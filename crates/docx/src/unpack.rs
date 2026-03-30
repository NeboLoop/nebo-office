use anyhow::{Context, Result, bail};
use nebo_office_core::units::*;
use nebo_spec::*;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::Path;

use crate::inline::runs_to_markdown;

pub fn unpack_docx<R: Read + Seek>(
    reader: R,
    assets_dir: Option<&Path>,
    _pretty: bool,
) -> Result<DocSpec> {
    let mut archive = zip::ZipArchive::new(reader).context("failed to open DOCX as ZIP")?;

    // Read all files into memory
    let mut files: HashMap<String, Vec<u8>> = HashMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let mut content = Vec::new();
        file.read_to_end(&mut content)?;
        files.insert(name, content);
    }

    // Parse relationships
    let doc_rels = if let Some(data) = files.get("word/_rels/document.xml.rels") {
        parse_relationships(data)?
    } else {
        HashMap::new()
    };

    // Parse document.xml
    let doc_data = files
        .get("word/document.xml")
        .context("missing word/document.xml")?;
    let doc_xml = String::from_utf8_lossy(doc_data);

    // Parse styles
    let styles = if let Some(data) = files.get("word/styles.xml") {
        let xml = String::from_utf8_lossy(data);
        Some(parse_styles(&xml)?)
    } else {
        None
    };

    // Parse page setup from document
    let page = parse_page_setup(&doc_xml)?;

    // Parse body blocks
    let body = parse_body(&doc_xml, &doc_rels, &files, assets_dir)?;

    // Parse headers/footers
    let headers = parse_header_footer_set(&doc_xml, &doc_rels, &files, true)?;
    let footers = parse_header_footer_set(&doc_xml, &doc_rels, &files, false)?;

    // Parse footnotes
    let footnotes = if let Some(data) = files.get("word/footnotes.xml") {
        let xml = String::from_utf8_lossy(data);
        parse_footnotes(&xml)?
    } else {
        None
    };

    // Parse comments
    let comments = if let Some(data) = files.get("word/comments.xml") {
        let xml = String::from_utf8_lossy(data);
        parse_comments(&xml)?
    } else {
        None
    };

    // Parse metadata
    let metadata = if let Some(data) = files.get("docProps/core.xml") {
        let xml = String::from_utf8_lossy(data);
        parse_metadata(&xml)?
    } else {
        None
    };

    // Extract images to assets dir
    if let Some(dir) = assets_dir {
        std::fs::create_dir_all(dir)?;
        for (path, data) in &files {
            if path.starts_with("word/media/") {
                let filename = path.strip_prefix("word/media/").unwrap();
                let dest = dir.join(filename);
                std::fs::write(&dest, data)?;
            }
        }
    }

    Ok(DocSpec {
        version: 1,
        page,
        styles,
        headers,
        footers,
        footnotes,
        comments,
        metadata,
        body,
    })
}

fn parse_relationships(data: &[u8]) -> Result<HashMap<String, (String, String, Option<String>)>> {
    let xml = String::from_utf8_lossy(data);
    let mut rels = HashMap::new();
    let mut reader = Reader::from_str(&xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.name().as_ref() == b"Relationship" => {
                let mut id = String::new();
                let mut rel_type = String::new();
                let mut target = String::new();
                let mut target_mode = None;

                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = String::from_utf8_lossy(&attr.value).to_string(),
                        b"Type" => rel_type = String::from_utf8_lossy(&attr.value).to_string(),
                        b"Target" => target = String::from_utf8_lossy(&attr.value).to_string(),
                        b"TargetMode" => {
                            target_mode = Some(String::from_utf8_lossy(&attr.value).to_string())
                        }
                        _ => {}
                    }
                }

                rels.insert(id, (rel_type, target, target_mode));
            }
            Ok(Event::Eof) => break,
            Err(e) => bail!("error parsing relationships: {e}"),
            _ => {}
        }
    }

    Ok(rels)
}

fn parse_page_setup(xml: &str) -> Result<Option<PageSetup>> {
    let mut reader = Reader::from_str(xml);
    let mut in_sect_pr = false;
    let mut page_setup = PageSetup {
        size: None,
        orientation: None,
        margin: None,
    };
    let mut found = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if local_name(e) == "sectPr" => {
                in_sect_pr = true;
            }
            Ok(Event::End(ref e)) if local_name_end(e) == "sectPr" => {
                in_sect_pr = false;
            }
            Ok(Event::Empty(ref e)) if in_sect_pr && local_name(e) == "pgSz" => {
                found = true;
                let mut w = 12240i64;
                let mut h = 15840i64;
                let mut orient = None;

                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    match key.split(':').last().unwrap_or(&key) {
                        "w" => w = val.parse().unwrap_or(12240),
                        "h" => h = val.parse().unwrap_or(15840),
                        "orient" => orient = Some(val),
                        _ => {}
                    }
                }

                // Determine page size name
                let (check_w, check_h) = if orient.as_deref() == Some("landscape") {
                    (h, w)
                } else {
                    (w, h)
                };
                page_setup.size = Some(match (check_w, check_h) {
                    (12240, 15840) => PageSize::Named("letter".into()),
                    (12240, 20160) => PageSize::Named("legal".into()),
                    (11906, 16838) => PageSize::Named("a4".into()),
                    _ => PageSize::Custom {
                        width: dxa_to_inches(check_w),
                        height: dxa_to_inches(check_h),
                    },
                });

                if let Some(o) = orient {
                    page_setup.orientation = Some(if o == "landscape" {
                        Orientation::Landscape
                    } else {
                        Orientation::Portrait
                    });
                }
            }
            Ok(Event::Empty(ref e)) if in_sect_pr && local_name(e) == "pgMar" => {
                found = true;
                let mut top = 1440i64;
                let mut bottom = 1440i64;
                let mut left = 1800i64;
                let mut right = 1800i64;

                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    match key.split(':').last().unwrap_or(&key) {
                        "top" => top = val.parse().unwrap_or(1440),
                        "bottom" => bottom = val.parse().unwrap_or(1440),
                        "left" => left = val.parse().unwrap_or(1800),
                        "right" => right = val.parse().unwrap_or(1800),
                        _ => {}
                    }
                }

                let t = dxa_to_inches(top);
                let b = dxa_to_inches(bottom);
                let l = dxa_to_inches(left);
                let r = dxa_to_inches(right);

                if (t - b).abs() < 0.01 && (l - r).abs() < 0.01 && (t - l).abs() < 0.01 {
                    page_setup.margin = Some(Margin::Uniform(t));
                } else {
                    page_setup.margin = Some(Margin::Custom {
                        top: Some(t),
                        bottom: Some(b),
                        left: Some(l),
                        right: Some(r),
                    });
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    Ok(if found { Some(page_setup) } else { None })
}

fn parse_body(
    xml: &str,
    rels: &HashMap<String, (String, String, Option<String>)>,
    files: &HashMap<String, Vec<u8>>,
    assets_dir: Option<&Path>,
) -> Result<Vec<Block>> {
    let mut blocks = Vec::new();

    let body_start = xml.find("<w:body>");
    let body_end = xml.rfind("</w:body>");

    if let (Some(start), Some(end)) = (body_start, body_end) {
        let body_xml = &xml[start + 8..end];
        blocks = parse_blocks(body_xml, rels, files, assets_dir)?;
    }

    Ok(blocks)
}

fn parse_blocks(
    xml: &str,
    rels: &HashMap<String, (String, String, Option<String>)>,
    _files: &HashMap<String, Vec<u8>>,
    _assets_dir: Option<&Path>,
) -> Result<Vec<Block>> {
    let mut blocks = Vec::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e);
                match name.as_str() {
                    "p" => {
                        let p_xml = read_element_xml(&mut reader, "p")?;
                        let full_xml = format!("<w:p {}>{}",
                            attrs_str(e),
                            p_xml
                        );
                        if let Some(block) = parse_paragraph_element(&full_xml, rels)? {
                            blocks.push(block);
                        }
                    }
                    "tbl" => {
                        let tbl_xml = read_element_xml(&mut reader, "tbl")?;
                        if let Some(block) = parse_table_element(&tbl_xml, rels)? {
                            blocks.push(block);
                        }
                    }
                    "sectPr" => {
                        // Skip final section properties
                        let _ = read_element_xml(&mut reader, "sectPr")?;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => bail!("XML parse error in blocks: {e}"),
            _ => {}
        }
    }

    Ok(blocks)
}

fn parse_paragraph_element(
    xml: &str,
    rels: &HashMap<String, (String, String, Option<String>)>,
) -> Result<Option<Block>> {
    // Check for heading style
    if let Some(heading_level) = extract_heading_level(xml) {
        let text = extract_text_content(xml);
        let id = extract_bookmark_name(xml);
        return Ok(Some(Block::Heading {
            heading: heading_level,
            text: TextContent::Simple(text),
            id,
        }));
    }

    // Check for page break
    if xml.contains("w:type=\"page\"") && xml.contains("w:br") {
        return Ok(Some(Block::PageBreak { page_break: true }));
    }

    // Check for list paragraph
    if xml.contains("<w:numPr>") || xml.contains("w:numPr") {
        // We'll handle lists by collecting consecutive list items later
        // For now, return as a paragraph
    }

    // Extract runs and merge adjacent runs with identical formatting
    let raw_runs = extract_runs(xml, rels)?;
    let runs = merge_adjacent_runs(raw_runs);
    let text = extract_text_content(xml);
    let align = extract_alignment(xml);

    if text.is_empty() && runs.is_empty() {
        return Ok(None);
    }

    // Try to simplify to markdown
    if align.is_none() && !runs.is_empty() {
        if let Some(md) = runs_to_markdown(&runs) {
            return Ok(Some(Block::Paragraph {
                paragraph: ParagraphContent::Simple(md),
            }));
        }
    }

    if runs.is_empty() {
        Ok(Some(Block::Paragraph {
            paragraph: ParagraphContent::Simple(text),
        }))
    } else if align.is_some() {
        // Check if we can use simple markdown
        if let Some(md) = runs_to_markdown(&runs) {
            Ok(Some(Block::Paragraph {
                paragraph: ParagraphContent::Full(ParagraphFull {
                    text: Some(md),
                    runs: None,
                    align,
                    spacing: None,
                    indent: None,
                    style: None,
                    id: None,
                    inserted: None,
                    deleted: None,
                }),
            }))
        } else {
            Ok(Some(Block::Paragraph {
                paragraph: ParagraphContent::Full(ParagraphFull {
                    text: None,
                    runs: Some(runs),
                    align,
                    spacing: None,
                    indent: None,
                    style: None,
                    id: None,
                    inserted: None,
                    deleted: None,
                }),
            }))
        }
    } else {
        Ok(Some(Block::Paragraph {
            paragraph: ParagraphContent::Full(ParagraphFull {
                text: None,
                runs: Some(runs),
                align: None,
                spacing: None,
                indent: None,
                style: None,
                id: None,
                inserted: None,
                deleted: None,
            }),
        }))
    }
}

fn parse_table_element(
    xml: &str,
    _rels: &HashMap<String, (String, String, Option<String>)>,
) -> Result<Option<Block>> {
    // Simple extraction: find all rows and cells
    let mut rows: Vec<Vec<String>> = Vec::new();
    let full_xml = format!("<w:tbl>{xml}</w:tbl>");
    let mut reader = Reader::from_str(&full_xml);
    let mut in_cell = false;
    let mut current_row: Vec<String> = Vec::new();
    let mut cell_text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e);
                match name.as_str() {
                    "tr" => {
                        current_row = Vec::new();
                    }
                    "tc" => {
                        in_cell = true;
                        cell_text = String::new();
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name_end(e);
                match name.as_str() {
                    "tr" => {
                        rows.push(std::mem::take(&mut current_row));
                    }
                    "tc" => {
                        in_cell = false;
                        current_row.push(std::mem::take(&mut cell_text));
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_cell => {
                cell_text.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    if rows.is_empty() {
        return Ok(None);
    }

    Ok(Some(Block::Table {
        table: TableContent::Simple(rows),
        header_rows: None,
    }))
}

fn extract_heading_level(xml: &str) -> Option<u8> {
    // Look for <w:pStyle w:val="Heading1"/> etc.
    let patterns = [
        ("Heading1", 1),
        ("Heading2", 2),
        ("Heading3", 3),
        ("Heading4", 4),
        ("Heading5", 5),
        ("Heading6", 6),
    ];
    for (pat, level) in &patterns {
        if xml.contains(&format!("w:val=\"{pat}\"")) {
            return Some(*level);
        }
    }
    None
}

fn extract_text_content(xml: &str) -> String {
    let mut text = String::new();
    let mut reader = Reader::from_str(xml);
    let mut in_t = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e);
                if name == "t" || name == "delText" {
                    in_t = true;
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name_end(e);
                if name == "t" || name == "delText" {
                    in_t = false;
                }
            }
            Ok(Event::Text(ref e)) if in_t => {
                text.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    text
}

fn extract_runs(
    xml: &str,
    rels: &HashMap<String, (String, String, Option<String>)>,
) -> Result<Vec<Run>> {
    let mut runs = Vec::new();
    let mut reader = Reader::from_str(xml);
    let mut in_r = false;
    let mut in_rpr = false;
    let mut in_hyperlink = false;
    let mut hyperlink_url = String::new();
    let mut current_text = String::new();
    let mut bold = false;
    let mut italic = false;
    let mut underline = false;
    let mut strike = false;
    let mut font: Option<String> = None;
    let mut size: Option<f64> = None;
    let mut color: Option<String> = None;
    let mut in_t = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e);
                match name.as_str() {
                    "r" => {
                        in_r = true;
                        current_text.clear();
                        bold = false;
                        italic = false;
                        underline = false;
                        strike = false;
                        font = None;
                        size = None;
                        color = None;
                    }
                    "rPr" if in_r => in_rpr = true,
                    "t" if in_r => in_t = true,
                    "hyperlink" => {
                        in_hyperlink = true;
                        // Get relationship ID
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("id") {
                                let rid =
                                    String::from_utf8_lossy(&attr.value).to_string();
                                if let Some((_, target, _)) = rels.get(&rid) {
                                    hyperlink_url = target.clone();
                                }
                            } else if key.ends_with("anchor") {
                                hyperlink_url = format!(
                                    "#{}",
                                    String::from_utf8_lossy(&attr.value)
                                );
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if in_rpr => {
                let name = local_name(e);
                match name.as_str() {
                    "b" => bold = true,
                    "i" => italic = true,
                    "u" => underline = true,
                    "strike" => strike = true,
                    "rFonts" => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("ascii") {
                                font = Some(
                                    String::from_utf8_lossy(&attr.value).to_string(),
                                );
                            }
                        }
                    }
                    "sz" => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("val") {
                                if let Ok(hp) = String::from_utf8_lossy(&attr.value).parse::<i64>()
                                {
                                    size = Some(half_points_to_points(hp));
                                }
                            }
                        }
                    }
                    "color" => {
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key.ends_with("val") {
                                let v = String::from_utf8_lossy(&attr.value).to_string();
                                if v != "auto" {
                                    color = Some(v);
                                }
                            }
                        }
                    }
                    "tab" if in_r && !in_rpr => {
                        runs.push(Run::Tab(TabRun { tab: true }));
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if in_r && !in_rpr => {
                let name = local_name(e);
                if name == "tab" {
                    runs.push(Run::Tab(TabRun { tab: true }));
                } else if name == "br" {
                    let mut break_type = "line".to_string();
                    for attr in e.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                        if key.ends_with("type") {
                            break_type = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }
                    runs.push(Run::Break(BreakRun {
                        break_type,
                    }));
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name_end(e);
                match name.as_str() {
                    "r" => {
                        in_r = false;
                        if !current_text.is_empty() {
                            let link = if in_hyperlink {
                                Some(hyperlink_url.clone())
                            } else {
                                None
                            };
                            runs.push(Run::Text(TextRun {
                                text: std::mem::take(&mut current_text),
                                bold: if bold { Some(true) } else { None },
                                italic: if italic { Some(true) } else { None },
                                underline: if underline { Some(true) } else { None },
                                strike: if strike { Some(true) } else { None },
                                superscript: None,
                                subscript: None,
                                font: font.take(),
                                size,
                                color: color.take(),
                                highlight: None,
                                link,
                                all_caps: None,
                                small_caps: None,
                            }));
                        }
                    }
                    "rPr" => in_rpr = false,
                    "t" => in_t = false,
                    "hyperlink" => {
                        in_hyperlink = false;
                        hyperlink_url.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_t => {
                current_text.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    Ok(runs)
}

fn extract_alignment(xml: &str) -> Option<String> {
    // Look for <w:jc w:val="center"/>
    if let Some(pos) = xml.find("w:jc") {
        if let Some(val_start) = xml[pos..].find("w:val=\"") {
            let start = pos + val_start + 7;
            if let Some(end) = xml[start..].find('"') {
                let val = &xml[start..start + end];
                return Some(
                    match val {
                        "both" => "justify",
                        other => other,
                    }
                    .to_string(),
                );
            }
        }
    }
    None
}

fn extract_bookmark_name(xml: &str) -> Option<String> {
    if let Some(pos) = xml.find("w:bookmarkStart") {
        if let Some(name_start) = xml[pos..].find("w:name=\"") {
            let start = pos + name_start + 8;
            if let Some(end) = xml[start..].find('"') {
                let name = &xml[start..start + end];
                if !name.starts_with('_') {
                    return Some(name.to_string());
                }
            }
        }
    }
    None
}

fn read_element_xml(reader: &mut Reader<&[u8]>, end_tag: &str) -> Result<String> {
    let mut depth = 1u32;
    let mut xml = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e);
                xml.push_str(&format!("<w:{}", name));
                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref());
                    let val = String::from_utf8_lossy(&attr.value);
                    xml.push_str(&format!(r#" {}="{}""#, key, val));
                }
                xml.push('>');
                if name == end_tag {
                    depth += 1;
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name_end(e);
                if name == end_tag {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                }
                xml.push_str(&format!("</w:{}>", name));
            }
            Ok(Event::Empty(ref e)) => {
                let name = local_name(e);
                xml.push_str(&format!("<w:{}", name));
                for attr in e.attributes().flatten() {
                    let key = String::from_utf8_lossy(attr.key.as_ref());
                    let val = String::from_utf8_lossy(&attr.value);
                    xml.push_str(&format!(r#" {}="{}""#, key, val));
                }
                xml.push_str("/>");
            }
            Ok(Event::Text(ref e)) => {
                xml.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::Eof) => break,
            Err(e) => bail!("XML error reading element: {e}"),
            _ => {}
        }
    }

    Ok(xml)
}

fn parse_styles(xml: &str) -> Result<Styles> {
    let mut styles = Styles {
        font: None,
        size: None,
        color: None,
        headings: None,
        custom: None,
    };

    // Extract default font from docDefaults
    if let Some(pos) = xml.find("w:rFonts") {
        if let Some(ascii_start) = xml[pos..].find("w:ascii=\"") {
            let start = pos + ascii_start + 9;
            if let Some(end) = xml[start..].find('"') {
                styles.font = Some(xml[start..start + end].to_string());
            }
        }
    }

    // Extract default size
    if xml.contains("rPrDefault") {
        if let Some(pos) = xml.find("w:sz") {
            if let Some(val_start) = xml[pos..].find("w:val=\"") {
                let start = pos + val_start + 7;
                if let Some(end) = xml[start..].find('"') {
                    if let Ok(hp) = xml[start..start + end].parse::<i64>() {
                        styles.size = Some(half_points_to_points(hp));
                    }
                }
            }
        }
    }

    Ok(styles)
}

fn parse_header_footer_set(
    doc_xml: &str,
    rels: &HashMap<String, (String, String, Option<String>)>,
    files: &HashMap<String, Vec<u8>>,
    is_header: bool,
) -> Result<Option<HeaderFooterSet>> {
    let ref_tag = if is_header {
        "headerReference"
    } else {
        "footerReference"
    };

    let mut set = HeaderFooterSet {
        default: None,
        first: None,
        even: None,
    };
    let mut found = false;

    // Find references in sectPr
    let mut search_pos = 0;
    while let Some(pos) = doc_xml[search_pos..].find(ref_tag) {
        let abs_pos = search_pos + pos;
        search_pos = abs_pos + ref_tag.len();

        // Extract type and r:id
        let chunk = &doc_xml[abs_pos..std::cmp::min(abs_pos + 200, doc_xml.len())];
        let hf_type = extract_attr_value(chunk, "w:type").unwrap_or_default();
        let rid = extract_attr_value(chunk, "r:id").unwrap_or_default();

        if let Some((_, target, _)) = rels.get(&rid) {
            let file_path = format!("word/{target}");
            if let Some(data) = files.get(&file_path) {
                let xml = String::from_utf8_lossy(data);
                let blocks = parse_hf_blocks(&xml, rels)?;
                found = true;
                match hf_type.as_str() {
                    "first" => set.first = Some(blocks),
                    "even" => set.even = Some(blocks),
                    _ => set.default = Some(blocks),
                }
            }
        }
    }

    Ok(if found { Some(set) } else { None })
}

fn parse_hf_blocks(
    xml: &str,
    rels: &HashMap<String, (String, String, Option<String>)>,
) -> Result<Vec<Block>> {
    // Simple: extract paragraphs from header/footer XML
    let mut blocks = Vec::new();
    let text = extract_text_content(xml);
    if !text.is_empty() {
        let runs = extract_runs(xml, rels)?;
        if let Some(md) = runs_to_markdown(&runs) {
            blocks.push(Block::Paragraph {
                paragraph: ParagraphContent::Simple(md),
            });
        } else if !runs.is_empty() {
            blocks.push(Block::Paragraph {
                paragraph: ParagraphContent::Full(ParagraphFull {
                    text: None,
                    runs: Some(runs),
                    align: extract_alignment(xml),
                    spacing: None,
                    indent: None,
                    style: None,
                    id: None,
                    inserted: None,
                    deleted: None,
                }),
            });
        } else {
            blocks.push(Block::Paragraph {
                paragraph: ParagraphContent::Simple(text),
            });
        }
    }
    Ok(blocks)
}

fn parse_footnotes(xml: &str) -> Result<Option<HashMap<String, String>>> {
    let mut footnotes = HashMap::new();
    let mut reader = Reader::from_str(xml);
    let mut in_footnote = false;
    let mut footnote_id = String::new();
    let mut footnote_text = String::new();
    let mut in_t = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e);
                if name == "footnote" {
                    in_footnote = true;
                    footnote_text.clear();
                    for attr in e.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                        if key.ends_with("id") {
                            footnote_id = String::from_utf8_lossy(&attr.value).to_string();
                        }
                        if key.ends_with("type") {
                            // Skip separator footnotes
                            in_footnote = false;
                        }
                    }
                } else if name == "t" && in_footnote {
                    in_t = true;
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name_end(e);
                if name == "footnote" && in_footnote {
                    in_footnote = false;
                    let text = footnote_text.trim().to_string();
                    if !text.is_empty() {
                        footnotes.insert(footnote_id.clone(), text);
                    }
                } else if name == "t" {
                    in_t = false;
                }
            }
            Ok(Event::Text(ref e)) if in_t && in_footnote => {
                footnote_text.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    Ok(if footnotes.is_empty() {
        None
    } else {
        Some(footnotes)
    })
}

fn parse_comments(xml: &str) -> Result<Option<HashMap<String, Comment>>> {
    let mut comments = HashMap::new();
    let mut reader = Reader::from_str(xml);
    let mut in_comment = false;
    let mut comment_id = String::new();
    let mut author = String::new();
    let mut date = String::new();
    let mut comment_text = String::new();
    let mut in_t = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e);
                if name == "comment" {
                    in_comment = true;
                    comment_text.clear();
                    for attr in e.attributes().flatten() {
                        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key.split(':').last().unwrap_or(&key) {
                            "id" => comment_id = val,
                            "author" => author = val,
                            "date" => date = val,
                            _ => {}
                        }
                    }
                } else if name == "t" && in_comment {
                    in_t = true;
                }
            }
            Ok(Event::End(ref e)) => {
                let name = local_name_end(e);
                if name == "comment" && in_comment {
                    in_comment = false;
                    comments.insert(
                        format!("c{comment_id}"),
                        Comment {
                            author: Some(author.clone()),
                            date: Some(date.clone()),
                            text: comment_text.trim().to_string(),
                            replies: None,
                        },
                    );
                } else if name == "t" {
                    in_t = false;
                }
            }
            Ok(Event::Text(ref e)) if in_t && in_comment => {
                comment_text.push_str(&e.unescape().unwrap_or_default());
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    Ok(if comments.is_empty() {
        None
    } else {
        Some(comments)
    })
}

fn parse_metadata(xml: &str) -> Result<Option<Metadata>> {
    let mut metadata = Metadata {
        title: None,
        subject: None,
        creator: None,
        description: None,
        keywords: None,
        category: None,
    };
    let mut found = false;

    let extract = |xml: &str, tag: &str| -> Option<String> {
        let open = format!("<dc:{tag}>");
        let close = format!("</dc:{tag}>");
        if let Some(start) = xml.find(&open) {
            let start = start + open.len();
            if let Some(end) = xml[start..].find(&close) {
                return Some(xml[start..start + end].to_string());
            }
        }
        // Try cp: namespace
        let open = format!("<cp:{tag}>");
        let close = format!("</cp:{tag}>");
        if let Some(start) = xml.find(&open) {
            let start = start + open.len();
            if let Some(end) = xml[start..].find(&close) {
                return Some(xml[start..start + end].to_string());
            }
        }
        None
    };

    if let Some(v) = extract(xml, "title") {
        metadata.title = Some(v);
        found = true;
    }
    if let Some(v) = extract(xml, "subject") {
        metadata.subject = Some(v);
        found = true;
    }
    if let Some(v) = extract(xml, "creator") {
        metadata.creator = Some(v);
        found = true;
    }
    if let Some(v) = extract(xml, "description") {
        metadata.description = Some(v);
        found = true;
    }
    if let Some(v) = extract(xml, "keywords") {
        metadata.keywords = Some(v.split(", ").map(|s| s.to_string()).collect());
        found = true;
    }
    if let Some(v) = extract(xml, "category") {
        metadata.category = Some(v);
        found = true;
    }

    Ok(if found { Some(metadata) } else { None })
}

fn local_name(e: &BytesStart) -> String {
    local_name_from_bytes(e.name().as_ref())
}

fn local_name_end(e: &quick_xml::events::BytesEnd) -> String {
    local_name_from_bytes(e.name().as_ref())
}

fn local_name_from_bytes(name: &[u8]) -> String {
    let full = String::from_utf8_lossy(name).to_string();
    full.split(':').last().unwrap_or(&full).to_string()
}

fn attrs_str(e: &BytesStart) -> String {
    let mut s = String::new();
    for attr in e.attributes().flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref());
        let val = String::from_utf8_lossy(&attr.value);
        s.push_str(&format!(r#" {}="{}""#, key, val));
    }
    s
}

fn extract_attr_value(xml: &str, attr_name: &str) -> Option<String> {
    let pattern = format!("{attr_name}=\"");
    if let Some(pos) = xml.find(&pattern) {
        let start = pos + pattern.len();
        if let Some(end) = xml[start..].find('"') {
            return Some(xml[start..start + end].to_string());
        }
    }
    None
}

/// Merge adjacent TextRuns that have identical formatting properties.
/// This mirrors the Python merge_runs.py logic: if two consecutive runs
/// only differ in text content (same bold/italic/underline/etc), merge them.
fn merge_adjacent_runs(runs: Vec<Run>) -> Vec<Run> {
    if runs.len() <= 1 {
        return runs;
    }

    let mut merged = Vec::with_capacity(runs.len());

    for run in runs {
        if let Run::Text(tr) = &run {
            if let Some(Run::Text(prev)) = merged.last_mut() {
                if text_run_props_match(prev, tr) {
                    prev.text.push_str(&tr.text);
                    continue;
                }
            }
        }
        merged.push(run);
    }

    merged
}

/// Check if two TextRuns have identical formatting (everything except text content).
fn text_run_props_match(a: &TextRun, b: &TextRun) -> bool {
    a.bold == b.bold
        && a.italic == b.italic
        && a.underline == b.underline
        && a.strike == b.strike
        && a.superscript == b.superscript
        && a.subscript == b.subscript
        && a.font == b.font
        && a.size == b.size
        && a.color == b.color
        && a.highlight == b.highlight
        && a.link == b.link
        && a.all_caps == b.all_caps
        && a.small_caps == b.small_caps
}
