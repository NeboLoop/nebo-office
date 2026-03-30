use anyhow::{Context, Result};
use nebo_spec::xlsx_types::*;
use std::collections::HashMap;
use std::io::{Read, Seek};

/// Unpack an XLSX file into an XlsxSpec.
pub fn unpack_xlsx<R: Read + Seek>(
    reader: R,
    _assets_dir: Option<&std::path::Path>,
    _pretty: bool,
) -> Result<XlsxSpec> {
    let mut archive = zip::ZipArchive::new(reader).context("failed to open XLSX as ZIP")?;
    let mut files: HashMap<String, String> = HashMap::new();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        if name.ends_with(".xml") || name.ends_with(".rels") {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            files.insert(name, content);
        }
    }

    // Parse workbook for sheet names
    let sheet_names = parse_workbook_sheets(&files);

    // Parse shared strings
    let shared_strings = parse_shared_strings(&files);

    // Parse sheets
    let mut sheets = Vec::new();
    for (i, name) in sheet_names.iter().enumerate() {
        let sheet_path = format!("xl/worksheets/sheet{}.xml", i + 1);
        if let Some(sheet_xml) = files.get(&sheet_path) {
            let sheet = parse_sheet(sheet_xml, name, &shared_strings);
            sheets.push(sheet);
        }
    }

    // Parse metadata
    let metadata = parse_metadata(&files);

    Ok(XlsxSpec {
        version: 1,
        metadata,
        styles: None,
        sheets,
        named_ranges: vec![],
    })
}

fn parse_workbook_sheets(files: &HashMap<String, String>) -> Vec<String> {
    let mut names = Vec::new();
    if let Some(wb) = files.get("xl/workbook.xml") {
        let mut pos = 0;
        while let Some(start) = wb[pos..].find("<sheet ") {
            let abs = pos + start;
            let end = wb[abs..].find("/>").unwrap_or(wb.len() - abs) + abs;
            let chunk = &wb[abs..end + 2];
            if let Some(name) = extract_attr(chunk, "name") {
                names.push(xml_unescape(&name));
            }
            pos = end + 2;
        }
    }
    names
}

fn parse_shared_strings(files: &HashMap<String, String>) -> Vec<String> {
    let mut strings = Vec::new();
    if let Some(sst) = files.get("xl/sharedStrings.xml") {
        let mut pos = 0;
        while let Some(start) = sst[pos..].find("<si>") {
            let abs = pos + start;
            let end = match sst[abs..].find("</si>") {
                Some(e) => abs + e,
                None => break,
            };
            let si = &sst[abs..end + 5];

            // Extract text from <t> elements (handle <t> and <t ...>)
            let mut text = String::new();
            let mut tpos = 0;
            while let Some(tstart) = si[tpos..].find("<t") {
                let tabs = tpos + tstart;
                // Find the end of opening tag
                let tag_end = match si[tabs..].find('>') {
                    Some(e) => tabs + e,
                    None => break,
                };
                // Self-closing?
                if si[tag_end - 1..tag_end] == *"/" {
                    tpos = tag_end + 1;
                    continue;
                }
                let close = match si[tag_end..].find("</t>") {
                    Some(e) => tag_end + e,
                    None => break,
                };
                text.push_str(&si[tag_end + 1..close]);
                tpos = close + 4;
            }
            strings.push(xml_unescape(&text));
            pos = end + 5;
        }
    }
    strings
}

fn parse_sheet(xml: &str, name: &str, shared_strings: &[String]) -> Sheet {
    let mut rows = Vec::new();
    let mut merged = Vec::new();

    // Parse rows
    let mut pos = 0;
    while let Some(start) = xml[pos..].find("<row ") {
        let abs = pos + start;
        let row_end = match xml[abs..].find("</row>") {
            Some(e) => abs + e,
            None => break,
        };
        let row_xml = &xml[abs..row_end + 6];

        let mut cells: Vec<CellValue> = Vec::new();
        let mut cpos = 0;

        // Find all <c elements in this row
        while let Some(cstart) = row_xml[cpos..].find("<c ") {
            let cabs = cpos + cstart;

            // Determine cell end (self-closing or with content)
            let (cell_xml, cell_end) = if let Some(close) = row_xml[cabs..].find("</c>") {
                let e = cabs + close + 4;
                (&row_xml[cabs..e], e)
            } else if let Some(sc) = row_xml[cabs..].find("/>") {
                let e = cabs + sc + 2;
                (&row_xml[cabs..e], e)
            } else {
                break;
            };

            // Get cell reference to determine column position
            let cell_ref = extract_attr(cell_xml, "r").unwrap_or_default();
            let col_idx = cell_ref_to_col(&cell_ref);

            // Pad cells with empty values
            while cells.len() < col_idx {
                cells.push(CellValue::String(String::new()));
            }

            let cell_type = extract_attr(cell_xml, "t");
            let value = extract_element_text(cell_xml, "v");
            let formula = extract_element_text(cell_xml, "f");

            let cell_value = if formula.is_some() {
                CellValue::Rich(RichCell {
                    value: None,
                    formula,
                    format: None,
                    bold: None,
                    italic: None,
                    underline: None,
                    font: None,
                    size: None,
                    color: None,
                    shading: None,
                    align: None,
                    valign: None,
                    wrap: None,
                    colspan: None,
                })
            } else if let Some(ref v) = value {
                match cell_type.as_deref() {
                    Some("s") => {
                        // Shared string index
                        let idx: usize = v.parse().unwrap_or(0);
                        let s = shared_strings.get(idx).cloned().unwrap_or_default();
                        CellValue::String(s)
                    }
                    Some("b") => CellValue::Bool(v == "1"),
                    _ => {
                        // Try to parse as number
                        if let Ok(n) = v.parse::<f64>() {
                            CellValue::Number(n)
                        } else {
                            CellValue::String(v.clone())
                        }
                    }
                }
            } else {
                CellValue::String(String::new())
            };

            cells.push(cell_value);
            cpos = cell_end;
        }

        if !cells.is_empty() {
            rows.push(Row {
                cells,
                bold: None,
                italic: None,
                shading: None,
                color: None,
                font: None,
                size: None,
                height: None,
            });
        }

        pos = row_end + 6;
    }

    // Parse merged cells
    let mut mpos = 0;
    while let Some(start) = xml[mpos..].find("<mergeCell ") {
        let abs = mpos + start;
        let end = xml[abs..].find("/>").unwrap_or(xml.len() - abs) + abs;
        let chunk = &xml[abs..end + 2];
        if let Some(r) = extract_attr(chunk, "ref") {
            merged.push(r);
        }
        mpos = end + 2;
    }

    Sheet {
        name: name.to_string(),
        columns: vec![],
        rows,
        merged,
        freeze: None,
        conditional: vec![],
        charts: vec![],
        images: vec![],
        autofilter: None,
        validations: vec![],
        print: None,
    }
}

fn parse_metadata(files: &HashMap<String, String>) -> Option<XlsxMetadata> {
    let core = files.get("docProps/core.xml")?;
    let title = extract_element_text(core, "dc:title");
    let creator = extract_element_text(core, "dc:creator");
    let subject = extract_element_text(core, "dc:subject");
    let description = extract_element_text(core, "dc:description");

    if title.is_none() && creator.is_none() && subject.is_none() && description.is_none() {
        return None;
    }

    Some(XlsxMetadata {
        title,
        creator,
        subject,
        description,
    })
}

// --- Utilities ---

fn extract_attr(tag: &str, attr: &str) -> Option<String> {
    let pattern = format!("{attr}=\"");
    if let Some(start) = tag.find(&pattern) {
        let val_start = start + pattern.len();
        if let Some(end) = tag[val_start..].find('"') {
            return Some(tag[val_start..val_start + end].to_string());
        }
    }
    None
}

fn extract_element_text(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{tag}");
    let close = format!("</{tag}>");
    if let Some(start) = xml.find(&open) {
        let tag_end = xml[start..].find('>')?;
        let abs_tag_end = start + tag_end;
        // Self-closing check
        if xml[abs_tag_end - 1..abs_tag_end] == *"/" {
            return None;
        }
        let close_pos = xml[abs_tag_end..].find(&close)?;
        let text = &xml[abs_tag_end + 1..abs_tag_end + close_pos];
        Some(xml_unescape(text))
    } else {
        None
    }
}

fn cell_ref_to_col(cell_ref: &str) -> usize {
    let mut col = 0usize;
    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as usize - 'A' as usize + 1);
        } else {
            break;
        }
    }
    if col > 0 { col - 1 } else { 0 }
}

fn xml_unescape(s: &str) -> String {
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}
