use anyhow::{Context, Result};
use nebo_spec::pptx_types::*;
use std::collections::HashMap;
use std::io::{Read, Seek};

/// Unpack a PPTX file into a PptxSpec.
pub fn unpack_pptx<R: Read + Seek>(
    reader: R,
    _assets_dir: Option<&std::path::Path>,
    _pretty: bool,
) -> Result<PptxSpec> {
    let mut archive = zip::ZipArchive::new(reader).context("failed to open PPTX as ZIP")?;
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

    // Count slides
    let mut slide_count = 0;
    loop {
        let path = format!("ppt/slides/slide{}.xml", slide_count + 1);
        if files.contains_key(&path) {
            slide_count += 1;
        } else {
            break;
        }
    }

    let mut slides = Vec::new();
    for i in 1..=slide_count {
        let path = format!("ppt/slides/slide{i}.xml");
        if let Some(xml) = files.get(&path) {
            let slide = parse_slide(xml);
            slides.push(slide);
        }
    }

    let metadata = parse_metadata(&files);

    Ok(PptxSpec {
        version: 1,
        metadata,
        theme: None,
        size: None,
        slides,
    })
}

fn parse_slide(xml: &str) -> Slide {
    // Extract text from all <a:t> elements
    let mut texts = Vec::new();
    let mut pos = 0;
    while let Some(start) = xml[pos..].find("<a:t>") {
        let abs = pos + start + 5;
        let end = match xml[abs..].find("</a:t>") {
            Some(e) => abs + e,
            None => break,
        };
        let text = xml_unescape(&xml[abs..end]);
        if !text.is_empty() {
            texts.push(text);
        }
        pos = end + 6;
    }

    // First text is usually the title
    let title = texts.first().cloned();
    let body_texts: Vec<_> = texts.into_iter().skip(1).collect();

    let body = if body_texts.is_empty() {
        vec![]
    } else {
        vec![SlideBlock::Bullets {
            bullets: body_texts,
        }]
    };

    Slide {
        layout: "content".to_string(),
        title,
        subtitle: None,
        body,
        shapes: vec![],
        left: None,
        right: None,
        background: None,
        transition: None,
        notes: None,
    }
}

fn parse_metadata(files: &HashMap<String, String>) -> Option<PptxMetadata> {
    let core = files.get("docProps/core.xml")?;
    let title = extract_element_text(core, "dc:title");
    let creator = extract_element_text(core, "dc:creator");

    if title.is_none() && creator.is_none() {
        return None;
    }

    Some(PptxMetadata {
        title,
        creator,
        subject: None,
        description: None,
    })
}

fn extract_element_text(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{tag}");
    let close = format!("</{tag}>");
    if let Some(start) = xml.find(&open) {
        let tag_end = xml[start..].find('>')?;
        let abs_tag_end = start + tag_end;
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

fn xml_unescape(s: &str) -> String {
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}
