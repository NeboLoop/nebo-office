use anyhow::Result;
use nebo_office_core::relationships::{
    build_content_types, RelationshipManager, OOXML_IMAGE_EXTENSIONS, REL_CORE_PROPS,
};
use nebo_office_core::zip_utils::create_zip;
use nebo_office_core::inches_to_emu;
use nebo_spec::pptx_types::*;
use std::collections::HashMap;
use std::io::{Seek, Write};

// PresentationML relationship types
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_SLIDE_LAYOUT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_SLIDE_MASTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_THEME: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_NOTES_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

// Content types
const CT_PRESENTATION: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const CT_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_SLIDE_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
const CT_SLIDE_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const CT_NOTES_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const CT_CORE_PROPS: &str =
    "application/vnd.openxmlformats-package.core-properties+xml";

// Namespaces
const NS_PRES: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_DC: &str = "http://purl.org/dc/elements/1.1/";
const NS_DCTERMS: &str = "http://purl.org/dc/terms/";
const NS_CP: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const NS_XSI: &str = "http://www.w3.org/2001/XMLSchema-instance";

/// Create a PPTX file from a spec.
pub fn create_pptx<W: Write + Seek>(
    spec: &PptxSpec,
    writer: W,
    assets_dir: Option<&std::path::Path>,
) -> Result<()> {
    let (slide_w, slide_h) = spec
        .size
        .as_ref()
        .map(|s| s.dimensions())
        .unwrap_or((10.0, 5.625));

    let cx = inches_to_emu(slide_w);
    let cy = inches_to_emu(slide_h);

    let theme_font = spec
        .theme
        .as_ref()
        .and_then(|t| t.font.clone())
        .unwrap_or_else(|| "Calibri".to_string());

    let mut content_types: HashMap<String, String> = HashMap::new();
    let mut files_data: Vec<(String, String)> = Vec::new();
    let mut image_files: Vec<(String, Vec<u8>)> = Vec::new();
    let mut image_counter = 0u32;

    // Build slide master
    let master_xml = build_slide_master(cx, cy, &theme_font, spec);
    files_data.push(("ppt/slideMasters/slideMaster1.xml".to_string(), master_xml));
    content_types.insert(
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
        CT_SLIDE_MASTER.to_string(),
    );

    // Build slide layouts (7 built-in layouts)
    let layout_names = [
        "title", "content", "section", "two-column", "blank", "title-only", "comparison",
    ];
    for (i, layout_name) in layout_names.iter().enumerate() {
        let layout_xml = build_slide_layout(cx, cy, layout_name);
        files_data.push((
            format!("ppt/slideLayouts/slideLayout{}.xml", i + 1),
            layout_xml,
        ));
        content_types.insert(
            format!("/ppt/slideLayouts/slideLayout{}.xml", i + 1),
            CT_SLIDE_LAYOUT.to_string(),
        );
    }

    // Build slides
    let mut slide_rels_files: Vec<(String, String)> = Vec::new();
    for (i, slide) in spec.slides.iter().enumerate() {
        let slide_num = i + 1;
        let layout_idx = layout_names
            .iter()
            .position(|&n| n == slide.layout)
            .unwrap_or(4) // default to blank
            + 1;

        let mut slide_rels = RelationshipManager::new();
        slide_rels.add(
            REL_SLIDE_LAYOUT,
            &format!("../slideLayouts/slideLayout{layout_idx}.xml"),
        );

        // Collect images from shapes
        let mut slide_images: Vec<(String, String)> = Vec::new();
        for shape in &slide.shapes {
            if let Some(ref img_file) = shape.image {
                if shape.shape_type == "image" {
                    image_counter += 1;
                    let ext = img_file.rsplit('.').next().unwrap_or("png");
                    let media_name = format!("image{image_counter}.{ext}");
                    let rel_id = slide_rels.add(
                        REL_IMAGE,
                        &format!("../media/{media_name}"),
                    );
                    slide_images.push((rel_id, img_file.clone()));

                    // Load the image
                    if let Some(ref dir) = assets_dir {
                        let path = dir.join(img_file);
                        if let Ok(data) = std::fs::read(&path) {
                            image_files.push((format!("ppt/media/{media_name}"), data));
                        }
                    }
                }
            }
        }

        // Notes
        if slide.notes.is_some() {
            let notes_rel_id = slide_rels.add(
                REL_NOTES_SLIDE,
                &format!("../notesSlides/notesSlide{slide_num}.xml"),
            );
            let notes_xml = build_notes_slide(slide, slide_num);
            files_data.push((
                format!("ppt/notesSlides/notesSlide{slide_num}.xml"),
                notes_xml,
            ));
            content_types.insert(
                format!("/ppt/notesSlides/notesSlide{slide_num}.xml"),
                CT_NOTES_SLIDE.to_string(),
            );
        }

        let slide_xml = build_slide(slide, cx, cy, &theme_font, &slide_images, spec);
        files_data.push((format!("ppt/slides/slide{slide_num}.xml"), slide_xml));
        content_types.insert(
            format!("/ppt/slides/slide{slide_num}.xml"),
            CT_SLIDE.to_string(),
        );

        slide_rels_files.push((
            format!("ppt/slides/_rels/slide{slide_num}.xml.rels"),
            slide_rels.to_xml(),
        ));
    }

    // Slide master rels
    let mut master_rels = RelationshipManager::new();
    for i in 0..layout_names.len() {
        master_rels.add(
            REL_SLIDE_LAYOUT,
            &format!("../slideLayouts/slideLayout{}.xml", i + 1),
        );
    }
    master_rels.add(REL_THEME, "../theme/theme1.xml");

    // Layout rels (each layout references the master)
    let mut layout_rels_files: Vec<(String, String)> = Vec::new();
    for i in 0..layout_names.len() {
        let mut layout_rels = RelationshipManager::new();
        layout_rels.add(REL_SLIDE_MASTER, "../slideMasters/slideMaster1.xml");
        layout_rels_files.push((
            format!("ppt/slideLayouts/_rels/slideLayout{}.xml.rels", i + 1),
            layout_rels.to_xml(),
        ));
    }

    // Theme
    let theme_xml = build_theme(spec);
    files_data.push(("ppt/theme/theme1.xml".to_string(), theme_xml));
    content_types.insert("/ppt/theme/theme1.xml".to_string(), CT_THEME.to_string());

    // Presentation
    let mut pres_rels = RelationshipManager::new();
    for i in 0..spec.slides.len() {
        pres_rels.add(REL_SLIDE, &format!("slides/slide{}.xml", i + 1));
    }
    pres_rels.add(REL_SLIDE_MASTER, "slideMasters/slideMaster1.xml");
    pres_rels.add(REL_THEME, "theme/theme1.xml");

    let pres_xml = build_presentation(spec, cx, cy);
    files_data.push(("ppt/presentation.xml".to_string(), pres_xml));
    content_types.insert(
        "/ppt/presentation.xml".to_string(),
        CT_PRESENTATION.to_string(),
    );

    // Package rels
    let mut pkg_rels = RelationshipManager::new();
    pkg_rels.add(REL_OFFICE_DOCUMENT, "ppt/presentation.xml");

    let has_metadata = spec.metadata.is_some();
    if has_metadata {
        pkg_rels.add(REL_CORE_PROPS, "docProps/core.xml");
        content_types.insert(
            "/docProps/core.xml".to_string(),
            CT_CORE_PROPS.to_string(),
        );
    }

    let content_types_xml = build_content_types(&content_types, OOXML_IMAGE_EXTENSIONS);
    let core_xml = if has_metadata {
        Some(build_core_xml(spec))
    } else {
        None
    };

    // Assemble ZIP
    let mut zip_files: Vec<(String, Vec<u8>)> = Vec::new();
    zip_files.push(("[Content_Types].xml".to_string(), content_types_xml.into_bytes()));
    zip_files.push(("_rels/.rels".to_string(), pkg_rels.to_xml().into_bytes()));
    zip_files.push((
        "ppt/_rels/presentation.xml.rels".to_string(),
        pres_rels.to_xml().into_bytes(),
    ));
    zip_files.push((
        "ppt/slideMasters/_rels/slideMaster1.xml.rels".to_string(),
        master_rels.to_xml().into_bytes(),
    ));

    for (path, xml) in &files_data {
        zip_files.push((path.clone(), xml.as_bytes().to_vec()));
    }
    for (path, xml) in &slide_rels_files {
        zip_files.push((path.clone(), xml.as_bytes().to_vec()));
    }
    for (path, xml) in &layout_rels_files {
        zip_files.push((path.clone(), xml.as_bytes().to_vec()));
    }
    for (path, data) in &image_files {
        zip_files.push((path.clone(), data.clone()));
    }
    if let Some(ref core) = core_xml {
        zip_files.push(("docProps/core.xml".to_string(), core.as_bytes().to_vec()));
    }

    // Convert to slice-of-slices for create_zip
    let refs: Vec<(&str, &[u8])> = zip_files.iter().map(|(p, d)| (p.as_str(), d.as_slice())).collect();
    create_zip(writer, &refs)?;
    Ok(())
}

fn build_presentation(spec: &PptxSpec, cx: i64, cy: i64) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<p:presentation xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">"#
    ));
    xml.push_str(&format!(r#"<p:sldMasterIdLst><p:sldMasterId r:id="rId{}"/></p:sldMasterIdLst>"#, spec.slides.len() + 1));
    xml.push_str("<p:sldIdLst>");
    for i in 0..spec.slides.len() {
        xml.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + i,
            i + 1
        ));
    }
    xml.push_str("</p:sldIdLst>");
    xml.push_str(&format!(r#"<p:sldSz cx="{cx}" cy="{cy}" type="custom"/>"#));
    xml.push_str(&format!(r#"<p:notesSz cx="{cy}" cy="{cx}"/>"#));
    xml.push_str("</p:presentation>");
    xml
}

fn build_slide(
    slide: &Slide,
    cx: i64,
    cy: i64,
    font: &str,
    images: &[(String, String)],
    spec: &PptxSpec,
) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">"#
    ));
    xml.push_str("<p:cSld>");

    // Background
    if let Some(ref bg) = slide.background {
        write_background(&mut xml, bg);
    }

    xml.push_str("<p:spTree>");
    xml.push_str(r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#);
    xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#);

    let mut shape_id = 2u32;

    // Title
    if let Some(ref title) = slide.title {
        let title_color = get_text_color(slide, spec);
        write_text_shape(
            &mut xml,
            shape_id,
            "Title",
            title,
            inches_to_emu(0.5),
            inches_to_emu(0.3),
            cx - inches_to_emu(1.0),
            inches_to_emu(0.8),
            3200,
            true,
            font,
            &title_color,
        );
        shape_id += 1;
    }

    // Subtitle (title layout)
    if let Some(ref subtitle) = slide.subtitle {
        let text_color = get_text_color(slide, spec);
        write_text_shape(
            &mut xml,
            shape_id,
            "Subtitle",
            subtitle,
            inches_to_emu(0.5),
            inches_to_emu(1.2),
            cx - inches_to_emu(1.0),
            inches_to_emu(0.6),
            2000,
            false,
            font,
            &text_color,
        );
        shape_id += 1;
    }

    // Body blocks
    let body_top = if slide.title.is_some() {
        inches_to_emu(1.3)
    } else {
        inches_to_emu(0.5)
    };
    let body_width = match slide.layout.as_str() {
        "two-column" => (cx - inches_to_emu(1.5)) / 2,
        _ => cx - inches_to_emu(1.0),
    };

    let text_color = get_text_color(slide, spec);
    let mut y_offset = body_top;

    // Handle two-column layout
    if slide.layout == "two-column" {
        if let Some(ref left) = slide.left {
            let mut left_y = body_top;
            for block in left {
                shape_id = write_slide_block(
                    &mut xml,
                    block,
                    shape_id,
                    inches_to_emu(0.5),
                    left_y,
                    body_width,
                    font,
                    &text_color,
                );
                left_y += inches_to_emu(0.5);
            }
        }
        if let Some(ref right) = slide.right {
            let right_x = inches_to_emu(0.5) + body_width + inches_to_emu(0.5);
            let mut right_y = body_top;
            for block in right {
                shape_id = write_slide_block(
                    &mut xml,
                    block,
                    shape_id,
                    right_x,
                    right_y,
                    body_width,
                    font,
                    &text_color,
                );
                right_y += inches_to_emu(0.5);
            }
        }
    } else {
        for block in &slide.body {
            shape_id = write_slide_block(
                &mut xml,
                block,
                shape_id,
                inches_to_emu(0.5),
                y_offset,
                body_width,
                font,
                &text_color,
            );
            y_offset += inches_to_emu(0.5);
        }
    }

    // Custom shapes
    let mut img_idx = 0;
    for shape in &slide.shapes {
        let x = inches_to_emu(shape.x.unwrap_or(0.0));
        let y = inches_to_emu(shape.y.unwrap_or(0.0));
        let w = inches_to_emu(shape.w.unwrap_or(1.0));
        let h = inches_to_emu(shape.h.unwrap_or(1.0));

        match shape.shape_type.as_str() {
            "rect" | "rounded-rect" | "oval" => {
                let preset = match shape.shape_type.as_str() {
                    "oval" => "ellipse",
                    "rounded-rect" => "roundRect",
                    _ => "rect",
                };
                write_geom_shape(
                    &mut xml,
                    shape_id,
                    &shape.shape_type,
                    preset,
                    x, y, w, h,
                    shape.fill.as_deref(),
                    shape.opacity,
                    shape.line_color.as_deref(),
                    shape.line_width,
                );
                shape_id += 1;

                // Text inside shape
                if let Some(ref text) = shape.text {
                    // The text was already written as part of the shape in a real implementation
                    // For now we add it as a separate text box
                }
            }
            "text" => {
                let color = shape.color.as_deref().unwrap_or("000000");
                let size = (shape.font_size.unwrap_or(18.0) * 100.0) as i64;
                let bold = shape.bold.unwrap_or(false);
                let text = shape.text.as_deref().unwrap_or("");
                write_text_shape(
                    &mut xml,
                    shape_id,
                    "TextBox",
                    text,
                    x, y, w, h,
                    size,
                    bold,
                    font,
                    color,
                );
                shape_id += 1;
            }
            "image" => {
                if img_idx < images.len() {
                    let (ref rel_id, _) = images[img_idx];
                    write_image_shape(&mut xml, shape_id, rel_id, x, y, w, h);
                    shape_id += 1;
                    img_idx += 1;
                }
            }
            "line" => {
                write_line_shape(
                    &mut xml,
                    shape_id,
                    x, y, w, h,
                    shape.line_color.as_deref().unwrap_or("000000"),
                    shape.line_width.unwrap_or(1.0),
                );
                shape_id += 1;
            }
            _ => {}
        }
    }

    xml.push_str("</p:spTree>");
    xml.push_str("</p:cSld>");

    // Transition
    if let Some(ref trans) = slide.transition {
        let dur = (trans.duration.unwrap_or(0.5) * 1000.0) as i64;
        let trans_type = match trans.transition_type.as_str() {
            "fade" => "fade",
            "push" => "push",
            "wipe" => "wipe",
            _ => "fade",
        };
        xml.push_str(&format!(
            r#"<p:transition spd="med" advClick="1"><p:{trans_type} dur="{dur}"/></p:transition>"#
        ));
    }

    xml.push_str("</p:sld>");
    xml
}

fn write_text_shape(
    xml: &mut String,
    id: u32,
    name: &str,
    text: &str,
    x: i64, y: i64, w: i64, h: i64,
    font_size: i64,
    bold: bool,
    font: &str,
    color: &str,
) {
    xml.push_str(&format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#
    ));
    xml.push_str(&format!(
        r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"#
    ));
    xml.push_str("<p:txBody>");
    xml.push_str(r#"<a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>"#);

    // Split text by newlines for multiple paragraphs
    for line in text.split('\n') {
        xml.push_str("<a:p>");
        let bold_attr = if bold { r#" b="1""# } else { "" };

        // Parse markdown bold
        let parts = parse_inline_bold(line);
        for (part_text, part_bold) in &parts {
            let is_bold = *part_bold || bold;
            let b = if is_bold { r#" b="1""# } else { "" };
            xml.push_str(&format!(
                r#"<a:r><a:rPr lang="en-US" sz="{font_size}"{b} dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{font}"/></a:rPr><a:t>{}</a:t></a:r>"#,
                xml_escape(part_text)
            ));
        }
        xml.push_str("</a:p>");
    }
    xml.push_str("</p:txBody></p:sp>");
}

fn write_slide_block(
    xml: &mut String,
    block: &SlideBlock,
    mut shape_id: u32,
    x: i64, y: i64, w: i64,
    font: &str,
    color: &str,
) -> u32 {
    match block {
        SlideBlock::Paragraph { paragraph } => {
            write_text_shape(xml, shape_id, "Content", paragraph, x, y, w, inches_to_emu(0.4), 1800, false, font, color);
            shape_id + 1
        }
        SlideBlock::Bullets { bullets } => {
            let text = bullets.join("\n");
            // Build bullet list as separate paragraphs with bullet markers
            xml.push_str(&format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="Content"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#
            ));
            let h = inches_to_emu(0.35 * bullets.len() as f64);
            xml.push_str(&format!(
                r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"#
            ));
            xml.push_str("<p:txBody>");
            xml.push_str(r#"<a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>"#);
            for bullet in bullets {
                xml.push_str("<a:p>");
                xml.push_str(r#"<a:pPr><a:buChar char="&#x2022;"/></a:pPr>"#);
                let parts = parse_inline_bold(bullet);
                for (part_text, part_bold) in &parts {
                    let b = if *part_bold { r#" b="1""# } else { "" };
                    xml.push_str(&format!(
                        r#"<a:r><a:rPr lang="en-US" sz="1800"{b} dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="{font}"/></a:rPr><a:t>{}</a:t></a:r>"#,
                        xml_escape(part_text)
                    ));
                }
                xml.push_str("</a:p>");
            }
            xml.push_str("</p:txBody></p:sp>");
            shape_id + 1
        }
        SlideBlock::Table { table, header_rows } => {
            if table.is_empty() {
                return shape_id;
            }
            let num_cols = table[0].len();
            let col_w = w / num_cols.max(1) as i64;
            let row_h = inches_to_emu(0.35);
            let h = row_h * table.len() as i64;

            xml.push_str("<p:graphicFrame>");
            xml.push_str(&format!(
                r#"<p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>"#
            ));
            xml.push_str(&format!(
                r#"<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>"#
            ));
            xml.push_str(r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">"#);
            xml.push_str("<a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"/><a:tblGrid>");
            for _ in 0..num_cols {
                xml.push_str(&format!(r#"<a:gridCol w="{col_w}"/>"#));
            }
            xml.push_str("</a:tblGrid>");

            let header_count = header_rows.unwrap_or(0) as usize;
            for (ri, row) in table.iter().enumerate() {
                xml.push_str(&format!(r#"<a:tr h="{row_h}">"#));
                for cell_text in row {
                    xml.push_str("<a:tc>");
                    xml.push_str("<a:txBody><a:bodyPr/><a:lstStyle/><a:p>");
                    let is_header = ri < header_count;
                    let b = if is_header { r#" b="1""# } else { "" };
                    xml.push_str(&format!(
                        r#"<a:r><a:rPr lang="en-US" sz="1400"{b}/><a:t>{}</a:t></a:r>"#,
                        xml_escape(cell_text)
                    ));
                    xml.push_str("</a:p></a:txBody>");
                    xml.push_str("<a:tcPr/>");
                    xml.push_str("</a:tc>");
                }
                xml.push_str("</a:tr>");
            }

            xml.push_str("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
            shape_id + 1
        }
        _ => shape_id,
    }
}

fn write_geom_shape(
    xml: &mut String,
    id: u32,
    name: &str,
    preset: &str,
    x: i64, y: i64, w: i64, h: i64,
    fill: Option<&str>,
    opacity: Option<f64>,
    line_color: Option<&str>,
    line_width: Option<f64>,
) {
    xml.push_str(&format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#
    ));
    xml.push_str(&format!(
        r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom>"#
    ));

    if let Some(color) = fill {
        if let Some(op) = opacity {
            let alpha = (op * 100000.0) as i64;
            xml.push_str(&format!(
                r#"<a:solidFill><a:srgbClr val="{color}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>"#
            ));
        } else {
            xml.push_str(&format!(
                r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#
            ));
        }
    } else {
        xml.push_str("<a:noFill/>");
    }

    if let Some(lc) = line_color {
        let lw = (line_width.unwrap_or(1.0) * 12700.0) as i64;
        xml.push_str(&format!(
            r#"<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{lc}"/></a:solidFill></a:ln>"#
        ));
    }

    xml.push_str("</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>");
}

fn write_image_shape(
    xml: &mut String,
    id: u32,
    rel_id: &str,
    x: i64, y: i64, w: i64, h: i64,
) {
    xml.push_str(&format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="Image"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>"#
    ));
    xml.push_str(&format!(
        r#"<p:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>"#
    ));
    xml.push_str(&format!(
        r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
    ));
}

fn write_line_shape(
    xml: &mut String,
    id: u32,
    x: i64, y: i64, w: i64, h: i64,
    color: &str,
    width: f64,
) {
    let lw = (width * 12700.0) as i64;
    xml.push_str(&format!(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{id}" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>"#
    ));
    xml.push_str(&format!(
        r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="{lw}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>"#
    ));
}

fn write_background(xml: &mut String, bg: &SlideBackground) {
    xml.push_str("<p:bg><p:bgPr>");
    match bg {
        SlideBackground::Solid { color } => {
            xml.push_str(&format!(
                r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#
            ));
        }
        SlideBackground::Gradient { gradient } => {
            let angle = (gradient.angle.unwrap_or(270.0) * 60000.0) as i64;
            xml.push_str(&format!(r#"<a:gradFill><a:gsLst>"#));
            xml.push_str(&format!(
                r#"<a:gs pos="0"><a:srgbClr val="{}"/></a:gs>"#,
                gradient.from
            ));
            xml.push_str(&format!(
                r#"<a:gs pos="100000"><a:srgbClr val="{}"/></a:gs>"#,
                gradient.to
            ));
            xml.push_str(&format!(
                r#"</a:gsLst><a:lin ang="{angle}" scaled="1"/></a:gradFill>"#
            ));
        }
        SlideBackground::Image { .. } => {
            // Image backgrounds need relationship handling — simplified for now
            xml.push_str("<a:noFill/>");
        }
    }
    xml.push_str("<a:effectLst/></p:bgPr></p:bg>");
}

fn build_slide_master(cx: i64, cy: i64, font: &str, spec: &PptxSpec) -> String {
    let bg_color = spec
        .theme
        .as_ref()
        .and_then(|t| t.colors.as_ref())
        .and_then(|c| c.background.as_deref())
        .unwrap_or("FFFFFF");

    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<p:sldMaster xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">"#
    ));
    xml.push_str("<p:cSld>");
    xml.push_str(&format!(r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#));
    xml.push_str("<p:spTree>");
    xml.push_str(r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#);
    xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#);
    xml.push_str("</p:spTree>");
    xml.push_str("</p:cSld>");

    // Slide layout references
    xml.push_str("<p:sldLayoutIdLst>");
    for i in 1..=7 {
        xml.push_str(&format!(
            r#"<p:sldLayoutId id="{}" r:id="rId{i}"/>"#,
            2147483648u64 + i
        ));
    }
    xml.push_str("</p:sldLayoutIdLst>");

    xml.push_str("</p:sldMaster>");
    xml
}

fn build_slide_layout(cx: i64, cy: i64, name: &str) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<p:sldLayout xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}" type="{}" preserve="1">"#,
        layout_type_name(name)
    ));
    xml.push_str("<p:cSld><p:spTree>");
    xml.push_str(r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#);
    xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#);
    xml.push_str("</p:spTree></p:cSld>");
    xml.push_str("</p:sldLayout>");
    xml
}

fn layout_type_name(name: &str) -> &str {
    match name {
        "title" => "title",
        "content" => "obj",
        "section" => "secHead",
        "two-column" => "twoObj",
        "blank" => "blank",
        "title-only" => "titleOnly",
        "comparison" => "twoTxTwoObj",
        _ => "blank",
    }
}

fn build_notes_slide(slide: &Slide, slide_num: usize) -> String {
    let notes_text = slide.notes.as_deref().unwrap_or("");
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<p:notes xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">"#
    ));
    xml.push_str("<p:cSld><p:spTree>");
    xml.push_str(r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#);
    xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#);

    // Notes text shape
    xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>"#);
    xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>"#);
    for line in notes_text.split('\n') {
        xml.push_str(&format!(
            r#"<a:p><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p>"#,
            xml_escape(line)
        ));
    }
    xml.push_str("</p:txBody></p:sp>");

    xml.push_str("</p:spTree></p:cSld></p:notes>");
    xml
}

fn build_theme(spec: &PptxSpec) -> String {
    let colors = spec.theme.as_ref().and_then(|t| t.colors.as_ref());
    let font = spec
        .theme
        .as_ref()
        .and_then(|t| t.font.as_deref())
        .unwrap_or("Calibri");

    let dk1 = colors.and_then(|c| c.text.as_deref()).unwrap_or("333333");
    let lt1 = colors
        .and_then(|c| c.background.as_deref())
        .unwrap_or("FFFFFF");
    let accent1 = colors
        .and_then(|c| c.accent1.as_deref())
        .unwrap_or("4472C4");
    let accent2 = colors
        .and_then(|c| c.accent2.as_deref())
        .unwrap_or("ED7D31");
    let dk2 = colors
        .and_then(|c| c.primary.as_deref())
        .unwrap_or("1F4E79");

    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">"#);
    xml.push_str(r#"<a:themeElements><a:clrScheme name="Custom">"#);
    xml.push_str(&format!(r#"<a:dk1><a:srgbClr val="{dk1}"/></a:dk1>"#));
    xml.push_str(&format!(r#"<a:lt1><a:srgbClr val="{lt1}"/></a:lt1>"#));
    xml.push_str(&format!(r#"<a:dk2><a:srgbClr val="{dk2}"/></a:dk2>"#));
    xml.push_str(r#"<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>"#);
    xml.push_str(&format!(r#"<a:accent1><a:srgbClr val="{accent1}"/></a:accent1>"#));
    xml.push_str(&format!(r#"<a:accent2><a:srgbClr val="{accent2}"/></a:accent2>"#));
    xml.push_str(r#"<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>"#);
    xml.push_str(r#"<a:accent4><a:srgbClr val="FFC000"/></a:accent4>"#);
    xml.push_str(r#"<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>"#);
    xml.push_str(r#"<a:accent6><a:srgbClr val="70AD47"/></a:accent6>"#);
    xml.push_str(r#"<a:hlink><a:srgbClr val="0563C1"/></a:hlink>"#);
    xml.push_str(r#"<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>"#);
    xml.push_str(r#"</a:clrScheme>"#);
    xml.push_str(&format!(r#"<a:fontScheme name="Custom"><a:majorFont><a:latin typeface="{font}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="{font}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>"#));
    xml.push_str(r#"<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>"#);
    xml.push_str("</a:themeElements></a:theme>");
    xml
}

fn build_core_xml(spec: &PptxSpec) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<cp:coreProperties xmlns:cp="{NS_CP}" xmlns:dc="{NS_DC}" xmlns:dcterms="{NS_DCTERMS}" xmlns:xsi="{NS_XSI}">"#
    ));

    if let Some(ref meta) = spec.metadata {
        if let Some(ref title) = meta.title {
            xml.push_str(&format!("<dc:title>{}</dc:title>", xml_escape(title)));
        }
        if let Some(ref creator) = meta.creator {
            xml.push_str(&format!("<dc:creator>{}</dc:creator>", xml_escape(creator)));
        }
    }

    let now = chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ");
    xml.push_str(&format!(r#"<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>"#));
    xml.push_str(&format!(r#"<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>"#));
    xml.push_str("<cp:revision>1</cp:revision>");
    xml.push_str("</cp:coreProperties>");
    xml
}

fn get_text_color(slide: &Slide, spec: &PptxSpec) -> String {
    // If slide has a dark background, use white text
    if let Some(ref bg) = slide.background {
        match bg {
            SlideBackground::Solid { color } => {
                if is_dark_color(color) {
                    return "FFFFFF".to_string();
                }
            }
            SlideBackground::Gradient { gradient } => {
                if is_dark_color(&gradient.from) {
                    return "FFFFFF".to_string();
                }
            }
            _ => {}
        }
    }
    spec.theme
        .as_ref()
        .and_then(|t| t.colors.as_ref())
        .and_then(|c| c.text.clone())
        .unwrap_or_else(|| "333333".to_string())
}

fn is_dark_color(hex: &str) -> bool {
    if hex.len() < 6 {
        return false;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);
    let luminance = 0.299 * r as f64 + 0.587 * g as f64 + 0.114 * b as f64;
    luminance < 128.0
}

/// Parse inline **bold** markdown in text, returning (text, is_bold) segments.
fn parse_inline_bold(text: &str) -> Vec<(String, bool)> {
    let mut result = Vec::new();
    let mut pos = 0;
    while let Some(start) = text[pos..].find("**") {
        let abs_start = pos + start;
        if abs_start > pos {
            result.push((text[pos..abs_start].to_string(), false));
        }
        let after = abs_start + 2;
        if let Some(end) = text[after..].find("**") {
            result.push((text[after..after + end].to_string(), true));
            pos = after + end + 2;
        } else {
            result.push((text[abs_start..].to_string(), false));
            pos = text.len();
        }
    }
    if pos < text.len() {
        result.push((text[pos..].to_string(), false));
    }
    if result.is_empty() {
        result.push((text.to_string(), false));
    }
    result
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_basic_pptx() {
        let spec = PptxSpec {
            version: 1,
            metadata: Some(PptxMetadata {
                title: Some("Test Deck".to_string()),
                creator: Some("Test".to_string()),
                subject: None,
                description: None,
            }),
            theme: None,
            size: Some(SlideSize::Named("16:9".to_string())),
            slides: vec![
                Slide {
                    layout: "title".to_string(),
                    title: Some("Welcome".to_string()),
                    subtitle: Some("A test presentation".to_string()),
                    body: vec![],
                    shapes: vec![],
                    left: None,
                    right: None,
                    background: Some(SlideBackground::Solid {
                        color: "1F4E79".to_string(),
                    }),
                    transition: None,
                    notes: Some("Speaker notes here".to_string()),
                },
                Slide {
                    layout: "content".to_string(),
                    title: Some("Key Points".to_string()),
                    subtitle: None,
                    body: vec![SlideBlock::Bullets {
                        bullets: vec![
                            "First point".to_string(),
                            "**Bold** second point".to_string(),
                        ],
                    }],
                    shapes: vec![],
                    left: None,
                    right: None,
                    background: None,
                    transition: Some(SlideTransition {
                        transition_type: "fade".to_string(),
                        duration: Some(0.5),
                    }),
                    notes: None,
                },
            ],
        };

        let mut buf = std::io::Cursor::new(Vec::new());
        create_pptx(&spec, &mut buf, None).unwrap();
        let data = buf.into_inner();
        assert!(data.len() > 100);

        let cursor = std::io::Cursor::new(&data);
        let archive = zip::ZipArchive::new(cursor).unwrap();
        let names: Vec<_> = archive.file_names().collect();
        assert!(names.contains(&"[Content_Types].xml"));
        assert!(names.contains(&"ppt/presentation.xml"));
        assert!(names.contains(&"ppt/slides/slide1.xml"));
        assert!(names.contains(&"ppt/slides/slide2.xml"));
        assert!(names.contains(&"ppt/theme/theme1.xml"));
    }

    #[test]
    fn test_parse_inline_bold() {
        let result = parse_inline_bold("Revenue: **$12.5M** (+15%)");
        assert_eq!(result.len(), 3);
        assert_eq!(result[0], ("Revenue: ".to_string(), false));
        assert_eq!(result[1], ("$12.5M".to_string(), true));
        assert_eq!(result[2], (" (+15%)".to_string(), false));
    }

    #[test]
    fn test_is_dark_color() {
        assert!(is_dark_color("1F4E79"));
        assert!(is_dark_color("000000"));
        assert!(!is_dark_color("FFFFFF"));
        assert!(!is_dark_color("E7E6E6"));
    }
}
