use anyhow::{Context, Result};
use nebo_office_core::units::*;
use nebo_spec::*;
use std::collections::HashMap;
use std::io::{Seek, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;

use crate::inline::parse_inline_text;
use crate::relationships::*;

const NS_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_DC: &str = "http://purl.org/dc/elements/1.1/";
const NS_CP: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const NS_DCTERMS: &str = "http://purl.org/dc/terms/";

struct DocxBuilder {
    doc_rels: RelationshipManager,
    content_types: HashMap<String, String>,
    images: Vec<(String, Vec<u8>)>,
    image_counter: u32,
    hyperlink_map: HashMap<String, String>,
    footnote_counter: u32,
    footnotes_xml: Vec<String>,
    comment_id_map: HashMap<String, u32>,
    comments_xml: Vec<String>,
    has_footnotes: bool,
    has_comments: bool,
    has_numbering: bool,
    has_headers: bool,
    has_footers: bool,
    header_parts: Vec<(String, String, String)>, // (rel_id, filename, xml)
    footer_parts: Vec<(String, String, String)>,
    num_id_counter: u32,          // next numId to allocate
    extra_num_entries: Vec<u32>,   // numIds that need lvlOverride restart
}

impl DocxBuilder {
    fn new() -> Self {
        Self {
            doc_rels: RelationshipManager::new(),
            content_types: HashMap::new(),
            images: Vec::new(),
            image_counter: 0,
            hyperlink_map: HashMap::new(),
            footnote_counter: 0,
            footnotes_xml: Vec::new(),
            comment_id_map: HashMap::new(),
            comments_xml: Vec::new(),
            has_footnotes: false,
            has_comments: false,
            has_numbering: false,
            has_headers: false,
            has_footers: false,
            header_parts: Vec::new(),
            footer_parts: Vec::new(),
            num_id_counter: 3,      // 1=bullets, 2=numbered, 3+ = restart overrides
            extra_num_entries: Vec::new(),
        }
    }

    fn add_image(&mut self, filename: &str, data: Vec<u8>) -> String {
        self.image_counter += 1;
        let ext = Path::new(filename)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("png");
        let internal_name = format!("image{}.{}", self.image_counter, ext);
        let target = format!("media/{internal_name}");
        let rel_id = self.doc_rels.add(REL_IMAGE, &target);
        self.images.push((target, data));
        rel_id
    }

    fn add_hyperlink(&mut self, url: &str) -> String {
        if let Some(id) = self.hyperlink_map.get(url) {
            return id.clone();
        }
        let rel_id = self.doc_rels.add_external(REL_HYPERLINK, url);
        self.hyperlink_map.insert(url.to_string(), rel_id.clone());
        rel_id
    }

    fn add_footnote(&mut self, _id: &str, text: &str) {
        self.has_footnotes = true;
        self.footnote_counter += 1;
        let fid = self.footnote_counter;
        self.footnotes_xml.push(format!(
            r#"<w:footnote w:id="{fid}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> {}</w:t></w:r></w:p></w:footnote>"#,
            escape_xml(text)
        ));
    }

    fn register_comments(&mut self, comments: &HashMap<String, Comment>) {
        self.has_comments = true;
        let mut cid: u32 = 0;
        for (key, comment) in comments {
            self.comment_id_map.insert(key.clone(), cid);
            self.comments_xml.push(format!(
                r#"<w:comment w:id="{cid}" w:author="{}" w:date="{}"><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:comment>"#,
                escape_xml(comment.author.as_deref().unwrap_or("Author")),
                escape_xml(comment.date.as_deref().unwrap_or("")),
                escape_xml(&comment.text),
            ));
            cid += 1;
        }
    }
}

pub fn create_docx<W: Write + Seek>(
    spec: &DocSpec,
    writer: W,
    assets_dir: Option<&Path>,
) -> Result<()> {
    let mut builder = DocxBuilder::new();

    // Register relationships for standard parts
    let _styles_id = builder.doc_rels.add(REL_STYLES, "styles.xml");
    let _settings_id = builder.doc_rels.add(REL_SETTINGS, "settings.xml");

    // Pre-process: register footnotes
    if let Some(footnotes) = &spec.footnotes {
        for (key, text) in footnotes {
            builder.add_footnote(key, text);
        }
        builder
            .doc_rels
            .add(REL_FOOTNOTES, "footnotes.xml");
    }

    // Pre-process: register comments
    if let Some(comments) = &spec.comments {
        builder.register_comments(comments);
        builder.doc_rels.add(REL_COMMENTS, "comments.xml");
    }

    // Check if we need numbering
    for block in &spec.body {
        match block {
            Block::Bullets { .. } | Block::Numbered { .. } => {
                builder.has_numbering = true;
                break;
            }
            _ => {}
        }
    }
    if builder.has_numbering {
        builder.doc_rels.add(REL_NUMBERING, "numbering.xml");
    }

    // Build headers/footers
    if let Some(headers) = &spec.headers {
        build_header_footer_parts(&mut builder, headers, true, &spec.styles, assets_dir)?;
    }
    if let Some(footers) = &spec.footers {
        build_header_footer_parts(&mut builder, footers, false, &spec.styles, assets_dir)?;
    }

    // Build document.xml
    let document_xml = build_document_xml(spec, &mut builder, assets_dir)?;

    // Build styles.xml
    let styles_xml = build_styles_xml(spec);

    // Build settings.xml
    let settings_xml = build_settings_xml();

    // Build [Content_Types].xml
    builder.content_types.insert(
        "/word/document.xml".into(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml".into(),
    );
    builder.content_types.insert(
        "/word/styles.xml".into(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml".into(),
    );
    builder.content_types.insert(
        "/word/settings.xml".into(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml".into(),
    );
    if builder.has_numbering {
        builder.content_types.insert(
            "/word/numbering.xml".into(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml".into(),
        );
    }
    if builder.has_footnotes {
        builder.content_types.insert(
            "/word/footnotes.xml".into(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml".into(),
        );
    }
    if builder.has_comments {
        builder.content_types.insert(
            "/word/comments.xml".into(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml".into(),
        );
    }
    for (_, filename, _) in &builder.header_parts {
        builder.content_types.insert(
            format!("/word/{filename}"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml".into(),
        );
    }
    for (_, filename, _) in &builder.footer_parts {
        builder.content_types.insert(
            format!("/word/{filename}"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml".into(),
        );
    }
    if spec.metadata.is_some() {
        builder.content_types.insert(
            "/docProps/core.xml".into(),
            "application/vnd.openxmlformats-package.core-properties+xml".into(),
        );
    }

    let content_types_xml = build_content_types(&builder.content_types, OOXML_IMAGE_EXTENSIONS);
    let doc_rels_xml = builder.doc_rels.to_xml();

    // Package rels
    let mut pkg_rels = RelationshipManager::new();
    pkg_rels.add(REL_OFFICE_DOCUMENT, "word/document.xml");
    if spec.metadata.is_some() {
        pkg_rels.add(REL_CORE_PROPS, "docProps/core.xml");
    }
    let pkg_rels_xml = pkg_rels.to_xml();

    // Write ZIP
    let mut zip = zip::ZipWriter::new(writer);
    let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", opts)?;
    zip.write_all(content_types_xml.as_bytes())?;

    zip.start_file("_rels/.rels", opts)?;
    zip.write_all(pkg_rels_xml.as_bytes())?;

    zip.start_file("word/document.xml", opts)?;
    zip.write_all(document_xml.as_bytes())?;

    zip.start_file("word/_rels/document.xml.rels", opts)?;
    zip.write_all(doc_rels_xml.as_bytes())?;

    zip.start_file("word/styles.xml", opts)?;
    zip.write_all(styles_xml.as_bytes())?;

    zip.start_file("word/settings.xml", opts)?;
    zip.write_all(settings_xml.as_bytes())?;

    if builder.has_numbering {
        let numbering_xml = build_numbering_xml(&builder);
        zip.start_file("word/numbering.xml", opts)?;
        zip.write_all(numbering_xml.as_bytes())?;
    }

    if builder.has_footnotes {
        let footnotes_xml = build_footnotes_xml(&builder);
        zip.start_file("word/footnotes.xml", opts)?;
        zip.write_all(footnotes_xml.as_bytes())?;
    }

    if builder.has_comments {
        let comments_xml = build_comments_xml(&builder);
        zip.start_file("word/comments.xml", opts)?;
        zip.write_all(comments_xml.as_bytes())?;
    }

    for (_, filename, xml) in &builder.header_parts {
        zip.start_file(format!("word/{filename}"), opts)?;
        zip.write_all(xml.as_bytes())?;
    }
    for (_, filename, xml) in &builder.footer_parts {
        zip.start_file(format!("word/{filename}"), opts)?;
        zip.write_all(xml.as_bytes())?;
    }

    // Images
    for (path, data) in &builder.images {
        zip.start_file(format!("word/{path}"), opts)?;
        zip.write_all(data)?;
    }

    // Metadata
    if let Some(metadata) = &spec.metadata {
        let core_xml = build_core_xml(metadata);
        zip.start_file("docProps/core.xml", opts)?;
        zip.write_all(core_xml.as_bytes())?;
    }

    zip.finish()?;
    Ok(())
}

fn build_document_xml(
    spec: &DocSpec,
    builder: &mut DocxBuilder,
    assets_dir: Option<&Path>,
) -> Result<String> {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<w:document xmlns:w="{NS_W}" xmlns:r="{NS_R}" xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" xmlns:pic="{NS_PIC}">"#
    ));
    xml.push_str("<w:body>");

    for block in &spec.body {
        write_block(&mut xml, block, builder, &spec.styles, assets_dir, 0)?;
    }

    // Section properties — OOXML requires specific child element order:
    // headerReference/footerReference, then pgSz, pgMar, then titlePg
    xml.push_str("<w:sectPr>");
    write_header_footer_refs(&mut xml, builder);
    write_page_setup(&mut xml, spec.page.as_ref());
    // titlePg must come after pgMar
    if builder.has_headers || builder.has_footers {
        let has_first = builder
            .header_parts
            .iter()
            .any(|(_, f, _)| f.contains("first"))
            || builder
                .footer_parts
                .iter()
                .any(|(_, f, _)| f.contains("first"));
        if has_first {
            xml.push_str("<w:titlePg/>");
        }
    }
    xml.push_str("</w:sectPr>");

    xml.push_str("</w:body></w:document>");
    Ok(xml)
}

fn write_page_setup(xml: &mut String, page: Option<&PageSetup>) {
    let (w, h) = match page.and_then(|p| p.size.as_ref()) {
        Some(PageSize::Named(name)) => match name.as_str() {
            "letter" => (12240, 15840),
            "legal" => (12240, 20160),
            "a4" => (11906, 16838),
            _ => (12240, 15840),
        },
        Some(PageSize::Custom { width, height }) => {
            (inches_to_dxa(*width), inches_to_dxa(*height))
        }
        None => (12240, 15840),
    };

    let is_landscape = matches!(
        page.and_then(|p| p.orientation.as_ref()),
        Some(Orientation::Landscape)
    );

    let (pw, ph) = if is_landscape { (h, w) } else { (w, h) };

    // Only emit w:orient for landscape; portrait is the default
    if is_landscape {
        xml.push_str(&format!(
            r#"<w:pgSz w:w="{pw}" w:h="{ph}" w:orient="landscape"/>"#
        ));
    } else {
        xml.push_str(&format!(r#"<w:pgSz w:w="{pw}" w:h="{ph}"/>"#));
    }

    let (top, bottom, left, right) = match page.and_then(|p| p.margin.as_ref()) {
        Some(Margin::Uniform(m)) => {
            let d = inches_to_dxa(*m);
            (d, d, d, d)
        }
        Some(Margin::Custom {
            top,
            bottom,
            left,
            right,
        }) => (
            inches_to_dxa(top.unwrap_or(1.0)),
            inches_to_dxa(bottom.unwrap_or(1.0)),
            inches_to_dxa(left.unwrap_or(1.25)),
            inches_to_dxa(right.unwrap_or(1.25)),
        ),
        None => (
            inches_to_dxa(1.0),
            inches_to_dxa(1.0),
            inches_to_dxa(1.25),
            inches_to_dxa(1.25),
        ),
    };

    xml.push_str(&format!(
        r#"<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}" w:header="720" w:footer="720" w:gutter="0"/>"#
    ));
}

fn write_header_footer_refs(xml: &mut String, builder: &DocxBuilder) {
    for (rel_id, filename, _) in &builder.header_parts {
        let hf_type = if filename.contains("first") {
            "first"
        } else if filename.contains("even") {
            "even"
        } else {
            "default"
        };
        xml.push_str(&format!(
            r#"<w:headerReference w:type="{hf_type}" r:id="{rel_id}"/>"#
        ));
    }
    for (rel_id, filename, _) in &builder.footer_parts {
        let hf_type = if filename.contains("first") {
            "first"
        } else if filename.contains("even") {
            "even"
        } else {
            "default"
        };
        xml.push_str(&format!(
            r#"<w:footerReference w:type="{hf_type}" r:id="{rel_id}"/>"#
        ));
    }
}

fn write_block(
    xml: &mut String,
    block: &Block,
    builder: &mut DocxBuilder,
    styles: &Option<Styles>,
    assets_dir: Option<&Path>,
    _list_level: u32,
) -> Result<()> {
    match block {
        Block::Heading { heading, text, id } => {
            xml.push_str("<w:p><w:pPr>");
            xml.push_str(&format!(r#"<w:pStyle w:val="Heading{heading}"/>"#));
            xml.push_str("</w:pPr>");

            if let Some(bookmark_id) = id {
                let bid = bookmark_id.len() as u32; // simple numeric id
                xml.push_str(&format!(
                    r#"<w:bookmarkStart w:id="{bid}" w:name="{bookmark_id}"/>"#
                ));
            }

            let text_str = text.as_str();
            let runs = parse_inline_text(text_str);
            write_runs(xml, &runs, builder)?;

            if id.is_some() {
                let bid = id.as_ref().unwrap().len() as u32;
                xml.push_str(&format!(r#"<w:bookmarkEnd w:id="{bid}"/>"#));
            }

            xml.push_str("</w:p>");
        }

        Block::Paragraph { paragraph } => {
            write_paragraph(xml, paragraph, builder, styles, None)?;
        }

        Block::Bullets { bullets } => {
            builder.has_numbering = true;
            for item in bullets {
                write_list_item(xml, item, builder, styles, assets_dir, 0, true)?;
            }
        }

        Block::Numbered {
            numbered, restart, ..
        } => {
            builder.has_numbering = true;
            let num_id = if restart.unwrap_or(false) {
                let id = builder.num_id_counter;
                builder.num_id_counter += 1;
                builder.extra_num_entries.push(id);
                id
            } else {
                2 // default numbered list numId
            };
            for item in numbered {
                write_list_item_with_num(xml, item, builder, styles, assets_dir, 0, num_id)?;
            }
        }

        Block::Table { table, header_rows } => {
            write_table(xml, table, header_rows.unwrap_or(0), builder, styles, assets_dir)?;
        }

        Block::Image {
            image,
            width,
            height,
            alt,
            align,
            caption,
            image_data,
        } => {
            write_image(xml, image, *width, *height, alt.as_deref(), align.as_deref(), caption.as_deref(), image_data.as_deref(), builder, assets_dir)?;
        }

        Block::PageBreak { .. } => {
            xml.push_str(
                r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#,
            );
        }

        Block::Toc { toc } => {
            write_toc(xml, toc);
        }

        Block::SectionBreak { section_break } => {
            write_section_break(xml, section_break);
        }

        Block::Bookmark { bookmark } => {
            let bid = bookmark.len() as u32;
            xml.push_str(&format!(
                r#"<w:p><w:bookmarkStart w:id="{bid}" w:name="{bookmark}"/><w:bookmarkEnd w:id="{bid}"/></w:p>"#
            ));
        }

        Block::Raw { raw } => {
            xml.push_str(raw);
        }
    }
    Ok(())
}

fn write_paragraph(
    xml: &mut String,
    content: &ParagraphContent,
    builder: &mut DocxBuilder,
    _styles: &Option<Styles>,
    list_info: Option<(u32, bool)>,
) -> Result<()> {
    match content {
        ParagraphContent::Simple(text) => {
            xml.push_str("<w:p>");
            if let Some((level, is_bullet)) = list_info {
                xml.push_str("<w:pPr>");
                write_list_ppr(xml, level, is_bullet);
                xml.push_str("</w:pPr>");
            }
            let runs = parse_inline_text(text);
            write_runs(xml, &runs, builder)?;
            xml.push_str("</w:p>");
        }
        ParagraphContent::Full(full) => {
            xml.push_str("<w:p>");
            let has_ppr = full.align.is_some()
                || full.spacing.is_some()
                || full.indent.is_some()
                || full.style.is_some()
                || full.inserted.is_some()
                || full.deleted.is_some()
                || full.id.is_some()
                || list_info.is_some();

            if has_ppr {
                xml.push_str("<w:pPr>");
                // Schema order: pStyle, numPr, spacing, ind, jc, rPr
                if let Some(style) = &full.style {
                    xml.push_str(&format!(r#"<w:pStyle w:val="{style}"/>"#));
                }
                if let Some((level, is_bullet)) = list_info {
                    write_list_ppr(xml, level, is_bullet);
                }
                if let Some(spacing) = &full.spacing {
                    write_spacing(xml, spacing);
                }
                if let Some(indent) = &full.indent {
                    write_indent(xml, indent);
                }
                if let Some(align) = &full.align {
                    let val = match align.as_str() {
                        "center" => "center",
                        "right" => "right",
                        "justify" => "both",
                        _ => "left",
                    };
                    xml.push_str(&format!(r#"<w:jc w:val="{val}"/>"#));
                }
                if let Some(change) = &full.inserted {
                    xml.push_str(&format!(
                        r#"<w:rPr><w:ins w:id="0" w:author="{}" w:date="{}"/></w:rPr>"#,
                        escape_xml(change.author.as_deref().unwrap_or("Author")),
                        escape_xml(change.date.as_deref().unwrap_or(""))
                    ));
                }
                if let Some(change) = &full.deleted {
                    xml.push_str(&format!(
                        r#"<w:rPr><w:del w:id="0" w:author="{}" w:date="{}"/></w:rPr>"#,
                        escape_xml(change.author.as_deref().unwrap_or("Author")),
                        escape_xml(change.date.as_deref().unwrap_or(""))
                    ));
                }
                xml.push_str("</w:pPr>");
            }

            if let Some(bookmark_id) = &full.id {
                let bid = bookmark_id.len() as u32;
                xml.push_str(&format!(
                    r#"<w:bookmarkStart w:id="{bid}" w:name="{bookmark_id}"/>"#
                ));
            }

            if let Some(text) = &full.text {
                let runs = parse_inline_text(text);
                write_runs(xml, &runs, builder)?;
            } else if let Some(runs) = &full.runs {
                write_runs(xml, runs, builder)?;
            }

            if let Some(bookmark_id) = &full.id {
                let bid = bookmark_id.len() as u32;
                xml.push_str(&format!(r#"<w:bookmarkEnd w:id="{bid}"/>"#));
            }

            xml.push_str("</w:p>");
        }
    }
    Ok(())
}

fn write_runs(xml: &mut String, runs: &[Run], builder: &mut DocxBuilder) -> Result<()> {
    for run in runs {
        match run {
            Run::Text(tr) => {
                if let Some(url) = &tr.link {
                    if url.starts_with('#') {
                        // Internal bookmark link
                        let anchor = &url[1..];
                        xml.push_str(&format!(
                            r#"<w:hyperlink w:anchor="{anchor}">"#
                        ));
                        xml.push_str("<w:r><w:rPr>");
                        xml.push_str(r#"<w:rStyle w:val="Hyperlink"/>"#);
                        write_run_properties(xml, tr);
                        xml.push_str("</w:rPr>");
                        xml.push_str(&format!(
                            r#"<w:t xml:space="preserve">{}</w:t>"#,
                            escape_xml(&tr.text)
                        ));
                        xml.push_str("</w:r></w:hyperlink>");
                    } else {
                        // External hyperlink
                        let rel_id = builder.add_hyperlink(url);
                        xml.push_str(&format!(
                            r#"<w:hyperlink r:id="{rel_id}">"#
                        ));
                        xml.push_str("<w:r><w:rPr>");
                        xml.push_str(r#"<w:rStyle w:val="Hyperlink"/>"#);
                        write_run_properties(xml, tr);
                        xml.push_str("</w:rPr>");
                        xml.push_str(&format!(
                            r#"<w:t xml:space="preserve">{}</w:t>"#,
                            escape_xml(&tr.text)
                        ));
                        xml.push_str("</w:r></w:hyperlink>");
                    }
                } else {
                    xml.push_str("<w:r>");
                    let has_rpr = tr.bold.is_some()
                        || tr.italic.is_some()
                        || tr.underline.is_some()
                        || tr.strike.is_some()
                        || tr.superscript.is_some()
                        || tr.subscript.is_some()
                        || tr.font.is_some()
                        || tr.size.is_some()
                        || tr.color.is_some()
                        || tr.highlight.is_some()
                        || tr.all_caps.is_some()
                        || tr.small_caps.is_some();
                    if has_rpr {
                        xml.push_str("<w:rPr>");
                        write_run_properties(xml, tr);
                        xml.push_str("</w:rPr>");
                    }
                    xml.push_str(&format!(
                        r#"<w:t xml:space="preserve">{}</w:t>"#,
                        escape_xml(&tr.text)
                    ));
                    xml.push_str("</w:r>");
                }
            }
            Run::Tab(_) => {
                xml.push_str("<w:r><w:tab/></w:r>");
            }
            Run::Field(fr) => {
                match fr.field.as_str() {
                    "page-number" => {
                        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#);
                        xml.push_str(
                            r#"<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>"#,
                        );
                        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#);
                    }
                    "total-pages" => {
                        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#);
                        xml.push_str(
                            r#"<w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>"#,
                        );
                        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#);
                    }
                    "date" => {
                        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#);
                        xml.push_str(
                            r#"<w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>"#,
                        );
                        xml.push_str(r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#);
                    }
                    _ => {}
                }
            }
            Run::FootnoteRef(fr) => {
                let fid = fr.footnote.parse::<u32>().unwrap_or(1);
                xml.push_str("<w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr>");
                xml.push_str(&format!(r#"<w:footnoteReference w:id="{fid}"/>"#));
                xml.push_str("</w:r>");
            }
            Run::Delete(dr) => {
                xml.push_str(&format!(
                    r#"<w:del w:id="0" w:author="{}" w:date="{}">"#,
                    escape_xml(dr.author.as_deref().unwrap_or("Author")),
                    escape_xml(dr.date.as_deref().unwrap_or(""))
                ));
                xml.push_str(&format!(
                    r#"<w:r><w:delText xml:space="preserve">{}</w:delText></w:r>"#,
                    escape_xml(&dr.delete)
                ));
                xml.push_str("</w:del>");
            }
            Run::Insert(ir) => {
                xml.push_str(&format!(
                    r#"<w:ins w:id="0" w:author="{}" w:date="{}">"#,
                    escape_xml(ir.author.as_deref().unwrap_or("Author")),
                    escape_xml(ir.date.as_deref().unwrap_or(""))
                ));
                xml.push_str(&format!(
                    r#"<w:r><w:t xml:space="preserve">{}</w:t></w:r>"#,
                    escape_xml(&ir.insert)
                ));
                xml.push_str("</w:ins>");
            }
            Run::CommentStart(cs) => {
                if let Some(cid) = builder.comment_id_map.get(&cs.comment_start) {
                    xml.push_str(&format!(r#"<w:commentRangeStart w:id="{cid}"/>"#));
                }
            }
            Run::CommentEnd(ce) => {
                if let Some(cid) = builder.comment_id_map.get(&ce.comment_end) {
                    xml.push_str(&format!(r#"<w:commentRangeEnd w:id="{cid}"/>"#));
                    xml.push_str(&format!(
                        r#"<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>"#
                    ));
                }
            }
            Run::Break(br) => {
                let break_type = match br.break_type.as_str() {
                    "page" => r#"w:type="page""#,
                    "column" => r#"w:type="column""#,
                    _ => "",
                };
                xml.push_str(&format!(r#"<w:r><w:br {break_type}/></w:r>"#));
            }
        }
    }
    Ok(())
}

fn write_run_properties(xml: &mut String, tr: &TextRun) {
    if let Some(font) = &tr.font {
        xml.push_str(&format!(
            r#"<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>"#
        ));
    }
    if tr.bold == Some(true) {
        xml.push_str("<w:b/>");
    }
    if tr.italic == Some(true) {
        xml.push_str("<w:i/>");
    }
    if tr.underline == Some(true) {
        xml.push_str(r#"<w:u w:val="single"/>"#);
    }
    if tr.strike == Some(true) {
        xml.push_str("<w:strike/>");
    }
    if tr.superscript == Some(true) {
        xml.push_str(r#"<w:vertAlign w:val="superscript"/>"#);
    }
    if tr.subscript == Some(true) {
        xml.push_str(r#"<w:vertAlign w:val="subscript"/>"#);
    }
    if let Some(size) = tr.size {
        let hp = points_to_half_points(size);
        xml.push_str(&format!(r#"<w:sz w:val="{hp}"/>"#));
    }
    if let Some(color) = &tr.color {
        xml.push_str(&format!(r#"<w:color w:val="{color}"/>"#));
    }
    if let Some(highlight) = &tr.highlight {
        xml.push_str(&format!(r#"<w:highlight w:val="{highlight}"/>"#));
    }
    if tr.all_caps == Some(true) {
        xml.push_str("<w:caps/>");
    }
    if tr.small_caps == Some(true) {
        xml.push_str("<w:smallCaps/>");
    }
}

fn write_spacing(xml: &mut String, spacing: &Spacing) {
    let mut attrs = String::new();
    if let Some(before) = spacing.before {
        attrs.push_str(&format!(r#" w:before="{}""#, points_to_twips(before)));
    }
    if let Some(after) = spacing.after {
        attrs.push_str(&format!(r#" w:after="{}""#, points_to_twips(after)));
    }
    if let Some(line) = spacing.line {
        // Line spacing: in points → twips, with lineRule=exact
        attrs.push_str(&format!(
            r#" w:line="{}" w:lineRule="exact""#,
            points_to_twips(line)
        ));
    }
    if !attrs.is_empty() {
        xml.push_str(&format!(r#"<w:spacing{attrs}/>"#));
    }
}

fn write_indent(xml: &mut String, indent: &Indent) {
    let mut attrs = String::new();
    if let Some(left) = indent.left {
        attrs.push_str(&format!(r#" w:left="{}""#, inches_to_dxa(left)));
    }
    if let Some(right) = indent.right {
        attrs.push_str(&format!(r#" w:right="{}""#, inches_to_dxa(right)));
    }
    if let Some(first_line) = indent.first_line {
        attrs.push_str(&format!(r#" w:firstLine="{}""#, inches_to_dxa(first_line)));
    }
    if let Some(hanging) = indent.hanging {
        attrs.push_str(&format!(r#" w:hanging="{}""#, inches_to_dxa(hanging)));
    }
    if !attrs.is_empty() {
        xml.push_str(&format!(r#"<w:ind{attrs}/>"#));
    }
}

fn write_list_ppr(xml: &mut String, level: u32, is_bullet: bool) {
    let num_id = if is_bullet { 1 } else { 2 };
    xml.push_str(&format!(
        r#"<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>"#
    ));
}

fn write_list_ppr_with_num(xml: &mut String, level: u32, num_id: u32) {
    xml.push_str(&format!(
        r#"<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>"#
    ));
}

fn write_list_item_with_num(
    xml: &mut String,
    item: &ListItem,
    builder: &mut DocxBuilder,
    styles: &Option<Styles>,
    assets_dir: Option<&Path>,
    level: u32,
    num_id: u32,
) -> Result<()> {
    let text = item.text();
    let para = ParagraphContent::Simple(text.to_string());
    // Write paragraph with custom numId
    xml.push_str("<w:p><w:pPr><w:pStyle w:val=\"ListParagraph\"/>");
    write_list_ppr_with_num(xml, level, num_id);
    xml.push_str("</w:pPr>");
    let runs = parse_inline_text(text);
    write_runs(xml, &runs, builder)?;
    xml.push_str("</w:p>");

    if let Some(children) = item.children() {
        for child in children {
            write_list_item_with_num(xml, child, builder, styles, assets_dir, level + 1, num_id)?;
        }
    }
    Ok(())
}

fn write_list_item(
    xml: &mut String,
    item: &ListItem,
    builder: &mut DocxBuilder,
    styles: &Option<Styles>,
    assets_dir: Option<&Path>,
    level: u32,
    is_bullet: bool,
) -> Result<()> {
    let text = item.text();
    let para = ParagraphContent::Simple(text.to_string());
    write_paragraph(xml, &para, builder, styles, Some((level, is_bullet)))?;

    if let Some(children) = item.children() {
        for child in children {
            write_list_item(xml, child, builder, styles, assets_dir, level + 1, is_bullet)?;
        }
    }
    Ok(())
}

fn write_table(
    xml: &mut String,
    table: &TableContent,
    header_rows: u32,
    builder: &mut DocxBuilder,
    styles: &Option<Styles>,
    assets_dir: Option<&Path>,
) -> Result<()> {
    xml.push_str("<w:tbl>");
    xml.push_str("<w:tblPr>");
    xml.push_str(r#"<w:tblStyle w:val="TableGrid"/>"#);
    xml.push_str(r#"<w:tblW w:w="0" w:type="auto"/>"#);
    xml.push_str(
        r#"<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>"#,
    );
    xml.push_str("</w:tblPr>");

    match table {
        TableContent::Simple(rows) => {
            // Column grid
            if let Some(first_row) = rows.first() {
                xml.push_str("<w:tblGrid>");
                for _ in first_row {
                    xml.push_str(r#"<w:gridCol w:w="2000"/>"#);
                }
                xml.push_str("</w:tblGrid>");
            }

            for (ri, row) in rows.iter().enumerate() {
                xml.push_str("<w:tr>");
                if (ri as u32) < header_rows {
                    xml.push_str("<w:trPr><w:tblHeader/></w:trPr>");
                }
                for cell in row {
                    xml.push_str("<w:tc><w:p><w:r>");
                    xml.push_str(&format!(
                        r#"<w:t xml:space="preserve">{}</w:t>"#,
                        escape_xml(cell)
                    ));
                    xml.push_str("</w:r></w:p></w:tc>");
                }
                xml.push_str("</w:tr>");
            }
        }
        TableContent::Full(full) => {
            // Column grid
            if let Some(columns) = &full.columns {
                xml.push_str("<w:tblGrid>");
                for col in columns {
                    let w = inches_to_dxa(col.width);
                    xml.push_str(&format!(r#"<w:gridCol w:w="{w}"/>"#));
                }
                xml.push_str("</w:tblGrid>");
            }

            let hdr_rows = full.header_rows.unwrap_or(header_rows);

            for (ri, row) in full.rows.iter().enumerate() {
                xml.push_str("<w:tr>");
                if (ri as u32) < hdr_rows {
                    xml.push_str("<w:trPr><w:tblHeader/></w:trPr>");
                }
                for cell in &row.cells {
                    xml.push_str("<w:tc>");

                    // Cell properties
                    let has_tc_pr = cell.colspan.is_some()
                        || cell.rowspan.is_some()
                        || cell.shading.is_some()
                        || cell.valign.is_some();
                    if has_tc_pr {
                        xml.push_str("<w:tcPr>");
                        if let Some(colspan) = cell.colspan {
                            if colspan > 1 {
                                xml.push_str(&format!(
                                    r#"<w:gridSpan w:val="{colspan}"/>"#
                                ));
                            }
                        }
                        if let Some(rowspan) = cell.rowspan {
                            if rowspan > 1 {
                                xml.push_str(r#"<w:vMerge w:val="restart"/>"#);
                            }
                        }
                        if let Some(shading) = &cell.shading {
                            xml.push_str(&format!(
                                r#"<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>"#
                            ));
                        }
                        if let Some(valign) = &cell.valign {
                            let val = match valign.as_str() {
                                "center" => "center",
                                "bottom" => "bottom",
                                _ => "top",
                            };
                            xml.push_str(&format!(r#"<w:vAlign w:val="{val}"/>"#));
                        }
                        xml.push_str("</w:tcPr>");
                    }

                    // Cell content
                    if let Some(body) = &cell.body {
                        for block in body {
                            write_block(xml, block, builder, styles, assets_dir, 0)?;
                        }
                    } else if let Some(runs) = &cell.runs {
                        xml.push_str("<w:p>");
                        if let Some(align) = &cell.align {
                            let val = match align.as_str() {
                                "center" => "center",
                                "right" => "right",
                                _ => "left",
                            };
                            xml.push_str(&format!(
                                r#"<w:pPr><w:jc w:val="{val}"/></w:pPr>"#
                            ));
                        }
                        write_runs(xml, runs, builder)?;
                        xml.push_str("</w:p>");
                    } else {
                        let text = cell.text.as_deref().unwrap_or("");
                        xml.push_str("<w:p>");
                        if cell.align.is_some() || cell.bold == Some(true) || cell.color.is_some() {
                            xml.push_str("<w:pPr>");
                            if let Some(align) = &cell.align {
                                let val = match align.as_str() {
                                    "center" => "center",
                                    "right" => "right",
                                    _ => "left",
                                };
                                xml.push_str(&format!(r#"<w:jc w:val="{val}"/>"#));
                            }
                            xml.push_str("</w:pPr>");
                        }
                        xml.push_str("<w:r>");
                        if cell.bold == Some(true) || cell.color.is_some() {
                            xml.push_str("<w:rPr>");
                            if cell.bold == Some(true) {
                                xml.push_str("<w:b/>");
                            }
                            if let Some(color) = &cell.color {
                                xml.push_str(&format!(r#"<w:color w:val="{color}"/>"#));
                            }
                            xml.push_str("</w:rPr>");
                        }
                        xml.push_str(&format!(
                            r#"<w:t xml:space="preserve">{}</w:t>"#,
                            escape_xml(text)
                        ));
                        xml.push_str("</w:r></w:p>");
                    }

                    xml.push_str("</w:tc>");
                }
                xml.push_str("</w:tr>");
            }
        }
    }

    xml.push_str("</w:tbl>");
    Ok(())
}

fn write_image(
    xml: &mut String,
    filename: &str,
    width: Option<f64>,
    height: Option<f64>,
    alt: Option<&str>,
    align: Option<&str>,
    caption: Option<&str>,
    image_data: Option<&str>,
    builder: &mut DocxBuilder,
    assets_dir: Option<&Path>,
) -> Result<()> {
    // Load image data
    let data = if let Some(b64) = image_data {
        use base64::Engine;
        base64::engine::general_purpose::STANDARD
            .decode(b64)
            .context("invalid base64 image data")?
    } else {
        let dir = assets_dir.unwrap_or_else(|| Path::new("."));
        let img_path = dir.join(filename);
        std::fs::read(&img_path)
            .with_context(|| format!("failed to read image: {}", img_path.display()))?
    };

    let rel_id = builder.add_image(filename, data);

    let w_emu = inches_to_emu(width.unwrap_or(4.0));
    let h_emu = inches_to_emu(height.unwrap_or(3.0));
    let alt_text = alt.unwrap_or(filename);

    // Paragraph with optional alignment
    xml.push_str("<w:p>");
    if let Some(a) = align {
        let val = match a {
            "center" => "center",
            "right" => "right",
            _ => "left",
        };
        xml.push_str(&format!(r#"<w:pPr><w:jc w:val="{val}"/></w:pPr>"#));
    }

    xml.push_str("<w:r><w:drawing>");
    xml.push_str(&format!(
        r#"<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{w_emu}" cy="{h_emu}"/><wp:docPr id="1" name="{}" descr="{}"/>"#,
        escape_xml(filename),
        escape_xml(alt_text)
    ));
    xml.push_str(r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">"#);
    xml.push_str(&format!(
        r#"<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="{}"/><pic:cNvPicPr/></pic:nvPicPr>"#,
        escape_xml(filename)
    ));
    xml.push_str(&format!(
        r#"<pic:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"#
    ));
    xml.push_str(&format!(
        r#"<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w_emu}" cy="{h_emu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>"#
    ));
    xml.push_str("</pic:pic></a:graphicData></a:graphic></wp:inline>");
    xml.push_str("</w:drawing></w:r></w:p>");

    // Caption
    if let Some(caption_text) = caption {
        xml.push_str("<w:p><w:pPr>");
        xml.push_str(r#"<w:pStyle w:val="Caption"/>"#);
        if align.is_some() {
            let val = match align.unwrap() {
                "center" => "center",
                "right" => "right",
                _ => "left",
            };
            xml.push_str(&format!(r#"<w:jc w:val="{val}"/>"#));
        }
        xml.push_str("</w:pPr><w:r>");
        xml.push_str(&format!(
            r#"<w:t xml:space="preserve">{}</w:t>"#,
            escape_xml(caption_text)
        ));
        xml.push_str("</w:r></w:p>");
    }

    Ok(())
}

fn write_toc(xml: &mut String, toc: &TocContent) {
    let (title, depth) = match toc {
        TocContent::Simple(_) => ("Table of Contents", 3u8),
        TocContent::Full { title, depth } => (
            title.as_deref().unwrap_or("Table of Contents"),
            depth.unwrap_or(3),
        ),
    };

    // TOC title
    xml.push_str("<w:p><w:pPr><w:pStyle w:val=\"TOCHeading\"/></w:pPr>");
    xml.push_str(&format!(
        r#"<w:r><w:t>{}</w:t></w:r></w:p>"#,
        escape_xml(title)
    ));

    // TOC field
    xml.push_str("<w:p>");
    xml.push_str(r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#);
    xml.push_str(&format!(
        r#"<w:r><w:instrText xml:space="preserve"> TOC \o "1-{depth}" \h \z \u </w:instrText></w:r>"#
    ));
    xml.push_str(r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#);
    xml.push_str(
        r#"<w:r><w:t>[Table of contents will be updated when document is opened]</w:t></w:r>"#,
    );
    xml.push_str(r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#);
    xml.push_str("</w:p>");
}

fn write_section_break(xml: &mut String, config: &SectionBreakConfig) {
    xml.push_str("<w:p><w:pPr><w:sectPr>");

    let break_type = config
        .break_type
        .as_deref()
        .unwrap_or("nextPage");
    let val = match break_type {
        "next-page" | "nextPage" => "nextPage",
        "continuous" => "continuous",
        "even-page" | "evenPage" => "evenPage",
        "odd-page" | "oddPage" => "oddPage",
        _ => "nextPage",
    };
    xml.push_str(&format!(r#"<w:type w:val="{val}"/>"#));

    // Schema order: type, pgSz, pgMar, cols
    if let Some(page) = &config.page {
        if let Some(orient) = &page.orientation {
            let (w, h) = match &page.size {
                Some(PageSize::Named(name)) => match name.as_str() {
                    "letter" => (12240i64, 15840i64),
                    "legal" => (12240, 20160),
                    "a4" => (11906, 16838),
                    _ => (12240, 15840),
                },
                Some(PageSize::Custom { width, height }) => {
                    (inches_to_dxa(*width), inches_to_dxa(*height))
                }
                None => (12240, 15840),
            };
            let (pw, ph, is_land) = match orient {
                Orientation::Landscape => (h, w, true),
                Orientation::Portrait => (w, h, false),
            };
            if is_land {
                xml.push_str(&format!(
                    r#"<w:pgSz w:w="{pw}" w:h="{ph}" w:orient="landscape"/>"#
                ));
            } else {
                xml.push_str(&format!(r#"<w:pgSz w:w="{pw}" w:h="{ph}"/>"#));
            }
        }
    }

    if let Some(cols) = config.columns {
        let gap = inches_to_dxa(config.column_gap.unwrap_or(0.5));
        xml.push_str(&format!(
            r#"<w:cols w:num="{cols}" w:space="{gap}"/>"#
        ));
    }

    if let Some(valign) = &config.valign {
        let val = match valign.as_str() {
            "center" => "center",
            "bottom" => "bottom",
            "both" | "justify" => "both",
            _ => "top",
        };
        xml.push_str(&format!(r#"<w:vAlign w:val="{val}"/>"#));
    }

    xml.push_str("</w:sectPr></w:pPr></w:p>");
}

fn build_styles_xml(spec: &DocSpec) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<w:styles xmlns:w="{NS_W}" xmlns:r="{NS_R}">"#));

    // Document defaults
    xml.push_str("<w:docDefaults><w:rPrDefault><w:rPr>");
    let default_font = spec
        .styles
        .as_ref()
        .and_then(|s| s.font.as_deref())
        .unwrap_or("Calibri");
    let default_size = spec
        .styles
        .as_ref()
        .and_then(|s| s.size)
        .unwrap_or(11.0);
    xml.push_str(&format!(
        r#"<w:rFonts w:ascii="{default_font}" w:hAnsi="{default_font}" w:eastAsia="{default_font}" w:cs="{default_font}"/>"#
    ));
    xml.push_str(&format!(
        r#"<w:sz w:val="{}"/>"#,
        points_to_half_points(default_size)
    ));
    if let Some(color) = spec.styles.as_ref().and_then(|s| s.color.as_deref()) {
        xml.push_str(&format!(r#"<w:color w:val="{color}"/>"#));
    }
    xml.push_str("</w:rPr></w:rPrDefault>");

    // Paragraph defaults
    xml.push_str("<w:pPrDefault><w:pPr>");
    xml.push_str(r#"<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>"#);
    xml.push_str("</w:pPr></w:pPrDefault>");
    xml.push_str("</w:docDefaults>");

    // Normal style
    xml.push_str(r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#);

    // Heading styles
    for level in 1..=6 {
        let heading_styles = spec.styles.as_ref().and_then(|s| s.headings.as_ref());
        let level_style = heading_styles.and_then(|h| match level {
            1 => h.h1.as_ref(),
            2 => h.h2.as_ref(),
            3 => h.h3.as_ref(),
            4 => h.h4.as_ref(),
            5 => h.h5.as_ref(),
            6 => h.h6.as_ref(),
            _ => None,
        });

        let font = level_style
            .and_then(|l| l.font.as_deref())
            .or_else(|| heading_styles.and_then(|h| h.font.as_deref()));
        let color = level_style
            .and_then(|l| l.color.as_deref())
            .or_else(|| heading_styles.and_then(|h| h.color.as_deref()));
        let size = level_style.and_then(|l| l.size);
        let bold = level_style.and_then(|l| l.bold);

        // Default sizes if not specified
        let default_size = match level {
            1 => 32.0,
            2 => 26.0,
            3 => 24.0,
            4 => 22.0,
            5 => 20.0,
            6 => 18.0,
            _ => 24.0,
        };

        // Heading spacing: more space before than after for visual separation
        let (sp_before, sp_after) = match level {
            1 => (360, 200), // 18pt before, 10pt after
            2 => (300, 120), // 15pt before, 6pt after
            3 => (240, 80),  // 12pt before, 4pt after
            _ => (200, 60),  // 10pt before, 3pt after
        };
        xml.push_str(&format!(
            r#"<w:style w:type="paragraph" w:styleId="Heading{level}"><w:name w:val="heading {level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="{sp_before}" w:after="{sp_after}"/><w:outlineLvl w:val="{}"/></w:pPr><w:rPr>"#,
            level - 1
        ));

        if let Some(f) = font {
            xml.push_str(&format!(
                r#"<w:rFonts w:ascii="{f}" w:hAnsi="{f}"/>"#
            ));
        }
        if bold.unwrap_or(true) {
            xml.push_str("<w:b/>");
        }
        let s = size.unwrap_or(default_size);
        xml.push_str(&format!(
            r#"<w:sz w:val="{}"/>"#,
            points_to_half_points(s)
        ));
        if let Some(c) = color {
            xml.push_str(&format!(r#"<w:color w:val="{c}"/>"#));
        }

        xml.push_str("</w:rPr></w:style>");
    }

    // Hyperlink style
    xml.push_str(
        r#"<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>"#,
    );

    // Caption style
    xml.push_str(
        r#"<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:sz w:val="20"/></w:rPr></w:style>"#,
    );

    // TOC heading style
    xml.push_str(
        r#"<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/></w:style>"#,
    );

    // Footnote styles
    xml.push_str(
        r#"<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>"#,
    );
    xml.push_str(
        r#"<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>"#,
    );

    // Comment reference style
    xml.push_str(
        r#"<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>"#,
    );

    // Custom styles
    if let Some(styles) = &spec.styles {
        if let Some(custom) = &styles.custom {
            for (name, cs) in custom {
                xml.push_str(&format!(
                    r#"<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/><w:basedOn w:val="Normal"/>"#
                ));
                // Only emit pPr if there are paragraph properties
                let has_ppr = cs.indent.is_some() || cs.spacing.is_some() || cs.align.is_some();
                if has_ppr {
                    xml.push_str("<w:pPr>");
                    if let Some(indent) = &cs.indent {
                        write_indent(&mut xml, indent);
                    }
                    if let Some(spacing) = &cs.spacing {
                        write_spacing(&mut xml, spacing);
                    }
                    if let Some(align) = &cs.align {
                        let val = match align.as_str() {
                            "center" => "center",
                            "right" => "right",
                            "justify" => "both",
                            _ => "left",
                        };
                        xml.push_str(&format!(r#"<w:jc w:val="{val}"/>"#));
                    }
                    xml.push_str("</w:pPr>");
                }
                xml.push_str("<w:rPr>");
                if let Some(font) = &cs.font {
                    xml.push_str(&format!(
                        r#"<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>"#
                    ));
                }
                if let Some(size) = cs.size {
                    xml.push_str(&format!(
                        r#"<w:sz w:val="{}"/>"#,
                        points_to_half_points(size)
                    ));
                }
                if let Some(color) = &cs.color {
                    xml.push_str(&format!(r#"<w:color w:val="{color}"/>"#));
                }
                if cs.bold == Some(true) {
                    xml.push_str("<w:b/>");
                }
                if cs.italic == Some(true) {
                    xml.push_str("<w:i/>");
                }
                xml.push_str("</w:rPr></w:style>");
            }
        }
    }

    // Table style
    xml.push_str(
        r#"<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>"#,
    );

    // List paragraph style
    xml.push_str(
        r#"<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>"#,
    );

    xml.push_str("</w:styles>");
    xml
}

fn build_settings_xml() -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<w:settings xmlns:w="{NS_W}">"#));
    xml.push_str(r#"<w:defaultTabStop w:val="720"/>"#);
    xml.push_str(r#"<w:characterSpacingControl w:val="doNotCompress"/>"#);
    xml.push_str(r#"<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>"#);
    xml.push_str("</w:settings>");
    xml
}

fn build_numbering_xml(builder: &DocxBuilder) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<w:numbering xmlns:w="{NS_W}">"#));

    // Abstract numbering for bullets (id=0)
    xml.push_str(r#"<w:abstractNum w:abstractNumId="0">"#);
    let bullet_chars = ["●", "○", "■", "●", "○", "■", "●", "○", "■"];
    for (i, ch) in bullet_chars.iter().enumerate() {
        xml.push_str(&format!(
            r#"<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="{ch}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{}" w:hanging="360"/></w:pPr></w:lvl>"#,
            720 * (i as i32 + 1)
        ));
    }
    xml.push_str("</w:abstractNum>");

    // Abstract numbering for numbered lists (id=1)
    xml.push_str(r#"<w:abstractNum w:abstractNumId="1">"#);
    let num_formats = [
        "decimal",
        "lowerLetter",
        "lowerRoman",
        "decimal",
        "lowerLetter",
        "lowerRoman",
    ];
    let num_texts = ["%1.", "%2.", "%3.", "%4.", "%5.", "%6."];
    for i in 0..9 {
        let fmt = num_formats.get(i).unwrap_or(&"decimal");
        let text = num_texts.get(i).unwrap_or(&"%1.");
        xml.push_str(&format!(
            r#"<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{}" w:hanging="360"/></w:pPr></w:lvl>"#,
            720 * (i as i32 + 1)
        ));
    }
    xml.push_str("</w:abstractNum>");

    // Concrete numbering instances
    xml.push_str(r#"<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#);
    xml.push_str(r#"<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>"#);

    // Extra numbered list instances with restart override
    for &num_id in &builder.extra_num_entries {
        xml.push_str(&format!(
            r#"<w:num w:numId="{num_id}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>"#
        ));
    }

    xml.push_str("</w:numbering>");
    xml
}

fn build_footnotes_xml(builder: &DocxBuilder) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<w:footnotes xmlns:w="{NS_W}">"#));

    // Separator footnotes (required by Word)
    xml.push_str(r#"<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"#);
    xml.push_str(r#"<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"#);

    for footnote_xml in &builder.footnotes_xml {
        xml.push_str(footnote_xml);
    }

    xml.push_str("</w:footnotes>");
    xml
}

fn build_comments_xml(builder: &DocxBuilder) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<w:comments xmlns:w="{NS_W}">"#));
    for comment_xml in &builder.comments_xml {
        xml.push_str(comment_xml);
    }
    xml.push_str("</w:comments>");
    xml
}

fn build_header_footer_parts(
    builder: &mut DocxBuilder,
    set: &HeaderFooterSet,
    is_header: bool,
    styles: &Option<Styles>,
    assets_dir: Option<&Path>,
) -> Result<()> {
    let prefix = if is_header { "header" } else { "footer" };

    let variants: Vec<(&str, Option<&Vec<Block>>)> = vec![
        ("default", set.default.as_ref()),
        ("first", set.first.as_ref()),
        ("even", set.even.as_ref()),
    ];

    for (variant, blocks) in variants {
        if let Some(blocks) = blocks {
            let filename = format!("{prefix}-{variant}.xml");
            let rel_type = if is_header { REL_HEADER } else { REL_FOOTER };
            let rel_id = builder.doc_rels.add(rel_type, &filename);

            let tag = if is_header { "hdr" } else { "ftr" };
            let mut xml = format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{tag} xmlns:w="{NS_W}" xmlns:r="{NS_R}" xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" xmlns:pic="{NS_PIC}">"#);

            for block in blocks {
                write_block(&mut xml, block, builder, styles, assets_dir, 0)?;
            }

            // Ensure at least one paragraph
            if blocks.is_empty() {
                xml.push_str("<w:p/>");
            }

            xml.push_str(&format!("</w:{tag}>"));

            if is_header {
                builder.has_headers = true;
                builder.header_parts.push((rel_id, filename, xml));
            } else {
                builder.has_footers = true;
                builder.footer_parts.push((rel_id, filename, xml));
            }
        }
    }

    Ok(())
}

fn build_core_xml(metadata: &Metadata) -> String {
    let now = chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ").to_string();
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<cp:coreProperties xmlns:cp="{NS_CP}" xmlns:dc="{NS_DC}" xmlns:dcterms="{NS_DCTERMS}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#
    ));

    if let Some(title) = &metadata.title {
        xml.push_str(&format!("<dc:title>{}</dc:title>", escape_xml(title)));
    }
    if let Some(subject) = &metadata.subject {
        xml.push_str(&format!(
            "<dc:subject>{}</dc:subject>",
            escape_xml(subject)
        ));
    }
    if let Some(creator) = &metadata.creator {
        xml.push_str(&format!(
            "<dc:creator>{}</dc:creator>",
            escape_xml(creator)
        ));
    }
    if let Some(description) = &metadata.description {
        xml.push_str(&format!(
            "<dc:description>{}</dc:description>",
            escape_xml(description)
        ));
    }
    if let Some(keywords) = &metadata.keywords {
        xml.push_str(&format!(
            "<cp:keywords>{}</cp:keywords>",
            escape_xml(&keywords.join(", "))
        ));
    }
    if let Some(category) = &metadata.category {
        xml.push_str(&format!(
            "<cp:category>{}</cp:category>",
            escape_xml(category)
        ));
    }
    // Word expects created/modified dates
    xml.push_str(&format!(
        r#"<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>"#
    ));
    xml.push_str(&format!(
        r#"<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>"#
    ));
    xml.push_str("<cp:revision>1</cp:revision>");

    xml.push_str("</cp:coreProperties>");
    xml
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
