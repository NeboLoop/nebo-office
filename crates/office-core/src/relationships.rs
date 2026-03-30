/// Shared OOXML relationship management (*.rels files and [Content_Types].xml).
use std::collections::HashMap;

pub const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";

/// Common relationship type used across all OOXML formats.
pub const REL_CORE_PROPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
pub const REL_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

pub struct RelationshipManager {
    counter: u32,
    pub relationships: Vec<Relationship>,
}

pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

impl RelationshipManager {
    pub fn new() -> Self {
        Self {
            counter: 0,
            relationships: Vec::new(),
        }
    }

    pub fn add(&mut self, rel_type: &str, target: &str) -> String {
        self.counter += 1;
        let id = format!("rId{}", self.counter);
        self.relationships.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        });
        id
    }

    pub fn add_external(&mut self, rel_type: &str, target: &str) -> String {
        self.counter += 1;
        let id = format!("rId{}", self.counter);
        self.relationships.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: Some("External".to_string()),
        });
        id
    }

    pub fn to_xml(&self) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        );
        xml.push_str(&format!(
            r#"<Relationships xmlns="{NS_RELATIONSHIPS}">"#
        ));
        for rel in &self.relationships {
            xml.push_str(&format!(
                r#"<Relationship Id="{}" Type="{}" Target="{}""#,
                rel.id, rel.rel_type, rel.target
            ));
            if let Some(mode) = &rel.target_mode {
                xml.push_str(&format!(r#" TargetMode="{mode}""#));
            }
            xml.push_str("/>");
        }
        xml.push_str("</Relationships>");
        xml
    }
}

/// Build a [Content_Types].xml document.
///
/// `parts` maps part names (e.g. "word/document.xml") to content types.
/// `image_extensions` controls which image Default entries are emitted (pass empty slice to skip).
pub fn build_content_types(
    parts: &HashMap<String, String>,
    image_extensions: &[(&str, &str)],
) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    );
    xml.push_str(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
    );

    // Default extensions (always needed for OOXML)
    xml.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    xml.push_str(
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    );

    // Image defaults
    for (ext, ct) in image_extensions {
        xml.push_str(&format!(
            r#"<Default Extension="{ext}" ContentType="{ct}"/>"#
        ));
    }

    // Override parts
    for (part, content_type) in parts {
        xml.push_str(&format!(
            r#"<Override PartName="{part}" ContentType="{content_type}"/>"#
        ));
    }

    xml.push_str("</Types>");
    xml
}

/// Standard image extension defaults used by DOCX, XLSX, and PPTX.
pub const OOXML_IMAGE_EXTENSIONS: &[(&str, &str)] = &[
    ("png", "image/png"),
    ("jpg", "image/jpeg"),
    ("jpeg", "image/jpeg"),
    ("gif", "image/gif"),
    ("bmp", "image/bmp"),
    ("tiff", "image/tiff"),
    ("svg", "image/svg+xml"),
];
