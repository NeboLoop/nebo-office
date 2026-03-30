/// DOCX-specific relationship constants.
/// RelationshipManager and build_content_types are in nebo-office-core.
pub use nebo_office_core::relationships::{
    build_content_types, RelationshipManager, Relationship,
    NS_RELATIONSHIPS, REL_CORE_PROPS, REL_IMAGE, REL_HYPERLINK,
    OOXML_IMAGE_EXTENSIONS,
};

// DOCX-specific relationship types
pub const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
pub const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
pub const REL_SETTINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
pub const REL_NUMBERING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
pub const REL_FOOTNOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
pub const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
pub const REL_HEADER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
pub const REL_FOOTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
