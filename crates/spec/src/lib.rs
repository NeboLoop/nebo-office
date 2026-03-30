pub mod types;
pub mod validate;
pub mod xlsx_types;
pub mod pptx_types;

pub use types::*;
pub use validate::validate_spec;
// xlsx_types and pptx_types are accessed via nebo_spec::xlsx_types::* or nebo_spec::XlsxSpec etc.
// We selectively re-export root-level spec types to avoid name collisions (e.g. ColumnDef).
pub use xlsx_types::XlsxSpec;
pub use xlsx_types::XlsxMetadata;
pub use pptx_types::PptxSpec;
pub use pptx_types::PptxMetadata;
