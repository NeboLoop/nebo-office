use serde::{Deserialize, Serialize};

/// Root XLSX specification.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XlsxSpec {
    pub version: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub metadata: Option<XlsxMetadata>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub styles: Option<XlsxStyles>,
    pub sheets: Vec<Sheet>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub named_ranges: Vec<NamedRange>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XlsxMetadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub creator: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XlsxStyles {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<ColumnDef>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rows: Vec<Row>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub merged: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub freeze: Option<FreezePane>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub conditional: Vec<ConditionalFormat>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub charts: Vec<ChartSpec>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub images: Vec<SheetImage>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub autofilter: Option<AutoFilter>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub validations: Vec<DataValidation>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print: Option<PrintSetup>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnDef {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Row {
    pub cells: Vec<CellValue>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
}

/// A cell value — either a simple string/number or a rich cell object.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum CellValue {
    String(String),
    Number(f64),
    Bool(bool),
    Null,
    Rich(RichCell),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RichCell {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<serde_json::Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub valign: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub colspan: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FreezePane {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConditionalFormat {
    pub range: String,
    pub rule: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<serde_json::Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<ConditionalStyle>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConditionalStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartSpec {
    #[serde(rename = "type")]
    pub chart_type: String,
    pub data: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub labels: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetImage {
    pub image: String,
    pub cell: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AutoFilter {
    pub range: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataValidation {
    pub range: String,
    #[serde(rename = "type")]
    pub validation_type: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub values: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PrintSetup {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<String>,
    #[serde(default, rename = "fit-to-page", skip_serializing_if = "Option::is_none")]
    pub fit_to_page: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NamedRange {
    pub name: String,
    pub range: String,
}
