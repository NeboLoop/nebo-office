use serde::{Deserialize, Serialize};

/// Root PPTX specification.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PptxSpec {
    pub version: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub metadata: Option<PptxMetadata>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme: Option<PptxTheme>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<SlideSize>,
    pub slides: Vec<Slide>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PptxMetadata {
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
pub struct PptxTheme {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub colors: Option<ThemeColors>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ThemeColors {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub primary: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent1: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent2: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
}

/// Slide size — named aspect ratio or custom dimensions.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum SlideSize {
    Named(String),
    Custom { width: f64, height: f64 },
}

impl SlideSize {
    /// Returns (width_inches, height_inches).
    pub fn dimensions(&self) -> (f64, f64) {
        match self {
            SlideSize::Named(name) => match name.as_str() {
                "16:9" => (10.0, 5.625),
                "16:10" => (10.0, 6.25),
                "4:3" => (10.0, 7.5),
                _ => (10.0, 5.625), // default to 16:9
            },
            SlideSize::Custom { width, height } => (*width, *height),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Slide {
    pub layout: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subtitle: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub body: Vec<SlideBlock>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shapes: Vec<Shape>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Vec<SlideBlock>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Vec<SlideBlock>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<SlideBackground>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transition: Option<SlideTransition>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub notes: Option<String>,
}

/// Blocks within slide body or columns.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum SlideBlock {
    Paragraph { paragraph: String },
    Bullets { bullets: Vec<String> },
    Numbered { numbered: Vec<String> },
    Table {
        table: Vec<Vec<String>>,
        #[serde(default, rename = "header-rows", skip_serializing_if = "Option::is_none")]
        header_rows: Option<u32>,
    },
    Image {
        image: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        width: Option<f64>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        height: Option<f64>,
    },
    Chart {
        chart: SlideChart,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        x: Option<f64>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        y: Option<f64>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        w: Option<f64>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        h: Option<f64>,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideChart {
    #[serde(rename = "type")]
    pub chart_type: String,
    pub data: serde_json::Value,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    #[serde(rename = "type")]
    pub shape_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
    #[serde(default, rename = "line-color", skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    #[serde(default, rename = "line-width", skip_serializing_if = "Option::is_none")]
    pub line_width: Option<f64>,
    #[serde(default, rename = "corner-radius", skip_serializing_if = "Option::is_none")]
    pub corner_radius: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rotate: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<bool>,
    // Text in shapes
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, rename = "font-size", skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub valign: Option<String>,
    // Image in shapes
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub image: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum SlideBackground {
    Solid { color: String },
    Image { image: String },
    Gradient {
        gradient: GradientDef,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GradientDef {
    pub from: String,
    pub to: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub angle: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideTransition {
    #[serde(rename = "type")]
    pub transition_type: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub duration: Option<f64>,
}
