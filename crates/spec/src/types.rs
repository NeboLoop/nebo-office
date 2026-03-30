use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// Root document spec.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DocSpec {
    pub version: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page: Option<PageSetup>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub styles: Option<Styles>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headers: Option<HeaderFooterSet>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footers: Option<HeaderFooterSet>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnotes: Option<HashMap<String, String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments: Option<HashMap<String, Comment>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub metadata: Option<Metadata>,
    pub body: Vec<Block>,
}

// --- Page Setup ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageSetup {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<PageSize>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<Orientation>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin: Option<Margin>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum PageSize {
    Named(String),
    Custom { width: f64, height: f64 },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Orientation {
    Portrait,
    Landscape,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Margin {
    Uniform(f64),
    Custom {
        #[serde(default)]
        top: Option<f64>,
        #[serde(default)]
        bottom: Option<f64>,
        #[serde(default)]
        left: Option<f64>,
        #[serde(default)]
        right: Option<f64>,
    },
}

// --- Styles ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Styles {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headings: Option<HeadingStyles>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom: Option<HashMap<String, CustomStyle>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeadingStyles {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h1: Option<HeadingLevel>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h2: Option<HeadingLevel>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h3: Option<HeadingLevel>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h4: Option<HeadingLevel>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h5: Option<HeadingLevel>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h6: Option<HeadingLevel>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeadingLevel {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<Indent>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Spacing>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
}

// --- Block Elements ---

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Block {
    Heading {
        heading: u8,
        text: TextContent,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        id: Option<String>,
    },
    Paragraph {
        paragraph: ParagraphContent,
    },
    Bullets {
        bullets: Vec<ListItem>,
    },
    Numbered {
        numbered: Vec<ListItem>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        restart: Option<bool>,
    },
    Table {
        table: TableContent,
        #[serde(default, rename = "header-rows", skip_serializing_if = "Option::is_none")]
        header_rows: Option<u32>,
    },
    Image {
        image: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        width: Option<f64>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        height: Option<f64>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        alt: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        align: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        caption: Option<String>,
        #[serde(default, rename = "image-data", skip_serializing_if = "Option::is_none")]
        image_data: Option<String>,
    },
    PageBreak {
        #[serde(rename = "page-break")]
        page_break: bool,
    },
    Toc {
        toc: TocContent,
    },
    SectionBreak {
        #[serde(rename = "section-break")]
        section_break: SectionBreakConfig,
    },
    Bookmark {
        bookmark: String,
    },
    Raw {
        #[serde(rename = "_raw")]
        raw: String,
    },
}

// --- Paragraph ---

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ParagraphContent {
    Simple(String),
    Full(ParagraphFull),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ParagraphFull {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub runs: Option<Vec<Run>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Spacing>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<Indent>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inserted: Option<ChangeInfo>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub deleted: Option<ChangeInfo>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum TextContent {
    Simple(String),
    Full { text: String },
}

impl TextContent {
    pub fn as_str(&self) -> &str {
        match self {
            TextContent::Simple(s) => s,
            TextContent::Full { text } => text,
        }
    }
}

// --- Runs ---

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Run {
    Text(TextRun),
    Tab(TabRun),
    Field(FieldRun),
    FootnoteRef(FootnoteRun),
    Delete(DeleteRun),
    Insert(InsertRun),
    CommentStart(CommentStartRun),
    CommentEnd(CommentEndRun),
    Break(BreakRun),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextRun {
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub superscript: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subscript: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link: Option<String>,
    #[serde(default, rename = "all-caps", skip_serializing_if = "Option::is_none")]
    pub all_caps: Option<bool>,
    #[serde(default, rename = "small-caps", skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TabRun {
    pub tab: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FieldRun {
    pub field: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FootnoteRun {
    pub footnote: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DeleteRun {
    pub delete: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InsertRun {
    pub insert: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentStartRun {
    #[serde(rename = "comment-start")]
    pub comment_start: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentEndRun {
    #[serde(rename = "comment-end")]
    pub comment_end: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BreakRun {
    #[serde(rename = "break")]
    pub break_type: String,
}

// --- Lists ---

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ListItem {
    Simple(String),
    WithChildren {
        text: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        children: Option<Vec<ListItem>>,
    },
}

impl ListItem {
    pub fn text(&self) -> &str {
        match self {
            ListItem::Simple(s) => s,
            ListItem::WithChildren { text, .. } => text,
        }
    }

    pub fn children(&self) -> Option<&[ListItem]> {
        match self {
            ListItem::Simple(_) => None,
            ListItem::WithChildren { children, .. } => children.as_deref(),
        }
    }
}

// --- Tables ---

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum TableContent {
    Simple(Vec<Vec<String>>),
    Full(TableFull),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableFull {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<Vec<ColumnDef>>,
    #[serde(default, rename = "header-rows", skip_serializing_if = "Option::is_none")]
    pub header_rows: Option<u32>,
    pub rows: Vec<TableRow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnDef {
    pub width: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub runs: Option<Vec<Run>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub body: Option<Vec<Block>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub colspan: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rowspan: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub valign: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
}

// --- TOC ---

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum TocContent {
    Simple(bool),
    Full {
        #[serde(default, skip_serializing_if = "Option::is_none")]
        title: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        depth: Option<u8>,
    },
}

// --- Section Break ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SectionBreakConfig {
    #[serde(default, rename = "type", skip_serializing_if = "Option::is_none")]
    pub break_type: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<u32>,
    #[serde(default, rename = "column-gap", skip_serializing_if = "Option::is_none")]
    pub column_gap: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page: Option<SectionPageSetup>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub valign: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SectionPageSetup {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<Orientation>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<PageSize>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin: Option<Margin>,
}

// --- Spacing/Indent ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Spacing {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub before: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub after: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Indent {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_line: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hanging: Option<f64>,
}

// --- Headers/Footers ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeaderFooterSet {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default: Option<Vec<Block>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first: Option<Vec<Block>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even: Option<Vec<Block>>,
}

// --- Comments ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub replies: Option<Vec<CommentReply>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentReply {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    pub text: String,
}

// --- Tracked Changes ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChangeInfo {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
}

// --- Metadata ---

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Metadata {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub creator: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keywords: Option<Vec<String>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category: Option<String>,
}
