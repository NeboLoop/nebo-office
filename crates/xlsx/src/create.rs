use anyhow::Result;
use nebo_office_core::relationships::{
    build_content_types, RelationshipManager, OOXML_IMAGE_EXTENSIONS, REL_CORE_PROPS,
};
use nebo_office_core::zip_utils::create_zip;
use nebo_spec::xlsx_types::*;
use std::collections::HashMap;
use std::io::{Seek, Write};

// SpreadsheetML relationship types
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_THEME: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

// SpreadsheetML content types
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const CT_CORE_PROPS: &str =
    "application/vnd.openxmlformats-package.core-properties+xml";

// SpreadsheetML namespaces
const NS_SPREADSHEET: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_DC: &str = "http://purl.org/dc/elements/1.1/";
const NS_DCTERMS: &str = "http://purl.org/dc/terms/";
const NS_CP: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const NS_XSI: &str = "http://www.w3.org/2001/XMLSchema-instance";

/// Internal style record for the cellXfs table.
#[derive(Debug, Clone, Hash, PartialEq, Eq)]
struct CellStyle {
    num_fmt_id: u32,
    font_idx: u32,
    fill_idx: u32,
    apply_font: bool,
    apply_fill: bool,
    apply_num_fmt: bool,
    apply_alignment: bool,
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap_text: bool,
}

/// Internal font record.
#[derive(Debug, Clone, Hash, PartialEq, Eq)]
struct FontRecord {
    name: String,
    size_half_points: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    color: Option<String>,
}

/// Internal fill record.
#[derive(Debug, Clone, Hash, PartialEq, Eq)]
struct FillRecord {
    fg_color: Option<String>,
}

struct XlsxBuilder {
    shared_strings: Vec<String>,
    shared_string_map: HashMap<String, u32>,
    fonts: Vec<FontRecord>,
    font_map: HashMap<FontRecord, u32>,
    fills: Vec<FillRecord>,
    fill_map: HashMap<FillRecord, u32>,
    num_fmts: Vec<(u32, String)>,
    num_fmt_map: HashMap<String, u32>,
    xfs: Vec<CellStyle>,
    xf_map: HashMap<CellStyle, u32>,
    next_num_fmt_id: u32,
}

impl XlsxBuilder {
    fn new(spec: &XlsxSpec) -> Self {
        let default_font = spec
            .styles
            .as_ref()
            .and_then(|s| s.font.clone())
            .unwrap_or_else(|| "Calibri".to_string());
        let default_size = spec
            .styles
            .as_ref()
            .and_then(|s| s.size)
            .unwrap_or(11.0);

        // Initialize with default font (index 0)
        let default_font_rec = FontRecord {
            name: default_font,
            size_half_points: (default_size * 2.0) as u32,
            bold: false,
            italic: false,
            underline: false,
            color: None,
        };

        let mut font_map = HashMap::new();
        font_map.insert(default_font_rec.clone(), 0);

        // Fills: index 0 = none, index 1 = gray125 (required by Excel)
        let fill_none = FillRecord { fg_color: None };
        let fill_gray = FillRecord {
            fg_color: Some("__gray125__".to_string()),
        };
        let mut fill_map = HashMap::new();
        fill_map.insert(fill_none.clone(), 0);
        fill_map.insert(fill_gray.clone(), 1);

        // Default xf (index 0) — the "Normal" style
        let default_xf = CellStyle {
            num_fmt_id: 0,
            font_idx: 0,
            fill_idx: 0,
            apply_font: false,
            apply_fill: false,
            apply_num_fmt: false,
            apply_alignment: false,
            horizontal: None,
            vertical: None,
            wrap_text: false,
        };
        let mut xf_map = HashMap::new();
        xf_map.insert(default_xf.clone(), 0);

        Self {
            shared_strings: Vec::new(),
            shared_string_map: HashMap::new(),
            fonts: vec![default_font_rec],
            font_map,
            fills: vec![fill_none, fill_gray],
            fill_map,
            num_fmts: Vec::new(),
            num_fmt_map: HashMap::new(),
            xfs: vec![default_xf],
            xf_map,
            next_num_fmt_id: 164, // custom formats start at 164
        }
    }

    fn intern_string(&mut self, s: &str) -> u32 {
        if let Some(&idx) = self.shared_string_map.get(s) {
            return idx;
        }
        let idx = self.shared_strings.len() as u32;
        self.shared_strings.push(s.to_string());
        self.shared_string_map.insert(s.to_string(), idx);
        idx
    }

    fn get_or_add_font(&mut self, rec: FontRecord) -> u32 {
        if let Some(&idx) = self.font_map.get(&rec) {
            return idx;
        }
        let idx = self.fonts.len() as u32;
        self.fonts.push(rec.clone());
        self.font_map.insert(rec, idx);
        idx
    }

    fn get_or_add_fill(&mut self, rec: FillRecord) -> u32 {
        if let Some(&idx) = self.fill_map.get(&rec) {
            return idx;
        }
        let idx = self.fills.len() as u32;
        self.fills.push(rec.clone());
        self.fill_map.insert(rec, idx);
        idx
    }

    fn get_or_add_num_fmt(&mut self, format_str: &str) -> u32 {
        // Check built-in formats first
        if let Some(id) = builtin_num_fmt(format_str) {
            return id;
        }
        if let Some(&id) = self.num_fmt_map.get(format_str) {
            return id;
        }
        let id = self.next_num_fmt_id;
        self.next_num_fmt_id += 1;
        self.num_fmts.push((id, format_str.to_string()));
        self.num_fmt_map.insert(format_str.to_string(), id);
        id
    }

    fn get_or_add_xf(&mut self, style: CellStyle) -> u32 {
        if let Some(&idx) = self.xf_map.get(&style) {
            return idx;
        }
        let idx = self.xfs.len() as u32;
        self.xfs.push(style.clone());
        self.xf_map.insert(style, idx);
        idx
    }

    /// Resolve cell style from row defaults + cell overrides. Returns xf index.
    fn resolve_cell_style(
        &mut self,
        row: &Row,
        cell: &CellValue,
        col_def: Option<&ColumnDef>,
        default_font: &str,
        default_size: f64,
    ) -> u32 {
        let (cell_bold, cell_italic, cell_underline, cell_font, cell_size, cell_color,
            cell_shading, cell_format, cell_align, cell_valign, cell_wrap) = match cell {
            CellValue::Rich(rc) => (
                rc.bold,
                rc.italic,
                rc.underline,
                rc.font.as_deref(),
                rc.size,
                rc.color.as_deref(),
                rc.shading.as_deref(),
                rc.format.as_deref(),
                rc.align.as_deref(),
                rc.valign.as_deref(),
                rc.wrap.unwrap_or(false),
            ),
            _ => (None, None, None, None, None, None, None, None, None, None, false),
        };

        let bold = cell_bold.or(row.bold).unwrap_or(false);
        let italic = cell_italic.or(row.italic).unwrap_or(false);
        let underline = cell_underline.unwrap_or(false);
        let font_name = cell_font
            .or(row.font.as_deref())
            .unwrap_or(default_font);
        let font_size = cell_size.or(row.size).unwrap_or(default_size);
        let color = cell_color.or(row.color.as_deref());
        let shading = cell_shading.or(row.shading.as_deref());
        let num_format = cell_format
            .or(col_def.and_then(|c| c.format.as_deref()));

        // Determine if we need a non-default font
        let needs_font = bold
            || italic
            || underline
            || font_name != default_font
            || (font_size - default_size).abs() > 0.01
            || color.is_some();

        let font_idx = if needs_font {
            self.get_or_add_font(FontRecord {
                name: font_name.to_string(),
                size_half_points: (font_size * 2.0) as u32,
                bold,
                italic,
                underline,
                color: color.map(|s| s.to_string()),
            })
        } else {
            0
        };

        let fill_idx = if let Some(shade) = shading {
            self.get_or_add_fill(FillRecord {
                fg_color: Some(shade.to_string()),
            })
        } else {
            0
        };

        let num_fmt_id = if let Some(fmt) = num_format {
            self.get_or_add_num_fmt(fmt)
        } else {
            0
        };

        let has_alignment = cell_align.is_some() || cell_valign.is_some() || cell_wrap;

        let style = CellStyle {
            num_fmt_id,
            font_idx,
            fill_idx,
            apply_font: font_idx > 0,
            apply_fill: fill_idx > 0,
            apply_num_fmt: num_fmt_id > 0,
            apply_alignment: has_alignment,
            horizontal: cell_align.map(|s| s.to_string()),
            vertical: cell_valign.map(|s| s.to_string()),
            wrap_text: cell_wrap,
        };

        if style.font_idx == 0
            && style.fill_idx == 0
            && style.num_fmt_id == 0
            && !style.apply_alignment
        {
            return 0;
        }

        self.get_or_add_xf(style)
    }
}

/// Create an XLSX file from a spec.
pub fn create_xlsx<W: Write + Seek>(
    spec: &XlsxSpec,
    writer: W,
    _assets_dir: Option<&std::path::Path>,
) -> Result<()> {
    let default_font = spec
        .styles
        .as_ref()
        .and_then(|s| s.font.clone())
        .unwrap_or_else(|| "Calibri".to_string());
    let default_size = spec
        .styles
        .as_ref()
        .and_then(|s| s.size)
        .unwrap_or(11.0);

    let mut builder = XlsxBuilder::new(spec);

    // Pre-pass: build all sheet XML while populating shared strings/styles
    let mut sheet_xmls: Vec<String> = Vec::new();

    for sheet in &spec.sheets {
        let xml = build_sheet_xml(sheet, &mut builder, &default_font, default_size);
        sheet_xmls.push(xml);
    }

    // Build workbook relationships
    let mut wb_rels = RelationshipManager::new();
    let mut content_types: HashMap<String, String> = HashMap::new();

    // Worksheets
    for (i, _sheet) in spec.sheets.iter().enumerate() {
        wb_rels.add(REL_WORKSHEET, &format!("worksheets/sheet{}.xml", i + 1));
    }

    // Styles
    wb_rels.add(REL_STYLES, "styles.xml");
    content_types.insert("/xl/styles.xml".to_string(), CT_STYLES.to_string());

    // Shared strings
    if !builder.shared_strings.is_empty() {
        wb_rels.add(REL_SHARED_STRINGS, "sharedStrings.xml");
        content_types.insert(
            "/xl/sharedStrings.xml".to_string(),
            CT_SHARED_STRINGS.to_string(),
        );
    }

    // Theme
    wb_rels.add(REL_THEME, "theme/theme1.xml");
    content_types.insert("/xl/theme/theme1.xml".to_string(), CT_THEME.to_string());

    // Content types for sheets
    content_types.insert("/xl/workbook.xml".to_string(), CT_WORKBOOK.to_string());
    for i in 0..spec.sheets.len() {
        content_types.insert(
            format!("/xl/worksheets/sheet{}.xml", i + 1),
            CT_WORKSHEET.to_string(),
        );
    }

    // Package rels
    let mut pkg_rels = RelationshipManager::new();
    pkg_rels.add(REL_OFFICE_DOCUMENT, "xl/workbook.xml");

    let has_metadata = spec.metadata.is_some();
    if has_metadata {
        pkg_rels.add(REL_CORE_PROPS, "docProps/core.xml");
        content_types.insert(
            "/docProps/core.xml".to_string(),
            CT_CORE_PROPS.to_string(),
        );
    }

    // Build XML strings
    let workbook_xml = build_workbook_xml(spec);
    let styles_xml = build_styles_xml(&builder, &default_font, default_size);
    let shared_strings_xml = build_shared_strings_xml(&builder);
    let theme_xml = build_theme_xml();
    let content_types_xml = build_content_types(&content_types, OOXML_IMAGE_EXTENSIONS);
    let pkg_rels_xml = pkg_rels.to_xml();
    let wb_rels_xml = wb_rels.to_xml();
    let core_xml = if has_metadata {
        Some(build_core_xml(spec))
    } else {
        None
    };

    // Assemble ZIP
    let mut files: Vec<(&str, &[u8])> = Vec::new();
    files.push(("[Content_Types].xml", content_types_xml.as_bytes()));
    files.push(("_rels/.rels", pkg_rels_xml.as_bytes()));
    files.push(("xl/workbook.xml", workbook_xml.as_bytes()));
    files.push(("xl/_rels/workbook.xml.rels", wb_rels_xml.as_bytes()));
    files.push(("xl/styles.xml", styles_xml.as_bytes()));
    files.push(("xl/theme/theme1.xml", theme_xml.as_bytes()));

    if !builder.shared_strings.is_empty() {
        files.push(("xl/sharedStrings.xml", shared_strings_xml.as_bytes()));
    }

    // Owned sheet paths for borrow checker
    let sheet_paths: Vec<String> = (0..spec.sheets.len())
        .map(|i| format!("xl/worksheets/sheet{}.xml", i + 1))
        .collect();
    for (i, xml) in sheet_xmls.iter().enumerate() {
        files.push((&sheet_paths[i], xml.as_bytes()));
    }

    if let Some(ref core) = core_xml {
        files.push(("docProps/core.xml", core.as_bytes()));
    }

    create_zip(writer, &files)?;
    Ok(())
}

fn build_workbook_xml(spec: &XlsxSpec) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<workbook xmlns="{NS_SPREADSHEET}" xmlns:r="{NS_RELATIONSHIPS}">"#
    ));
    xml.push_str("<sheets>");
    for (i, sheet) in spec.sheets.iter().enumerate() {
        let name = xml_escape(&sheet.name);
        xml.push_str(&format!(
            r#"<sheet name="{name}" sheetId="{}" r:id="rId{}"/>"#,
            i + 1,
            i + 1
        ));
    }
    xml.push_str("</sheets>");

    // Named ranges (definedNames)
    if !spec.named_ranges.is_empty() {
        xml.push_str("<definedNames>");
        for nr in &spec.named_ranges {
            let name = xml_escape(&nr.name);
            let range = xml_escape(&nr.range);
            xml.push_str(&format!(r#"<definedName name="{name}">{range}</definedName>"#));
        }
        xml.push_str("</definedNames>");
    }

    xml.push_str("</workbook>");
    xml
}

fn build_sheet_xml(
    sheet: &Sheet,
    builder: &mut XlsxBuilder,
    default_font: &str,
    default_size: f64,
) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<worksheet xmlns="{NS_SPREADSHEET}" xmlns:r="{NS_RELATIONSHIPS}">"#));

    // Sheet views (freeze panes)
    if let Some(ref freeze) = sheet.freeze {
        xml.push_str("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">");
        write_freeze_pane(&mut xml, freeze);
        xml.push_str("</sheetView></sheetViews>");
    }

    // Column definitions
    if !sheet.columns.is_empty() {
        xml.push_str("<cols>");
        for (i, col) in sheet.columns.iter().enumerate() {
            let col_num = i + 1;
            let width = col.width.unwrap_or(10.0);
            let hidden = if col.hidden.unwrap_or(false) {
                " hidden=\"1\""
            } else {
                ""
            };
            xml.push_str(&format!(
                r#"<col min="{col_num}" max="{col_num}" width="{width}" customWidth="1"{hidden}/>"#
            ));
        }
        xml.push_str("</cols>");
    }

    // Sheet data (rows and cells)
    xml.push_str("<sheetData>");
    for (row_idx, row) in sheet.rows.iter().enumerate() {
        let row_num = row_idx + 1;
        let mut row_attrs = format!(r#" r="{row_num}""#);
        if let Some(h) = row.height {
            row_attrs.push_str(&format!(r#" ht="{h}" customHeight="1""#));
        }
        xml.push_str(&format!("<row{row_attrs}>"));

        for (col_idx, cell) in row.cells.iter().enumerate() {
            let col_letter = col_to_letter(col_idx);
            let cell_ref = format!("{col_letter}{row_num}");
            let col_def = sheet.columns.get(col_idx);

            let xf_idx = builder.resolve_cell_style(row, cell, col_def, default_font, default_size);

            write_cell(&mut xml, &cell_ref, cell, xf_idx, builder);
        }

        xml.push_str("</row>");
    }
    xml.push_str("</sheetData>");

    // Merged cells
    if !sheet.merged.is_empty() {
        xml.push_str(&format!(
            r#"<mergeCells count="{}">"#,
            sheet.merged.len()
        ));
        for range in &sheet.merged {
            xml.push_str(&format!(r#"<mergeCell ref="{range}"/>"#));
        }
        xml.push_str("</mergeCells>");
    }

    // Conditional formatting
    for cf in &sheet.conditional {
        write_conditional_formatting(&mut xml, cf);
    }

    // Data validation
    if !sheet.validations.is_empty() {
        xml.push_str(&format!(
            r#"<dataValidations count="{}">"#,
            sheet.validations.len()
        ));
        for dv in &sheet.validations {
            write_data_validation(&mut xml, dv);
        }
        xml.push_str("</dataValidations>");
    }

    // Auto-filter
    if let Some(ref af) = sheet.autofilter {
        xml.push_str(&format!(r#"<autoFilter ref="{}"/>"#, af.range));
    }

    // Print setup
    if let Some(ref print) = sheet.print {
        write_print_setup(&mut xml, print);
    }

    xml.push_str("</worksheet>");
    xml
}

fn write_cell(
    xml: &mut String,
    cell_ref: &str,
    cell: &CellValue,
    xf_idx: u32,
    builder: &mut XlsxBuilder,
) {
    let style_attr = if xf_idx > 0 {
        format!(r#" s="{xf_idx}""#)
    } else {
        String::new()
    };

    match cell {
        CellValue::String(s) => {
            if s.is_empty() {
                return; // skip empty cells
            }
            // Strip markdown bold/italic for the cell value
            let clean = strip_markdown(s);
            let ssi = builder.intern_string(&clean);
            xml.push_str(&format!(
                r#"<c r="{cell_ref}" t="s"{style_attr}><v>{ssi}</v></c>"#
            ));
        }
        CellValue::Number(n) => {
            xml.push_str(&format!(
                r#"<c r="{cell_ref}"{style_attr}><v>{n}</v></c>"#
            ));
        }
        CellValue::Bool(b) => {
            let v = if *b { 1 } else { 0 };
            xml.push_str(&format!(
                r#"<c r="{cell_ref}" t="b"{style_attr}><v>{v}</v></c>"#
            ));
        }
        CellValue::Null => {
            // skip null cells
        }
        CellValue::Rich(rc) => {
            if let Some(ref formula) = rc.formula {
                xml.push_str(&format!(
                    r#"<c r="{cell_ref}"{style_attr}><f>{}</f></c>"#,
                    xml_escape(formula)
                ));
            } else if let Some(ref val) = rc.value {
                match val {
                    serde_json::Value::String(s) => {
                        let clean = strip_markdown(s);
                        let ssi = builder.intern_string(&clean);
                        xml.push_str(&format!(
                            r#"<c r="{cell_ref}" t="s"{style_attr}><v>{ssi}</v></c>"#
                        ));
                    }
                    serde_json::Value::Number(n) => {
                        xml.push_str(&format!(
                            r#"<c r="{cell_ref}"{style_attr}><v>{n}</v></c>"#
                        ));
                    }
                    serde_json::Value::Bool(b) => {
                        let v = if *b { 1 } else { 0 };
                        xml.push_str(&format!(
                            r#"<c r="{cell_ref}" t="b"{style_attr}><v>{v}</v></c>"#
                        ));
                    }
                    _ => {}
                }
            }
        }
    }
}

fn write_freeze_pane(xml: &mut String, freeze: &FreezePane) {
    let row = freeze.row.unwrap_or(0);
    let col = freeze.col.unwrap_or(0);

    if row == 0 && col == 0 {
        return;
    }

    let top_left = format!("{}{}", col_to_letter(col as usize), row + 1);

    let active_pane = match (row > 0, col > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        (false, true) => "topRight",
        (false, false) => return,
    };

    xml.push_str(&format!(
        r#"<pane xSplit="{col}" ySplit="{row}" topLeftCell="{top_left}" activePane="{active_pane}" state="frozen"/>"#
    ));
    xml.push_str(&format!(
        r#"<selection pane="{active_pane}" activeCell="{top_left}" sqref="{top_left}"/>"#
    ));
}

fn write_conditional_formatting(xml: &mut String, cf: &ConditionalFormat) {
    xml.push_str(&format!(r#"<conditionalFormatting sqref="{}">"#, cf.range));

    let operator = match cf.rule.as_str() {
        "greater-than" => "greaterThan",
        "less-than" => "lessThan",
        "equal" => "equal",
        "not-equal" => "notEqual",
        "between" => "between",
        other => other,
    };

    xml.push_str(&format!(
        r#"<cfRule type="cellIs" operator="{operator}" priority="1""#
    ));

    // DXF style index (simplified — inline style)
    xml.push_str(">");

    if let Some(ref val) = cf.value {
        xml.push_str(&format!("<formula>{val}</formula>"));
    }

    xml.push_str("</cfRule>");
    xml.push_str("</conditionalFormatting>");
}

fn write_data_validation(xml: &mut String, dv: &DataValidation) {
    match dv.validation_type.as_str() {
        "list" => {
            let formula = format!("\"{}\"", dv.values.join(","));
            xml.push_str(&format!(
                r#"<dataValidation type="list" allowBlank="1" showDropDown="0" sqref="{}"><formula1>{}</formula1></dataValidation>"#,
                dv.range,
                xml_escape(&formula)
            ));
        }
        "whole" | "decimal" => {
            let min = dv.min.map(|n| n.to_string()).unwrap_or_default();
            let max = dv.max.map(|n| n.to_string()).unwrap_or_default();
            xml.push_str(&format!(
                r#"<dataValidation type="{}" operator="between" allowBlank="1" sqref="{}"><formula1>{min}</formula1><formula2>{max}</formula2></dataValidation>"#,
                dv.validation_type, dv.range
            ));
        }
        _ => {}
    }
}

fn write_print_setup(xml: &mut String, print: &PrintSetup) {
    if let Some(ref orientation) = print.orientation {
        xml.push_str(&format!(
            r#"<pageSetup orientation="{orientation}"/>"#
        ));
    }
    if print.fit_to_page == Some(true) {
        xml.push_str(r#"<sheetFormatPr defaultRowHeight="15" />"#);
    }
}

fn build_styles_xml(builder: &XlsxBuilder, default_font: &str, default_size: f64) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<styleSheet xmlns="{NS_SPREADSHEET}">"#));

    // Number formats
    if !builder.num_fmts.is_empty() {
        xml.push_str(&format!(
            r#"<numFmts count="{}">"#,
            builder.num_fmts.len()
        ));
        for (id, fmt_str) in &builder.num_fmts {
            xml.push_str(&format!(
                r#"<numFmt numFmtId="{id}" formatCode="{}"/>"#,
                xml_escape(fmt_str)
            ));
        }
        xml.push_str("</numFmts>");
    }

    // Fonts
    xml.push_str(&format!(r#"<fonts count="{}">"#, builder.fonts.len()));
    for font in &builder.fonts {
        xml.push_str("<font>");
        if font.bold {
            xml.push_str("<b/>");
        }
        if font.italic {
            xml.push_str("<i/>");
        }
        if font.underline {
            xml.push_str("<u/>");
        }
        let sz = font.size_half_points as f64 / 2.0;
        xml.push_str(&format!(r#"<sz val="{sz}"/>"#));
        if let Some(ref color) = font.color {
            xml.push_str(&format!(r#"<color rgb="FF{color}"/>"#));
        }
        xml.push_str(&format!(r#"<name val="{}"/>"#, xml_escape(&font.name)));
        xml.push_str("</font>");
    }
    xml.push_str("</fonts>");

    // Fills
    xml.push_str(&format!(r#"<fills count="{}">"#, builder.fills.len()));
    for fill in &builder.fills {
        match fill.fg_color.as_deref() {
            None => {
                xml.push_str(r#"<fill><patternFill patternType="none"/></fill>"#);
            }
            Some("__gray125__") => {
                xml.push_str(r#"<fill><patternFill patternType="gray125"/></fill>"#);
            }
            Some(color) => {
                xml.push_str(&format!(
                    r#"<fill><patternFill patternType="solid"><fgColor rgb="FF{color}"/></patternFill></fill>"#
                ));
            }
        }
    }
    xml.push_str("</fills>");

    // Borders (just one empty default)
    xml.push_str(r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#);

    // Cell style xfs (base styles)
    xml.push_str(r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#);

    // Cell xfs (the actual cell format records)
    xml.push_str(&format!(r#"<cellXfs count="{}">"#, builder.xfs.len()));
    for xf in &builder.xfs {
        let mut attrs = format!(
            r#"numFmtId="{}" fontId="{}" fillId="{}" borderId="0""#,
            xf.num_fmt_id, xf.font_idx, xf.fill_idx
        );
        if xf.apply_font {
            attrs.push_str(r#" applyFont="1""#);
        }
        if xf.apply_fill {
            attrs.push_str(r#" applyFill="1""#);
        }
        if xf.apply_num_fmt {
            attrs.push_str(r#" applyNumberFormat="1""#);
        }
        if xf.apply_alignment {
            attrs.push_str(r#" applyAlignment="1""#);
            xml.push_str(&format!("<xf {attrs}>"));
            let mut align_attrs = String::new();
            if let Some(ref h) = xf.horizontal {
                align_attrs.push_str(&format!(r#" horizontal="{h}""#));
            }
            if let Some(ref v) = xf.vertical {
                align_attrs.push_str(&format!(r#" vertical="{v}""#));
            }
            if xf.wrap_text {
                align_attrs.push_str(r#" wrapText="1""#);
            }
            xml.push_str(&format!("<alignment{align_attrs}/>"));
            xml.push_str("</xf>");
        } else {
            xml.push_str(&format!("<xf {attrs}/>"));
        }
    }
    xml.push_str("</cellXfs>");

    // Cell styles
    xml.push_str(r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#);

    xml.push_str("</styleSheet>");
    xml
}

fn build_shared_strings_xml(builder: &XlsxBuilder) -> String {
    let count = builder.shared_strings.len();
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<sst xmlns="{NS_SPREADSHEET}" count="{count}" uniqueCount="{count}">"#
    ));
    for s in &builder.shared_strings {
        xml.push_str(&format!("<si><t>{}</t></si>", xml_escape(s)));
    }
    xml.push_str("</sst>");
    xml
}

fn build_theme_xml() -> String {
    // Minimal theme — required by Excel for proper rendering
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">"#);
    xml.push_str(r#"<a:themeElements>"#);
    xml.push_str(r#"<a:clrScheme name="Office">"#);
    xml.push_str(r#"<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>"#);
    xml.push_str(r#"<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>"#);
    xml.push_str(r#"<a:dk2><a:srgbClr val="44546A"/></a:dk2>"#);
    xml.push_str(r#"<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>"#);
    xml.push_str(r#"<a:accent1><a:srgbClr val="4472C4"/></a:accent1>"#);
    xml.push_str(r#"<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>"#);
    xml.push_str(r#"<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>"#);
    xml.push_str(r#"<a:accent4><a:srgbClr val="FFC000"/></a:accent4>"#);
    xml.push_str(r#"<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>"#);
    xml.push_str(r#"<a:accent6><a:srgbClr val="70AD47"/></a:accent6>"#);
    xml.push_str(r#"<a:hlink><a:srgbClr val="0563C1"/></a:hlink>"#);
    xml.push_str(r#"<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>"#);
    xml.push_str(r#"</a:clrScheme>"#);
    xml.push_str(r#"<a:fontScheme name="Office">"#);
    xml.push_str(r#"<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>"#);
    xml.push_str(r#"<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>"#);
    xml.push_str(r#"</a:fontScheme>"#);
    xml.push_str(r#"<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>"#);
    xml.push_str(r#"</a:themeElements></a:theme>"#);
    xml
}

fn build_core_xml(spec: &XlsxSpec) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<cp:coreProperties xmlns:cp="{NS_CP}" xmlns:dc="{NS_DC}" xmlns:dcterms="{NS_DCTERMS}" xmlns:xsi="{NS_XSI}">"#
    ));

    if let Some(ref meta) = spec.metadata {
        if let Some(ref title) = meta.title {
            xml.push_str(&format!("<dc:title>{}</dc:title>", xml_escape(title)));
        }
        if let Some(ref creator) = meta.creator {
            xml.push_str(&format!(
                "<dc:creator>{}</dc:creator>",
                xml_escape(creator)
            ));
        }
        if let Some(ref subject) = meta.subject {
            xml.push_str(&format!(
                "<dc:subject>{}</dc:subject>",
                xml_escape(subject)
            ));
        }
        if let Some(ref desc) = meta.description {
            xml.push_str(&format!(
                "<dc:description>{}</dc:description>",
                xml_escape(desc)
            ));
        }
    }

    let now = chrono::Utc::now().format("%Y-%m-%dT%H:%M:%SZ");
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

// --- Utilities ---

fn col_to_letter(col: usize) -> String {
    let mut result = String::new();
    let mut c = col;
    loop {
        result.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    result
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn strip_markdown(s: &str) -> String {
    s.replace("**", "").replace("__", "")
}

/// Map common Excel format strings to built-in numFmtId values.
fn builtin_num_fmt(fmt: &str) -> Option<u32> {
    match fmt {
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "mm-dd-yy" => Some(14),
        "d-mmm-yy" => Some(15),
        "h:mm" => Some(20),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(701), "ZZ");
    }

    #[test]
    fn test_strip_markdown() {
        assert_eq!(strip_markdown("**Total**"), "Total");
        assert_eq!(strip_markdown("Normal text"), "Normal text");
    }

    #[test]
    fn test_create_basic_xlsx() {
        let spec = XlsxSpec {
            version: 1,
            metadata: Some(XlsxMetadata {
                title: Some("Test".to_string()),
                creator: Some("Test User".to_string()),
                subject: None,
                description: None,
            }),
            styles: None,
            sheets: vec![Sheet {
                name: "Sheet1".to_string(),
                columns: vec![
                    ColumnDef { width: Some(20.0), format: None, hidden: None },
                    ColumnDef { width: Some(15.0), format: Some("$#,##0".to_string()), hidden: None },
                ],
                rows: vec![
                    Row {
                        cells: vec![
                            CellValue::String("Region".to_string()),
                            CellValue::String("Revenue".to_string()),
                        ],
                        bold: Some(true),
                        italic: None,
                        shading: Some("4472C4".to_string()),
                        color: Some("FFFFFF".to_string()),
                        font: None,
                        size: None,
                        height: None,
                    },
                    Row {
                        cells: vec![
                            CellValue::String("North America".to_string()),
                            CellValue::Number(1250000.0),
                        ],
                        bold: None,
                        italic: None,
                        shading: None,
                        color: None,
                        font: None,
                        size: None,
                        height: None,
                    },
                ],
                merged: vec![],
                freeze: Some(FreezePane { row: Some(1), col: None }),
                conditional: vec![],
                charts: vec![],
                images: vec![],
                autofilter: None,
                validations: vec![],
                print: None,
            }],
            named_ranges: vec![],
        };

        let mut buf = std::io::Cursor::new(Vec::new());
        create_xlsx(&spec, &mut buf, None).unwrap();
        let data = buf.into_inner();
        assert!(data.len() > 100, "XLSX should be non-trivial size");

        // Verify it's a valid ZIP
        let cursor = std::io::Cursor::new(&data);
        let archive = zip::ZipArchive::new(cursor).unwrap();
        let names: Vec<_> = archive.file_names().collect();
        assert!(names.contains(&"[Content_Types].xml"));
        assert!(names.contains(&"xl/workbook.xml"));
        assert!(names.contains(&"xl/worksheets/sheet1.xml"));
        assert!(names.contains(&"xl/styles.xml"));
        assert!(names.contains(&"xl/sharedStrings.xml"));
    }
}
