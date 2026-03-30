use std::io::Cursor;

use nebo_docx::{create, unpack};
use nebo_spec::*;

fn create_and_unpack(spec: &DocSpec) -> DocSpec {
    let mut buf = Cursor::new(Vec::new());
    create::create_docx(spec, &mut buf, None).expect("create failed");
    let data = buf.into_inner();
    assert!(!data.is_empty(), "DOCX output is empty");

    // Verify it's a valid ZIP
    assert_eq!(&data[0..2], b"PK", "not a valid ZIP file");

    let cursor = Cursor::new(data);
    unpack::unpack_docx(cursor, None, false).expect("unpack failed")
}

#[test]
fn test_simple_paragraph() {
    let spec = DocSpec {
        version: 1,
        page: None,
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![Block::Paragraph {
            paragraph: ParagraphContent::Simple("Hello World".into()),
        }],
    };

    let result = create_and_unpack(&spec);
    assert_eq!(result.version, 1);
    assert!(!result.body.is_empty());
}

#[test]
fn test_headings() {
    let spec = DocSpec {
        version: 1,
        page: None,
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![
            Block::Heading {
                heading: 1,
                text: TextContent::Simple("Title".into()),
                id: None,
            },
            Block::Heading {
                heading: 2,
                text: TextContent::Simple("Subtitle".into()),
                id: None,
            },
            Block::Paragraph {
                paragraph: ParagraphContent::Simple("Body text".into()),
            },
        ],
    };

    let result = create_and_unpack(&spec);
    assert!(result.body.len() >= 3);

    // First block should be heading 1
    match &result.body[0] {
        Block::Heading { heading, text, .. } => {
            assert_eq!(*heading, 1);
            assert_eq!(text.as_str(), "Title");
        }
        other => panic!("expected heading, got {:?}", other),
    }
}

#[test]
fn test_bold_italic_markdown() {
    let spec = DocSpec {
        version: 1,
        page: None,
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![Block::Paragraph {
            paragraph: ParagraphContent::Simple(
                "This is **bold** and *italic* text.".into(),
            ),
        }],
    };

    let result = create_and_unpack(&spec);
    // Should round-trip with markdown preserved
    match &result.body[0] {
        Block::Paragraph {
            paragraph: ParagraphContent::Simple(text),
        } => {
            assert!(text.contains("**bold**"));
            assert!(text.contains("*italic*"));
        }
        _ => panic!("expected simple paragraph"),
    }
}

#[test]
fn test_table_round_trip() {
    let spec = DocSpec {
        version: 1,
        page: None,
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![Block::Table {
            table: TableContent::Simple(vec![
                vec!["Name".into(), "Age".into()],
                vec!["Alice".into(), "30".into()],
            ]),
            header_rows: Some(1),
        }],
    };

    let result = create_and_unpack(&spec);
    // Should have a table block
    let has_table = result.body.iter().any(|b| matches!(b, Block::Table { .. }));
    assert!(has_table, "round-trip lost the table");
}

#[test]
fn test_page_break() {
    let spec = DocSpec {
        version: 1,
        page: None,
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![
            Block::Paragraph {
                paragraph: ParagraphContent::Simple("Page 1".into()),
            },
            Block::PageBreak { page_break: true },
            Block::Paragraph {
                paragraph: ParagraphContent::Simple("Page 2".into()),
            },
        ],
    };

    let result = create_and_unpack(&spec);
    let has_break = result
        .body
        .iter()
        .any(|b| matches!(b, Block::PageBreak { .. }));
    assert!(has_break, "round-trip lost the page break");
}

#[test]
fn test_metadata() {
    let spec = DocSpec {
        version: 1,
        page: None,
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: Some(Metadata {
            title: Some("Test Doc".into()),
            creator: Some("nebo-office".into()),
            subject: None,
            description: None,
            keywords: None,
            category: None,
        }),
        body: vec![Block::Paragraph {
            paragraph: ParagraphContent::Simple("Content".into()),
        }],
    };

    let result = create_and_unpack(&spec);
    let meta = result.metadata.expect("metadata lost");
    assert_eq!(meta.title.as_deref(), Some("Test Doc"));
    assert_eq!(meta.creator.as_deref(), Some("nebo-office"));
}

#[test]
fn test_page_setup() {
    let spec = DocSpec {
        version: 1,
        page: Some(PageSetup {
            size: Some(PageSize::Named("a4".into())),
            orientation: None,
            margin: Some(Margin::Uniform(1.0)),
        }),
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![Block::Paragraph {
            paragraph: ParagraphContent::Simple("A4 doc".into()),
        }],
    };

    let result = create_and_unpack(&spec);
    let page = result.page.expect("page setup lost");
    match page.size {
        Some(PageSize::Named(name)) => assert_eq!(name, "a4"),
        other => panic!("expected a4, got {:?}", other),
    }
    match page.margin {
        Some(Margin::Uniform(m)) => assert!((m - 1.0).abs() < 0.01),
        other => panic!("expected uniform 1.0, got {:?}", other),
    }
}

#[test]
fn test_xsd_compliant_output() {
    // Verify critical XML structure requirements
    let spec = DocSpec {
        version: 1,
        page: Some(PageSetup {
            size: Some(PageSize::Named("letter".into())),
            orientation: Some(Orientation::Portrait),
            margin: Some(Margin::Custom {
                top: Some(1.0),
                bottom: Some(1.0),
                left: Some(1.25),
                right: Some(1.25),
            }),
        }),
        styles: Some(Styles {
            font: Some("Arial".into()),
            size: Some(12.0),
            color: None,
            headings: None,
            custom: None,
        }),
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![
            Block::Paragraph {
                paragraph: ParagraphContent::Full(ParagraphFull {
                    text: Some("Centered with spacing".into()),
                    runs: None,
                    align: Some("center".into()),
                    spacing: Some(Spacing {
                        before: Some(6.0),
                        after: Some(12.0),
                        line: None,
                    }),
                    indent: None,
                    style: None,
                    id: None,
                    inserted: None,
                    deleted: None,
                }),
            },
        ],
    };

    let mut buf = Cursor::new(Vec::new());
    create::create_docx(&spec, &mut buf, None).expect("create failed");
    let data = buf.into_inner();
    let xml = extract_document_xml(&data);

    // Verify w:gutter is present in pgMar
    assert!(xml.contains("w:gutter=\"0\""), "missing w:gutter on pgMar");

    // Verify no w:orient="portrait"
    assert!(
        !xml.contains("w:orient=\"portrait\""),
        "should not emit orient=portrait"
    );

    // Verify element order: spacing before jc in pPr
    if let Some(spacing_pos) = xml.find("w:spacing w:") {
        if let Some(jc_pos) = xml.find("w:jc w:val=") {
            // Only check within the same pPr context
            let ppr_start = xml[..spacing_pos].rfind("<w:pPr>").unwrap_or(0);
            let ppr_end = xml[spacing_pos..].find("</w:pPr>").map(|p| p + spacing_pos);
            if let Some(end) = ppr_end {
                if jc_pos > ppr_start && jc_pos < end {
                    assert!(
                        spacing_pos < jc_pos,
                        "w:spacing must come before w:jc in pPr"
                    );
                }
            }
        }
    }
}

#[test]
fn test_run_merging_on_unpack() {
    // Create a doc with markdown text that produces multiple runs
    let spec = DocSpec {
        version: 1,
        page: None,
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: None,
        body: vec![Block::Paragraph {
            paragraph: ParagraphContent::Simple(
                "Normal **bold** normal again.".into(),
            ),
        }],
    };

    let result = create_and_unpack(&spec);
    // The unpack should merge adjacent plain runs and produce clean markdown
    match &result.body[0] {
        Block::Paragraph {
            paragraph: ParagraphContent::Simple(text),
        } => {
            assert!(text.contains("**bold**"), "should preserve bold markdown");
            assert!(text.contains("Normal"), "should have normal text");
        }
        _ => panic!("expected simple paragraph"),
    }
}

#[test]
fn test_validate_docx_output() {
    let spec = DocSpec {
        version: 1,
        page: Some(PageSetup {
            size: Some(PageSize::Named("letter".into())),
            orientation: None,
            margin: Some(Margin::Uniform(1.0)),
        }),
        styles: None,
        headers: None,
        footers: None,
        footnotes: None,
        comments: None,
        metadata: Some(Metadata {
            title: Some("Test".into()),
            creator: Some("test".into()),
            subject: None,
            description: None,
            keywords: None,
            category: None,
        }),
        body: vec![
            Block::Heading {
                heading: 1,
                text: TextContent::Simple("Title".into()),
                id: None,
            },
            Block::Paragraph {
                paragraph: ParagraphContent::Simple("Body text.".into()),
            },
        ],
    };

    let mut buf = Cursor::new(Vec::new());
    create::create_docx(&spec, &mut buf, None).expect("create failed");
    let data = buf.into_inner();

    let cursor = Cursor::new(data);
    let result = nebo_docx::validate_docx::validate_docx(cursor).expect("validate failed");
    assert!(result.is_valid(), "DOCX should pass validation: {result}");
}

fn extract_document_xml(docx_data: &[u8]) -> String {
    let cursor = Cursor::new(docx_data);
    let mut archive = zip::ZipArchive::new(cursor).unwrap();
    let mut file = archive.by_name("word/document.xml").unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut file, &mut xml).unwrap();
    xml
}
