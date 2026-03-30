//! Post-creation DOCX structural validation.
//!
//! Validates the internal XML of a DOCX file to catch common OOXML compliance
//! issues. Mirrors the checks from the Python validators (base.py + docx.py).

use anyhow::{Context, Result};
use nebo_office_core::validation::{
    extract_xml_attr, Severity, ValidationIssue, ValidationResult,
};
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

/// Validate a DOCX file's internal XML structure.
pub fn validate_docx<R: Read + Seek>(reader: R) -> Result<ValidationResult> {
    let mut archive = zip::ZipArchive::new(reader).context("failed to open DOCX as ZIP")?;
    let mut issues = Vec::new();

    // Read all files into memory
    let mut files: HashMap<String, String> = HashMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        if name.ends_with(".xml") || name.ends_with(".rels") {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            files.insert(name, content);
        }
    }

    // 1. Validate XML well-formedness
    validate_xml_wellformed(&files, &mut issues);

    // 2. Validate content types
    validate_content_types(&files, &mut issues);

    // 3. Validate relationship references
    validate_relationships(&files, &mut issues);

    // 4. Validate whitespace preservation
    if let Some(doc_xml) = files.get("word/document.xml") {
        validate_whitespace_preservation(doc_xml, "word/document.xml", &mut issues);
    }

    // 5. Validate tracked changes integrity
    if let Some(doc_xml) = files.get("word/document.xml") {
        validate_deletions(doc_xml, &mut issues);
        validate_insertions(doc_xml, &mut issues);
    }

    // 6. Validate comment markers
    validate_comment_markers(&files, &mut issues);

    // 7. Validate OOXML element ordering in document.xml
    if let Some(doc_xml) = files.get("word/document.xml") {
        validate_element_ordering(doc_xml, &mut issues);
    }

    // 8. Validate required attributes
    if let Some(doc_xml) = files.get("word/document.xml") {
        validate_required_attributes(doc_xml, &mut issues);
    }

    Ok(ValidationResult { issues })
}

/// Check all XML files parse correctly.
fn validate_xml_wellformed(files: &HashMap<String, String>, issues: &mut Vec<ValidationIssue>) {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    for (name, content) in files {
        let mut reader = Reader::from_str(content);
        loop {
            match reader.read_event() {
                Ok(Event::Eof) => break,
                Err(e) => {
                    issues.push(ValidationIssue {
                        file: name.clone(),
                        message: format!("XML parse error: {e}"),
                        severity: Severity::Error,
                    });
                    break;
                }
                _ => {}
            }
        }
    }
}

/// Validate [Content_Types].xml has entries for all XML parts.
fn validate_content_types(files: &HashMap<String, String>, issues: &mut Vec<ValidationIssue>) {
    let ct_xml = match files.get("[Content_Types].xml") {
        Some(xml) => xml,
        None => {
            issues.push(ValidationIssue {
                file: "[Content_Types].xml".into(),
                message: "missing [Content_Types].xml".into(),
                severity: Severity::Error,
            });
            return;
        }
    };

    // Collect declared part names and extensions
    let mut declared_parts: HashSet<String> = HashSet::new();
    let mut declared_extensions: HashSet<String> = HashSet::new();

    // Parse Override elements
    let mut pos = 0;
    while let Some(start) = ct_xml[pos..].find("<Override") {
        let abs = pos + start;
        let end = ct_xml[abs..].find("/>").unwrap_or(ct_xml.len() - abs) + abs;
        let chunk = &ct_xml[abs..end + 2];
        if let Some(pn) = extract_xml_attr(chunk, "PartName") {
            declared_parts.insert(pn.trim_start_matches('/').to_string());
        }
        pos = end + 2;
    }

    // Parse Default elements
    pos = 0;
    while let Some(start) = ct_xml[pos..].find("<Default") {
        let abs = pos + start;
        let end = ct_xml[abs..].find("/>").unwrap_or(ct_xml.len() - abs) + abs;
        let chunk = &ct_xml[abs..end + 2];
        if let Some(ext) = extract_xml_attr(chunk, "Extension") {
            declared_extensions.insert(ext.to_lowercase());
        }
        pos = end + 2;
    }

    // Check that key parts are declared
    let required_parts = ["word/document.xml"];
    for part in required_parts {
        if !declared_parts.contains(part) {
            issues.push(ValidationIssue {
                file: "[Content_Types].xml".into(),
                message: format!("missing Override for {part}"),
                severity: Severity::Error,
            });
        }
    }

    // Check that .rels and .xml extensions have defaults
    for ext in ["rels", "xml"] {
        if !declared_extensions.contains(ext) {
            issues.push(ValidationIssue {
                file: "[Content_Types].xml".into(),
                message: format!("missing Default for .{ext} extension"),
                severity: Severity::Error,
            });
        }
    }
}

/// Validate relationship files: all referenced targets must exist, no duplicate IDs.
fn validate_relationships(files: &HashMap<String, String>, issues: &mut Vec<ValidationIssue>) {
    for (name, content) in files {
        if !name.ends_with(".rels") {
            continue;
        }

        let mut seen_ids: HashSet<String> = HashSet::new();

        let mut pos = 0;
        while let Some(start) = content[pos..].find("<Relationship") {
            let abs = pos + start;
            let end = content[abs..].find("/>").unwrap_or(content.len() - abs) + abs;
            let chunk = &content[abs..end + 2];

            if let Some(id) = extract_xml_attr(chunk, "Id") {
                if !seen_ids.insert(id.clone()) {
                    issues.push(ValidationIssue {
                        file: name.clone(),
                        message: format!("duplicate relationship ID: {id}"),
                        severity: Severity::Error,
                    });
                }
            }

            if let Some(target) = extract_xml_attr(chunk, "Target") {
                let target_mode = extract_xml_attr(chunk, "TargetMode");
                // Only check internal targets
                if target_mode.as_deref() != Some("External")
                    && !target.starts_with("http")
                    && !target.starts_with("mailto:")
                {
                    // Resolve relative path
                    let base = if name.contains("word/") {
                        "word/"
                    } else {
                        ""
                    };
                    let full_path = if target.starts_with('/') {
                        target.trim_start_matches('/').to_string()
                    } else {
                        format!("{base}{target}")
                    };

                    // Check if the target file exists in the archive
                    if !files.contains_key(&full_path) {
                        // Check if it's a non-XML file (images etc.) — we only loaded XML
                        // so we can't check media files. Skip those.
                        if full_path.ends_with(".xml") || full_path.ends_with(".rels") {
                            issues.push(ValidationIssue {
                                file: name.clone(),
                                message: format!("broken reference to {target} (resolved: {full_path})"),
                                severity: Severity::Error,
                            });
                        }
                    }
                }
            }

            pos = end + 2;
        }
    }
}

/// Validate xml:space="preserve" on w:t elements with leading/trailing whitespace.
fn validate_whitespace_preservation(
    xml: &str,
    filename: &str,
    issues: &mut Vec<ValidationIssue>,
) {
    // Find all <w:t ...>text</w:t> patterns and check for whitespace
    let mut pos = 0;
    while let Some(start) = xml[pos..].find("<w:t") {
        let abs = pos + start;
        // Find the end of the opening tag
        let tag_end = match xml[abs..].find('>') {
            Some(e) => abs + e,
            None => break,
        };

        // Check if it's a self-closing tag
        if xml[tag_end - 1..tag_end] == *"/" {
            pos = tag_end + 1;
            continue;
        }

        // Find closing tag
        let close = match xml[tag_end..].find("</w:t>") {
            Some(e) => tag_end + e,
            None => break,
        };

        let text_content = &xml[tag_end + 1..close];
        let tag_content = &xml[abs..tag_end + 1];

        // Check if text has leading/trailing whitespace
        if !text_content.is_empty()
            && (text_content.starts_with(' ')
                || text_content.starts_with('\t')
                || text_content.ends_with(' ')
                || text_content.ends_with('\t'))
        {
            if !tag_content.contains("xml:space=\"preserve\"") {
                let preview = if text_content.len() > 30 {
                    format!("{}...", &text_content[..30])
                } else {
                    text_content.to_string()
                };
                issues.push(ValidationIssue {
                    file: filename.into(),
                    message: format!(
                        "w:t with whitespace missing xml:space=\"preserve\": {:?}",
                        preview
                    ),
                    severity: Severity::Error,
                });
            }
        }

        pos = close + 6;
    }
}

/// Validate that no w:t elements appear inside w:del (should use w:delText).
fn validate_deletions(xml: &str, issues: &mut Vec<ValidationIssue>) {
    // Find all <w:del> ... </w:del> blocks
    let mut pos = 0;
    while let Some(start) = xml[pos..].find("<w:del ") {
        let abs = pos + start;
        let end = match xml[abs..].find("</w:del>") {
            Some(e) => abs + e,
            None => break,
        };

        let del_content = &xml[abs..end];

        // Check for <w:t> inside deletion (should be <w:delText>)
        if del_content.contains("<w:t ") || del_content.contains("<w:t>") {
            issues.push(ValidationIssue {
                file: "word/document.xml".into(),
                message: "w:t found inside w:del (should use w:delText)".into(),
                severity: Severity::Error,
            });
        }

        pos = end + 8;
    }
}

/// Validate that no w:delText elements appear inside w:ins without w:del.
fn validate_insertions(xml: &str, issues: &mut Vec<ValidationIssue>) {
    let mut pos = 0;
    while let Some(start) = xml[pos..].find("<w:ins ") {
        let abs = pos + start;
        let end = match xml[abs..].find("</w:ins>") {
            Some(e) => abs + e,
            None => break,
        };

        let ins_content = &xml[abs..end];

        // Check for <w:delText> inside insertion (without a nested w:del)
        if (ins_content.contains("<w:delText ") || ins_content.contains("<w:delText>"))
            && !ins_content.contains("<w:del ")
        {
            issues.push(ValidationIssue {
                file: "word/document.xml".into(),
                message: "w:delText found inside w:ins without w:del ancestor".into(),
                severity: Severity::Error,
            });
        }

        pos = end + 8;
    }
}

/// Validate comment markers are properly paired.
fn validate_comment_markers(files: &HashMap<String, String>, issues: &mut Vec<ValidationIssue>) {
    let doc_xml = match files.get("word/document.xml") {
        Some(xml) => xml,
        None => return,
    };

    let mut range_starts: HashSet<String> = HashSet::new();
    let mut range_ends: HashSet<String> = HashSet::new();
    let mut references: HashSet<String> = HashSet::new();

    // Collect commentRangeStart IDs
    let mut pos = 0;
    while let Some(start) = doc_xml[pos..].find("commentRangeStart") {
        let abs = pos + start;
        let end = abs + 100.min(doc_xml.len() - abs);
        let chunk = &doc_xml[abs..end];
        if let Some(id) = extract_xml_attr(chunk, "w:id") {
            range_starts.insert(id);
        }
        pos = abs + 17;
    }

    // Collect commentRangeEnd IDs
    pos = 0;
    while let Some(start) = doc_xml[pos..].find("commentRangeEnd") {
        let abs = pos + start;
        let end = abs + 100.min(doc_xml.len() - abs);
        let chunk = &doc_xml[abs..end];
        if let Some(id) = extract_xml_attr(chunk, "w:id") {
            range_ends.insert(id);
        }
        pos = abs + 15;
    }

    // Collect commentReference IDs
    pos = 0;
    while let Some(start) = doc_xml[pos..].find("commentReference") {
        let abs = pos + start;
        let end = abs + 100.min(doc_xml.len() - abs);
        let chunk = &doc_xml[abs..end];
        if let Some(id) = extract_xml_attr(chunk, "w:id") {
            references.insert(id);
        }
        pos = abs + 16;
    }

    // Check for orphaned ends (end without matching start)
    for id in range_ends.difference(&range_starts) {
        issues.push(ValidationIssue {
            file: "word/document.xml".into(),
            message: format!("commentRangeEnd id=\"{id}\" has no matching commentRangeStart"),
            severity: Severity::Error,
        });
    }

    // Check for orphaned starts (start without matching end)
    for id in range_starts.difference(&range_ends) {
        issues.push(ValidationIssue {
            file: "word/document.xml".into(),
            message: format!("commentRangeStart id=\"{id}\" has no matching commentRangeEnd"),
            severity: Severity::Error,
        });
    }

    // If we have comments.xml, check that markers reference valid comment IDs
    if let Some(comments_xml) = files.get("word/comments.xml") {
        let mut comment_ids: HashSet<String> = HashSet::new();
        let mut cpos = 0;
        while let Some(start) = comments_xml[cpos..].find("<w:comment ") {
            let abs = cpos + start;
            let end = abs + 200.min(comments_xml.len() - abs);
            let chunk = &comments_xml[abs..end];
            if let Some(id) = extract_xml_attr(chunk, "w:id") {
                comment_ids.insert(id);
            }
            cpos = abs + 11;
        }

        let all_marker_ids: HashSet<&String> =
            range_starts.iter().chain(range_ends.iter()).chain(references.iter()).collect();

        for id in all_marker_ids {
            if !comment_ids.contains(id) {
                issues.push(ValidationIssue {
                    file: "word/document.xml".into(),
                    message: format!("marker id=\"{id}\" references non-existent comment"),
                    severity: Severity::Error,
                });
            }
        }
    }
}

/// Validate OOXML element ordering within pPr and sectPr.
fn validate_element_ordering(xml: &str, issues: &mut Vec<ValidationIssue>) {
    // Check pPr child ordering: pStyle, numPr, spacing, ind, jc, rPr
    let ppr_order = &["pStyle", "numPr", "spacing", "ind", "jc", "rPr"];
    validate_child_order(xml, "pPr", ppr_order, issues);

    // Check rPr child ordering in runs: rFonts, b, i, u, strike, vertAlign, sz, color, highlight
    let rpr_order = &["rStyle", "rFonts", "b", "i", "u", "strike", "vertAlign", "sz", "szCs", "color", "highlight", "caps", "smallCaps"];
    validate_child_order(xml, "rPr", rpr_order, issues);
}

fn validate_child_order(
    xml: &str,
    parent_tag: &str,
    expected_order: &[&str],
    issues: &mut Vec<ValidationIssue>,
) {
    let open_tag = format!("<w:{parent_tag}>");
    let close_tag = format!("</w:{parent_tag}>");

    let mut pos = 0;
    while let Some(start) = xml[pos..].find(&open_tag) {
        let abs = pos + start + open_tag.len();
        let end = match xml[abs..].find(&close_tag) {
            Some(e) => abs + e,
            None => break,
        };

        let content = &xml[abs..end];

        // Find positions of each child element
        let mut last_order_idx = 0;
        for &tag in expected_order {
            let patterns = [
                format!("<w:{tag} "),
                format!("<w:{tag}/>"),
                format!("<w:{tag}>"),
            ];
            for pattern in &patterns {
                if let Some(child_pos) = content.find(pattern.as_str()) {
                    // Find this tag's order index
                    if let Some(order_idx) = expected_order.iter().position(|&t| t == tag) {
                        if order_idx < last_order_idx {
                            issues.push(ValidationIssue {
                                file: "word/document.xml".into(),
                                message: format!(
                                    "w:{tag} appears after w:{} in w:{parent_tag} (violates schema order)",
                                    expected_order[last_order_idx]
                                ),
                                severity: Severity::Warning,
                            });
                        }
                        // Only update if we found it at this position
                        if child_pos > 0 || order_idx >= last_order_idx {
                            last_order_idx = order_idx;
                        }
                    }
                    break;
                }
            }
        }

        pos = end + close_tag.len();
    }
}

/// Validate required attributes on OOXML elements.
fn validate_required_attributes(xml: &str, issues: &mut Vec<ValidationIssue>) {
    // Check pgMar has w:gutter attribute
    if let Some(pos) = xml.find("<w:pgMar ") {
        let end = pos + 200.min(xml.len() - pos);
        let chunk = &xml[pos..end];
        if !chunk.contains("w:gutter=") {
            issues.push(ValidationIssue {
                file: "word/document.xml".into(),
                message: "w:pgMar missing required w:gutter attribute".into(),
                severity: Severity::Error,
            });
        }
    }

    // Check that orient="portrait" is NOT emitted (it's the default)
    if xml.contains("w:orient=\"portrait\"") {
        issues.push(ValidationIssue {
            file: "word/document.xml".into(),
            message: "w:orient=\"portrait\" should not be emitted (portrait is default)".into(),
            severity: Severity::Warning,
        });
    }
}

// extract_xml_attr is imported from nebo_office_core::validation

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_whitespace_validation() {
        let mut issues = Vec::new();

        // Good: has xml:space="preserve"
        let xml = r#"<w:t xml:space="preserve"> hello </w:t>"#;
        validate_whitespace_preservation(xml, "test.xml", &mut issues);
        assert!(issues.is_empty());

        // Bad: missing xml:space="preserve"
        let xml = r#"<w:t> hello </w:t>"#;
        validate_whitespace_preservation(xml, "test.xml", &mut issues);
        assert_eq!(issues.len(), 1);
        assert_eq!(issues[0].severity, Severity::Error);
    }

    #[test]
    fn test_deletion_validation() {
        let mut issues = Vec::new();

        // Good: uses w:delText inside w:del
        let xml = r#"<w:del w:id="0"><w:r><w:delText>removed</w:delText></w:r></w:del>"#;
        validate_deletions(xml, &mut issues);
        assert!(issues.is_empty());

        // Bad: uses w:t inside w:del
        let xml = r#"<w:del w:id="0"><w:r><w:t>removed</w:t></w:r></w:del>"#;
        validate_deletions(xml, &mut issues);
        assert_eq!(issues.len(), 1);
    }

    #[test]
    fn test_required_attributes() {
        let mut issues = Vec::new();

        // Good: has w:gutter
        let xml = r#"<w:pgMar w:top="1440" w:gutter="0"/>"#;
        validate_required_attributes(xml, &mut issues);
        assert!(issues.is_empty());

        // Bad: missing w:gutter
        let xml = r#"<w:pgMar w:top="1440"/>"#;
        validate_required_attributes(xml, &mut issues);
        assert_eq!(issues.len(), 1);
    }

    #[test]
    fn test_comment_markers_paired() {
        let mut files = HashMap::new();
        let doc_xml = r#"<w:document><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>"#;
        files.insert("word/document.xml".to_string(), doc_xml.to_string());

        let comments = r#"<w:comments><w:comment w:id="0" w:author="Test"><w:p><w:r><w:t>comment</w:t></w:r></w:p></w:comment></w:comments>"#;
        files.insert("word/comments.xml".to_string(), comments.to_string());

        let mut issues = Vec::new();
        validate_comment_markers(&files, &mut issues);
        assert!(issues.is_empty());
    }
}
