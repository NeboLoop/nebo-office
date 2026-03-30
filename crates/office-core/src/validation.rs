/// Shared OOXML validation types used by docx, xlsx, and pptx validators.

/// A single validation finding.
#[derive(Debug)]
pub struct ValidationIssue {
    pub file: String,
    pub message: String,
    pub severity: Severity,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Severity {
    Error,
    Warning,
}

/// Results from validating an OOXML file.
#[derive(Debug)]
pub struct ValidationResult {
    pub issues: Vec<ValidationIssue>,
}

impl ValidationResult {
    pub fn is_valid(&self) -> bool {
        !self.issues.iter().any(|i| i.severity == Severity::Error)
    }

    pub fn error_count(&self) -> usize {
        self.issues
            .iter()
            .filter(|i| i.severity == Severity::Error)
            .count()
    }

    pub fn warning_count(&self) -> usize {
        self.issues
            .iter()
            .filter(|i| i.severity == Severity::Warning)
            .count()
    }
}

impl std::fmt::Display for ValidationResult {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if self.is_valid() {
            write!(f, "PASSED - All validations passed")?;
            if self.warning_count() > 0 {
                write!(f, " ({} warnings)", self.warning_count())?;
            }
        } else {
            write!(
                f,
                "FAILED - {} errors, {} warnings",
                self.error_count(),
                self.warning_count()
            )?;
        }
        for issue in &self.issues {
            let severity = match issue.severity {
                Severity::Error => "ERROR",
                Severity::Warning => "WARN",
            };
            write!(f, "\n  [{}] {}: {}", severity, issue.file, issue.message)?;
        }
        Ok(())
    }
}

/// Extract an XML attribute value from a tag string. Shared utility for validators.
pub fn extract_xml_attr(tag: &str, attr_name: &str) -> Option<String> {
    let pattern = format!("{attr_name}=\"");
    if let Some(start) = tag.find(&pattern) {
        let val_start = start + pattern.len();
        if let Some(end) = tag[val_start..].find('"') {
            return Some(tag[val_start..val_start + end].to_string());
        }
    }
    None
}
