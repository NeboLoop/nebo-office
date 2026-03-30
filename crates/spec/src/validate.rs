use crate::types::*;
use std::path::Path;
use thiserror::Error;

#[derive(Error, Debug)]
pub enum ValidationError {
    #[error("at {path}: {message}")]
    Field { path: String, message: String },
    #[error("{}", .0.iter().map(|e| e.to_string()).collect::<Vec<_>>().join("\n"))]
    Multiple(Vec<ValidationError>),
}

pub struct ValidationOptions {
    pub strict: bool,
    pub assets_dir: Option<String>,
}

impl Default for ValidationOptions {
    fn default() -> Self {
        Self {
            strict: false,
            assets_dir: None,
        }
    }
}

pub fn validate_spec(
    spec: &DocSpec,
    options: &ValidationOptions,
) -> Result<(), ValidationError> {
    let mut errors = Vec::new();

    if spec.version != 1 {
        errors.push(ValidationError::Field {
            path: "version".into(),
            message: format!("expected 1, got {}", spec.version),
        });
    }

    if spec.body.is_empty() {
        errors.push(ValidationError::Field {
            path: "body".into(),
            message: "body must contain at least one element".into(),
        });
    }

    // Validate page setup
    if let Some(page) = &spec.page {
        validate_page(page, &mut errors);
    }

    // Validate body blocks
    for (i, block) in spec.body.iter().enumerate() {
        validate_block(block, &format!("body[{i}]"), options, &mut errors);
    }

    // Validate referenced footnotes exist
    if let Some(footnotes) = &spec.footnotes {
        for key in footnotes.keys() {
            if key.is_empty() {
                errors.push(ValidationError::Field {
                    path: "footnotes".into(),
                    message: "footnote key cannot be empty".into(),
                });
            }
        }
    }

    // Validate referenced comments exist
    if let Some(comments) = &spec.comments {
        for (key, comment) in comments {
            if comment.text.is_empty() {
                errors.push(ValidationError::Field {
                    path: format!("comments.{key}.text"),
                    message: "comment text cannot be empty".into(),
                });
            }
        }
    }

    if errors.is_empty() {
        Ok(())
    } else {
        Err(ValidationError::Multiple(errors))
    }
}

fn validate_page(page: &PageSetup, errors: &mut Vec<ValidationError>) {
    if let Some(PageSize::Named(name)) = &page.size {
        match name.as_str() {
            "letter" | "a4" | "legal" => {}
            _ => {
                errors.push(ValidationError::Field {
                    path: "page.size".into(),
                    message: format!(
                        "unknown page size '{name}'. Use 'letter', 'a4', 'legal', or {{width, height}}"
                    ),
                });
            }
        }
    }
}

fn validate_block(
    block: &Block,
    path: &str,
    options: &ValidationOptions,
    errors: &mut Vec<ValidationError>,
) {
    match block {
        Block::Heading { heading, .. } => {
            if *heading < 1 || *heading > 6 {
                errors.push(ValidationError::Field {
                    path: path.into(),
                    message: format!("heading level must be 1-6, got {heading}"),
                });
            }
        }
        Block::Image {
            image,
            image_data,
            width,
            height,
            ..
        } => {
            if image_data.is_none() {
                if options.strict {
                    if let Some(dir) = &options.assets_dir {
                        let img_path = Path::new(dir).join(image);
                        if !img_path.exists() {
                            errors.push(ValidationError::Field {
                                path: path.into(),
                                message: format!("image file not found: {}", img_path.display()),
                            });
                        }
                    }
                }
            }
            if let Some(w) = width {
                if *w <= 0.0 {
                    errors.push(ValidationError::Field {
                        path: format!("{path}.width"),
                        message: "width must be positive".into(),
                    });
                }
            }
            if let Some(h) = height {
                if *h <= 0.0 {
                    errors.push(ValidationError::Field {
                        path: format!("{path}.height"),
                        message: "height must be positive".into(),
                    });
                }
            }
        }
        Block::Table { table, .. } => match table {
            TableContent::Simple(rows) => {
                if rows.is_empty() {
                    errors.push(ValidationError::Field {
                        path: path.into(),
                        message: "table must have at least one row".into(),
                    });
                }
            }
            TableContent::Full(full) => {
                if full.rows.is_empty() {
                    errors.push(ValidationError::Field {
                        path: path.into(),
                        message: "table must have at least one row".into(),
                    });
                }
            }
        },
        _ => {}
    }
}

impl std::fmt::Display for ValidationOptions {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "strict={}", self.strict)
    }
}
