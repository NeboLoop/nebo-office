use anyhow::{Context, Result};
use clap::{Parser, Subcommand};
use std::io::{self, Read, Write};
use std::path::PathBuf;

#[cfg(feature = "docx")]
use nebo_docx::{create, unpack, validate_docx};
use nebo_spec::validate::{ValidationOptions, validate_spec};

#[derive(Parser)]
#[command(name = "nebo-office", version, about = "Generate and parse Office documents from JSON specs")]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Work with DOCX files
    #[cfg(feature = "docx")]
    Docx {
        #[command(subcommand)]
        action: DocxAction,
    },
    /// Work with XLSX files
    #[cfg(feature = "xlsx")]
    Xlsx {
        #[command(subcommand)]
        action: XlsxAction,
    },
    /// Work with PPTX files
    #[cfg(feature = "pptx")]
    Pptx {
        #[command(subcommand)]
        action: PptxAction,
    },
    /// Print version information
    Version,
}

#[cfg(feature = "docx")]
#[derive(Subcommand)]
enum DocxAction {
    /// Create a DOCX file from a JSON spec
    Create {
        /// Path to the JSON spec file, or "-" for stdin
        spec: String,
        /// Output DOCX file path, or "-" for stdout
        #[arg(short, long)]
        output: String,
        /// Directory containing image assets (defaults to spec file directory)
        #[arg(long)]
        assets: Option<PathBuf>,
        /// Validate the generated DOCX after creation
        #[arg(long)]
        validate: bool,
    },
    /// Unpack a DOCX file into a JSON spec
    Unpack {
        /// Path to the DOCX file, or "-" for stdin
        input: String,
        /// Output JSON spec file path, or "-" for stdout
        #[arg(short, long)]
        output: String,
        /// Directory to extract image assets to
        #[arg(long)]
        assets: Option<PathBuf>,
        /// Pretty-print JSON output
        #[arg(long)]
        pretty: bool,
    },
    /// Validate a JSON spec or DOCX file
    Validate {
        /// Path to the JSON spec or DOCX file, or "-" for stdin
        spec: String,
        /// Directory containing image assets
        #[arg(long)]
        assets: Option<PathBuf>,
        /// Enable strict validation (check image files exist, etc.)
        #[arg(long)]
        strict: bool,
    },
}

#[cfg(feature = "xlsx")]
#[derive(Subcommand)]
enum XlsxAction {
    /// Create an XLSX file from a JSON spec
    Create {
        spec: String,
        #[arg(short, long)]
        output: String,
        #[arg(long)]
        assets: Option<PathBuf>,
    },
    /// Unpack an XLSX file into a JSON spec
    Unpack {
        input: String,
        #[arg(short, long)]
        output: String,
        #[arg(long)]
        assets: Option<PathBuf>,
        #[arg(long)]
        pretty: bool,
    },
    /// Validate a JSON spec
    Validate {
        spec: String,
        #[arg(long)]
        assets: Option<PathBuf>,
        #[arg(long)]
        strict: bool,
    },
}

#[cfg(feature = "pptx")]
#[derive(Subcommand)]
enum PptxAction {
    /// Create a PPTX file from a JSON spec
    Create {
        spec: String,
        #[arg(short, long)]
        output: String,
        #[arg(long)]
        assets: Option<PathBuf>,
    },
    /// Unpack a PPTX file into a JSON spec
    Unpack {
        input: String,
        #[arg(short, long)]
        output: String,
        #[arg(long)]
        assets: Option<PathBuf>,
        #[arg(long)]
        pretty: bool,
    },
    /// Validate a JSON spec
    Validate {
        spec: String,
        #[arg(long)]
        assets: Option<PathBuf>,
        #[arg(long)]
        strict: bool,
    },
}

fn main() {
    if let Err(e) = run() {
        eprintln!("error: {e:#}");
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Version => {
            println!("nebo-office {}", env!("CARGO_PKG_VERSION"));
        }
        #[cfg(feature = "docx")]
        Commands::Docx { action } => run_docx(action)?,
        #[cfg(feature = "xlsx")]
        Commands::Xlsx { action } => run_xlsx(action)?,
        #[cfg(feature = "pptx")]
        Commands::Pptx { action } => run_pptx(action)?,
    }

    Ok(())
}

#[cfg(feature = "docx")]
fn run_docx(action: DocxAction) -> Result<()> {
    use nebo_spec::DocSpec;

    match action {
        DocxAction::Create {
            spec,
            output,
            assets,
            validate,
        } => {
            let json = read_input(&spec)?;
            let doc: DocSpec = serde_json::from_str(&json)
                .context("failed to parse JSON spec")?;

            let assets_dir = assets
                .or_else(|| {
                    if spec != "-" {
                        PathBuf::from(&spec).parent().map(|p| p.to_path_buf())
                    } else {
                        None
                    }
                });

            let mut buf = io::Cursor::new(Vec::new());
            create::create_docx(&doc, &mut buf, assets_dir.as_deref())?;
            let data = buf.into_inner();

            if validate {
                let cursor = io::Cursor::new(&data);
                let result = validate_docx::validate_docx(cursor)?;
                eprintln!("{result}");
                if !result.is_valid() {
                    std::process::exit(1);
                }
            }

            write_output(&output, &data)?;
            if output != "-" {
                eprintln!("Created: {output}");
            }
        }
        DocxAction::Unpack {
            input,
            output,
            assets,
            pretty,
        } => {
            let data = read_input_bytes(&input)?;
            let cursor = io::Cursor::new(data);
            let spec = unpack::unpack_docx(cursor, assets.as_deref(), pretty)?;
            write_json_output(&output, &spec, pretty)?;
            if output != "-" {
                eprintln!("Unpacked: {output}");
            }
        }
        DocxAction::Validate {
            spec,
            assets,
            strict,
        } => {
            let is_docx = spec.ends_with(".docx") || spec.ends_with(".DOCX");

            if is_docx {
                let data = read_input_bytes(&spec)?;
                let cursor = io::Cursor::new(data);
                let result = validate_docx::validate_docx(cursor)?;
                eprintln!("{result}");
                if !result.is_valid() {
                    std::process::exit(1);
                }
            } else {
                let json = read_input(&spec)?;
                let doc: DocSpec = serde_json::from_str(&json)
                    .context("failed to parse JSON spec")?;
                let options = ValidationOptions {
                    strict,
                    assets_dir: assets.map(|p| p.to_string_lossy().to_string()),
                };
                match validate_spec(&doc, &options) {
                    Ok(()) => eprintln!("Validation passed."),
                    Err(e) => {
                        eprintln!("Validation failed:\n{e}");
                        std::process::exit(1);
                    }
                }
            }
        }
    }
    Ok(())
}

#[cfg(feature = "xlsx")]
fn run_xlsx(action: XlsxAction) -> Result<()> {
    use nebo_spec::XlsxSpec;

    match action {
        XlsxAction::Create {
            spec,
            output,
            assets,
        } => {
            let json = read_input(&spec)?;
            let xlsx_spec: XlsxSpec = serde_json::from_str(&json)
                .context("failed to parse JSON spec")?;

            let assets_dir = assets
                .or_else(|| {
                    if spec != "-" {
                        PathBuf::from(&spec).parent().map(|p| p.to_path_buf())
                    } else {
                        None
                    }
                });

            let mut buf = io::Cursor::new(Vec::new());
            nebo_xlsx::create::create_xlsx(&xlsx_spec, &mut buf, assets_dir.as_deref())?;
            let data = buf.into_inner();
            write_output(&output, &data)?;
            if output != "-" {
                eprintln!("Created: {output}");
            }
        }
        XlsxAction::Unpack {
            input,
            output,
            assets,
            pretty,
        } => {
            let data = read_input_bytes(&input)?;
            let cursor = io::Cursor::new(data);
            let spec = nebo_xlsx::unpack::unpack_xlsx(cursor, assets.as_deref(), pretty)?;
            write_json_output(&output, &spec, pretty)?;
            if output != "-" {
                eprintln!("Unpacked: {output}");
            }
        }
        XlsxAction::Validate {
            spec,
            assets,
            strict,
        } => {
            let json = read_input(&spec)?;
            let _xlsx_spec: XlsxSpec = serde_json::from_str(&json)
                .context("failed to parse JSON spec")?;
            eprintln!("Validation passed.");
        }
    }
    Ok(())
}

#[cfg(feature = "pptx")]
fn run_pptx(action: PptxAction) -> Result<()> {
    use nebo_spec::PptxSpec;

    match action {
        PptxAction::Create {
            spec,
            output,
            assets,
        } => {
            let json = read_input(&spec)?;
            let pptx_spec: PptxSpec = serde_json::from_str(&json)
                .context("failed to parse JSON spec")?;

            let assets_dir = assets
                .or_else(|| {
                    if spec != "-" {
                        PathBuf::from(&spec).parent().map(|p| p.to_path_buf())
                    } else {
                        None
                    }
                });

            let mut buf = io::Cursor::new(Vec::new());
            nebo_pptx::create::create_pptx(&pptx_spec, &mut buf, assets_dir.as_deref())?;
            let data = buf.into_inner();
            write_output(&output, &data)?;
            if output != "-" {
                eprintln!("Created: {output}");
            }
        }
        PptxAction::Unpack {
            input,
            output,
            assets,
            pretty,
        } => {
            let data = read_input_bytes(&input)?;
            let cursor = io::Cursor::new(data);
            let spec = nebo_pptx::unpack::unpack_pptx(cursor, assets.as_deref(), pretty)?;
            write_json_output(&output, &spec, pretty)?;
            if output != "-" {
                eprintln!("Unpacked: {output}");
            }
        }
        PptxAction::Validate {
            spec,
            assets,
            strict,
        } => {
            let json = read_input(&spec)?;
            let _pptx_spec: PptxSpec = serde_json::from_str(&json)
                .context("failed to parse JSON spec")?;
            eprintln!("Validation passed.");
        }
    }
    Ok(())
}

// --- I/O helpers ---

fn read_input(path: &str) -> Result<String> {
    if path == "-" {
        let mut buf = String::new();
        io::stdin().read_to_string(&mut buf)?;
        Ok(buf)
    } else {
        std::fs::read_to_string(path)
            .with_context(|| format!("failed to read {path}"))
    }
}

fn read_input_bytes(path: &str) -> Result<Vec<u8>> {
    if path == "-" {
        let mut buf = Vec::new();
        io::stdin().read_to_end(&mut buf)?;
        Ok(buf)
    } else {
        std::fs::read(path)
            .with_context(|| format!("failed to read {path}"))
    }
}

fn write_output(path: &str, data: &[u8]) -> Result<()> {
    if path == "-" {
        io::stdout().write_all(data)?;
    } else {
        std::fs::write(path, data)
            .with_context(|| format!("failed to write {path}"))?;
    }
    Ok(())
}

fn write_json_output<T: serde::Serialize>(path: &str, value: &T, pretty: bool) -> Result<()> {
    let json = if pretty {
        serde_json::to_string_pretty(value)?
    } else {
        serde_json::to_string(value)?
    };
    if path == "-" {
        io::stdout().write_all(json.as_bytes())?;
        io::stdout().write_all(b"\n")?;
    } else {
        std::fs::write(path, json.as_bytes())
            .with_context(|| format!("failed to write {path}"))?;
    }
    Ok(())
}
