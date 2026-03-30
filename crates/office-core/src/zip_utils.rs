use std::io::{Read, Seek, Write};
use thiserror::Error;
use zip::write::SimpleFileOptions;

#[derive(Error, Debug)]
pub enum ZipError {
    #[error("ZIP I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),
}

/// Create a new ZIP archive, adding files from a list of (path, content) pairs.
pub fn create_zip<W: Write + Seek>(
    writer: W,
    files: &[(&str, &[u8])],
) -> Result<(), ZipError> {
    let mut zip = zip::ZipWriter::new(writer);
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (path, content) in files {
        zip.start_file(*path, options)?;
        zip.write_all(content)?;
    }

    zip.finish()?;
    Ok(())
}

/// Read all files from a ZIP archive into a Vec of (path, content) pairs.
pub fn read_zip<R: Read + Seek>(reader: R) -> Result<Vec<(String, Vec<u8>)>, ZipError> {
    let mut archive = zip::ZipArchive::new(reader)?;
    let mut files = Vec::new();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let mut content = Vec::new();
        file.read_to_end(&mut content)?;
        files.push((name, content));
    }

    Ok(files)
}
