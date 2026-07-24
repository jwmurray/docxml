//! Open Packaging Conventions (OPC) layer: the zip container, its parts, and
//! package relationships.
//!
//! Fidelity contract: every part's bytes are preserved exactly as read unless the
//! caller explicitly replaces them. Opening a package and saving it produces a
//! package whose parts are byte-identical to the original.

mod rels;

pub use rels::{Relationship, parse_relationships};

use std::fs::File;
use std::io::{BufReader, BufWriter, Read, Seek, Write};
use std::path::Path;

use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::error::{Error, Result};

/// Relationship type identifying the main document part (transitional and strict).
const OFFICE_DOCUMENT_REL_TYPES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
];

/// One part (file) inside the package, e.g. `word/document.xml`.
#[derive(Debug, Clone)]
pub struct Part {
    /// Zip entry name, without a leading slash.
    pub name: String,
    /// Raw bytes of the part, preserved exactly as read.
    pub data: Vec<u8>,
}

/// An OPC package: an ordered collection of parts read from a `.docx` zip.
///
/// Part order is preserved from the source archive so saves are as close to the
/// original layout as possible.
#[derive(Debug, Clone)]
pub struct Package {
    parts: Vec<Part>,
}

impl Package {
    /// Open a package from a file on disk.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::read(BufReader::new(File::open(path)?))
    }

    /// Read a package from any seekable reader.
    pub fn read(reader: impl Read + Seek) -> Result<Self> {
        let mut archive = ZipArchive::new(reader)?;
        let mut parts = Vec::with_capacity(archive.len());
        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            if entry.is_dir() {
                continue;
            }
            let mut data = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut data)?;
            parts.push(Part {
                name: entry.name().to_owned(),
                data,
            });
        }
        if parts.is_empty() {
            return Err(Error::InvalidPackage("archive contains no parts".into()));
        }
        Ok(Self { parts })
    }

    /// Save the package to a file on disk.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.write(BufWriter::new(File::create(path)?))
    }

    /// Write the package to any seekable writer.
    pub fn write(&self, writer: impl Write + Seek) -> Result<()> {
        let mut zip = ZipWriter::new(writer);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for part in &self.parts {
            zip.start_file(&part.name, options)?;
            zip.write_all(&part.data)?;
        }
        zip.finish()?;
        Ok(())
    }

    /// All parts, in original archive order.
    pub fn parts(&self) -> &[Part] {
        &self.parts
    }

    /// Add a part to the package.
    ///
    /// The name is normalized to its OPC part-name form (any leading `/` is dropped). If
    /// a part with that name already exists its bytes are **replaced in place**, keeping
    /// its position in the part order; otherwise the new part is appended after the
    /// existing parts. Replacing rather than erroring keeps callers that regenerate a
    /// derived part (e.g. a media file or `[Content_Types].xml`) idempotent.
    pub fn add_part(&mut self, name: String, data: Vec<u8>) {
        let name = name.trim_start_matches('/').to_string();
        if let Some(existing) = self.parts.iter_mut().find(|p| p.name == name) {
            existing.data = data;
        } else {
            self.parts.push(Part { name, data });
        }
    }

    /// Look up a part by name (leading slashes ignored, per OPC part-name form).
    pub fn part(&self, name: &str) -> Option<&Part> {
        let name = name.trim_start_matches('/');
        self.parts.iter().find(|p| p.name == name)
    }

    /// Mutable lookup of a part by name.
    pub fn part_mut(&mut self, name: &str) -> Option<&mut Part> {
        let name = name.trim_start_matches('/');
        self.parts.iter_mut().find(|p| p.name == name)
    }

    /// Package-level relationships from `_rels/.rels`.
    pub fn relationships(&self) -> Result<Vec<Relationship>> {
        let rels = self
            .part("_rels/.rels")
            .ok_or_else(|| Error::InvalidPackage("missing _rels/.rels".into()))?;
        parse_relationships(&rels.data)
    }

    /// The main document part (`word/document.xml` in practice), located via the
    /// officeDocument package relationship rather than by hard-coded name.
    pub fn main_document_part(&self) -> Result<&Part> {
        let rels = self.relationships()?;
        let rel = rels
            .iter()
            .find(|r| OFFICE_DOCUMENT_REL_TYPES.contains(&r.rel_type.as_str()))
            .ok_or_else(|| Error::InvalidPackage("no officeDocument relationship".into()))?;
        self.part(&rel.target).ok_or_else(|| {
            Error::InvalidPackage(format!("officeDocument target missing: {}", rel.target))
        })
    }
}
