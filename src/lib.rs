//! Create and edit `.docx` files — a [python-docx](https://python-docx.readthedocs.io/)
//! for Rust.
//!
//! **Status: early development.** The OPC packaging layer, the lossless XML tree, a
//! typed document API ([`Document`], [`Paragraph`], [`Run`]), character/paragraph
//! formatting (bold, italic, underline, size, color, font, alignment, styles), tables
//! ([`Table`], [`Row`], [`Cell`] — read rows/cells/text, merge awareness, create, and
//! edit), sections with headers/footers ([`Section`], [`HeaderFooter`], [`Length`] —
//! page geometry read/set, and header/footer text read and edit via lazily parsed parts),
//! and inline images ([`Picture`] — read existing pictures, add new ones with EMU
//! geometry, creating the media part, content-type, and relationship) are implemented.
//! See the [repository](https://github.com/jwmurray/docxml) for the architecture and
//! roadmap.
//!
//! `docxml` is built on a lossless core: every part of a package is preserved
//! byte-for-byte unless explicitly modified, so editing existing documents and
//! templates is safe. The typed handle API layers ergonomics on top.
//!
//! ```rust,ignore
//! use docxml::Document;
//!
//! let mut doc = Document::open("contract.docx")?;
//! for para in doc.paragraphs() {
//!     println!("{}", para.text(&doc));
//! }
//! let p = doc.add_paragraph("Signed and agreed:");
//! p.add_run(&mut doc, "John Murray").bold(&mut doc, true);
//! doc.save("contract-signed.docx")?;
//! # Ok::<(), docxml::Error>(())
//! ```

#![forbid(unsafe_code)]

mod api;
mod error;
pub mod opc;
pub mod xml;

pub use api::{
    Alignment, Cell, Document, HeaderFooter, Length, Paragraph, Picture, Pt, RgbColor, Row, Run,
    Section, Table, VMerge,
};
pub use error::{Error, Result};
