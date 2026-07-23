//! Create and edit `.docx` files — a [python-docx](https://python-docx.readthedocs.io/)
//! for Rust.
//!
//! **Status: early development.** The OPC packaging layer is implemented; the typed
//! document API is not yet. See the [repository](https://github.com/jwmurray/docxml)
//! for the architecture and roadmap.
//!
//! `docxml` is built on a lossless core: every part of a package is preserved
//! byte-for-byte unless explicitly modified, so editing existing documents and
//! templates is safe. A typed handle API (`Document`, `Paragraph`, `Run`, `Table`)
//! will layer ergonomics on top.

#![forbid(unsafe_code)]

mod error;
pub mod opc;

pub use error::{Error, Result};
