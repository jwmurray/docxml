//! Create and edit `.docx` files — a [python-docx](https://python-docx.readthedocs.io/)
//! for Rust.
//!
//! **Status: early development.** This release reserves the crate name; the API is not
//! yet implemented. See the [repository](https://github.com/jwmurray/docxml) for the
//! architecture and roadmap.
//!
//! `docxml` is built on a lossless, mutable XML tree: everything it doesn't understand
//! in a document passes through untouched on save, so editing existing documents and
//! templates is safe. A typed handle API (`Document`, `Paragraph`, `Run`, `Table`)
//! layers ergonomics on top.

#![forbid(unsafe_code)]
