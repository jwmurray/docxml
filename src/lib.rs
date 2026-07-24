//! Create and edit `.docx` files — a [python-docx](https://python-docx.readthedocs.io/)
//! for Rust.
//!
//! **Status: functional, pre-1.0 — full python-docx-parity coverage of the measured
//! production feature set (every milestone complete).** The OPC packaging layer, the lossless XML tree, a
//! typed document API ([`Document`], [`Paragraph`], [`Run`]), character/paragraph
//! formatting (bold, italic, underline, size, color, font, small-caps/all-caps, alignment,
//! line spacing, space before/after, indents, tab stops, keep-together/keep-with-next/
//! page-break-before, styles), breaks and field codes (page/column/line breaks, `PAGE`/
//! `TOC` fields), tables ([`Table`], [`Row`], [`Cell`] — read rows/cells/text, merge
//! awareness, create, and edit, plus column widths / fixed layout, grid-based cell
//! addressing, and merge creation), sections with headers/footers ([`Section`],
//! [`HeaderFooter`], [`Length`] — page geometry read/set, and header/footer text read and
//! edit via lazily parsed parts), inline images ([`Picture`] — read existing pictures,
//! add new ones with EMU geometry, creating the media part, content-type, and
//! relationship), and numbering / lists ([`NumberFormat`] — read a paragraph's numbering
//! (direct or style-resolved), apply/clear it, `add_bullet_paragraph` /
//! `add_numbered_paragraph` convenience, and `create_numbering` for independent restartable
//! list definitions, creating the numbering part when absent), and header/footer part
//! creation ([`HeaderFooterType`] — create default/first/even headers and footers, building
//! the part, content-type, and relationship; `add_paragraph` into them; the
//! `w:titlePg` different-first-page flag and the `w:evenAndOddHeaders` settings flag,
//! creating `word/settings.xml` when absent), and hyperlinks ([`HyperlinkInfo`] — read a
//! paragraph's hyperlinks (external `r:id` targets resolved through the owning part's rels,
//! internal `w:anchor` targets), `add_hyperlink` creating the External relationship,
//! `add_anchor_hyperlink`, and `add_bookmark`), and section line numbering, paragraph
//! frames/borders, and hidden text ([`LineNumbering`] / [`LineNumberRestart`] —
//! `Section::set_line_numbering` for `w:lnNumType` pleading-paper numbering, and
//! `Paragraph::suppress_line_numbers`; [`FrameOptions`] / [`FrameAnchor`] / [`FrameWrap`] —
//! `Paragraph::set_frame` for `w:framePr`; [`BorderEdge`] / [`BorderStyle`] —
//! `Paragraph::set_borders` for `w:pBdr`; and `Run::set_vanish` for `w:vanish` hidden text),
//! and styles ([`Style`] / [`StyleType`] — the `word/styles.xml` catalog via
//! [`Document::styles`], [`style_by_id`](Document::style_by_id), and
//! [`style_by_name`](Document::style_by_name); [`Document::create_style`] authoring with
//! `set_based_on`/`set_next` and bold/size/color/font/alignment/spacing setters (reusing the
//! run and paragraph writers); [`Paragraph::style_name`] and [`Run::set_style_id`];
//! [`Document::set_default_font`] writing `w:docDefaults`; and the style-aware effective
//! reads [`Run::effective_bold`], `effective_italic`, `effective_size`, and `effective_font`
//! that resolve through the `w:rStyle`/`w:pStyle` `w:basedOn` chains and `w:docDefaults`)
//! are implemented. See the
//! [repository](https://github.com/jwmurray/docxml) for the architecture and roadmap.
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
    Alignment, BorderEdge, BorderStyle, BreakType, Cell, Document, FrameAnchor, FrameOptions,
    FrameWrap, HeaderFooter, HeaderFooterType, HyperlinkInfo, Length, LineNumberRestart,
    LineNumbering, LineSpacing, NumberFormat, Paragraph, Picture, Pt, RgbColor, Row, Run, Section,
    Style, StyleType, TabAlignment, TabLeader, Table, VMerge,
};
pub use error::{Error, Result};
