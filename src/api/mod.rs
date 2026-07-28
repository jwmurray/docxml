//! Typed document API: the ergonomic layer over the lossless XML tree.
//!
//! This is python-docx's proxy pattern, Rust-flavored. [`Document`] owns the
//! [`Package`](crate::opc::Package) and the parsed main-document tree; [`Paragraph`]
//! and [`Run`] are lightweight `Copy` handles — an arena [`NodeId`](crate::xml::NodeId)
//! with phantom typing — that borrow nothing. Every read goes through `&Document` and
//! every mutation through `&mut Document`, so the borrow checker stays out of the way.
//!
//! Element matching is namespace-correct: elements are identified by their resolved
//! WordprocessingML namespace URI (transitional *or* strict) plus local name, never by
//! a literal `w:` prefix. New elements are created with whatever prefix the document
//! root maps to the main URI (conventionally `w:`).
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

mod document;
mod field;
mod header;
mod notes;
mod numbering;
mod paragraph;
mod picture;
mod revisions;
mod run;
mod section;
mod settings;
mod style;
mod table;
mod units;

pub use document::Document;
pub use header::{HeaderFooter, HeaderFooterType};
pub use notes::{Endnote, Footnote};
pub use numbering::NumberFormat;
pub use paragraph::{HyperlinkInfo, Paragraph};
pub use picture::Picture;
pub use revisions::{Comment, Revision, RevisionKind};
pub use run::Run;
pub use section::Section;
pub use style::{Style, StyleType};
pub use table::{Cell, Row, Table, VMerge};
pub use units::{
    Alignment, BorderEdge, BorderStyle, BreakType, DocGrid, DocGridType, FrameAnchor, FrameOptions,
    FrameWrap, Length, LineNumberRestart, LineNumbering, LineSpacing, PageBorderDisplay,
    PageBorderOffset, PageBorders, PageNumbering, Pt, RgbColor, TabAlignment, TabLeader,
};

use crate::xml::{NodeId, XmlTree};

/// Index of a parsed part in a [`Document`]'s `parsed` vector.
///
/// Index `0` is always the main document part (parsed eagerly on open/new). Header and
/// footer parts are parsed lazily on first access and appended, each taking the next id.
/// Every handle ([`Paragraph`], [`Run`], [`Table`], [`Row`], [`Cell`]) carries the
/// `PartId` of the part it lives in, so reads and mutations route to the right tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct PartId(pub(crate) u16);

impl PartId {
    /// The main document part — always parsed, always index `0`.
    pub(crate) const MAIN: PartId = PartId(0);
}

/// The two WordprocessingML main namespace URIs: transitional (the one Word writes)
/// and strict (ISO/IEC 29500 strict).
const WML_MAIN_URIS: [&str; 2] = [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
];

/// Split a qualified name into its optional prefix and local part (`"w:p"` →
/// `(Some("w"), "p")`, `"body"` → `(None, "body")`).
fn split_qname(qname: &str) -> (Option<&str>, &str) {
    match qname.split_once(':') {
        Some((prefix, local)) => (Some(prefix), local),
        None => (None, qname),
    }
}

/// True when `id` is an element in a WordprocessingML main namespace (transitional or
/// strict) whose local name is `local`.
///
/// The element's prefix is resolved to a URI through [`XmlTree::namespace_uri`] rather
/// than matched literally, so both `w:`-prefixed and default-namespace documents, in
/// either the transitional or strict namespace, are handled.
fn is_wml_element(tree: &XmlTree, id: NodeId, local: &str) -> bool {
    is_element_in(tree, id, &WML_MAIN_URIS, local)
}

/// True when `id` is an element whose local name is `local` and whose resolved namespace
/// URI is one of `uris`.
///
/// Generalizes [`is_wml_element`] to any namespace: the element's prefix is resolved to a
/// URI through [`XmlTree::namespace_uri`] rather than matched literally, so DrawingML,
/// content-types, and relationships elements match regardless of the prefix a given
/// document happens to use (transitional or strict URI).
fn is_element_in(tree: &XmlTree, id: NodeId, uris: &[&str], local: &str) -> bool {
    let Some(name) = tree.name(id) else {
        return false;
    };
    let (prefix, local_name) = split_qname(name);
    if local_name != local {
        return false;
    }
    match tree.namespace_uri(id, prefix) {
        Some(uri) => uris.contains(&uri),
        None => false,
    }
}

/// Whether a `w:t` needs `xml:space="preserve"` to keep its text intact: true when the
/// text has leading or trailing whitespace (which XML would otherwise be free to trim).
fn needs_space_preserve(text: &str) -> bool {
    text.starts_with(|c: char| c.is_whitespace()) || text.ends_with(|c: char| c.is_whitespace())
}

/// Schema-order rank of a child by local name, from a canonical order list: the child's
/// position in `order`, or [`u32::MAX`] when it is not listed (unknown / pass-through
/// content sorts last, so authored properties always slot in ahead of it).
fn rank_in(order: &[&str], local: &str) -> u32 {
    order
        .iter()
        .position(|&n| n == local)
        .map(|i| i as u32)
        .unwrap_or(u32::MAX)
}

/// Index at which to insert a new child of rank `new_rank` into `parent` so its children
/// stay in ascending schema order: before the first existing child that ranks after it.
/// `order` is the canonical local-name sequence for `parent`'s content model.
fn ordered_insert_index(tree: &XmlTree, parent: NodeId, new_rank: u32, order: &[&str]) -> usize {
    let children = tree.children(parent);
    children
        .iter()
        .position(|&c| {
            let rank = match tree.name(c) {
                Some(name) => rank_in(order, split_qname(name).1),
                None => u32::MAX,
            };
            rank > new_rank
        })
        .unwrap_or(children.len())
}

/// Read a `CT_Border` edge child (`local`) of `parent` into a [`BorderEdge`], or `None`
/// when that child is absent or its `w:val` is a style this API does not model (the
/// [`BorderStyle`] enum is closed). `w:sz`/`w:space` default to `0` when missing or
/// unparsable, and `w:color="auto"` (or an unparsable color) reads as `color` `None`.
///
/// Shared by paragraph borders (`w:pBdr`), table borders (`w:tblBorders`), and cell
/// borders (`w:tcBorders`) — every `CT_Border` reads identically.
fn read_border_edge(
    doc: &Document,
    part: PartId,
    parent: NodeId,
    local: &str,
) -> Option<BorderEdge> {
    let tree = doc.tree(part);
    let el = tree
        .children(parent)
        .iter()
        .copied()
        .find(|&c| is_wml_element(tree, c, local))?;
    let style = BorderStyle::from_val(tree.attr(el, &doc.qn(part, "val"))?)?;
    let size = tree
        .attr(el, &doc.qn(part, "sz"))
        .and_then(|v| v.trim().parse::<u8>().ok())
        .unwrap_or(0);
    let space = tree
        .attr(el, &doc.qn(part, "space"))
        .and_then(|v| v.trim().parse::<u8>().ok())
        .unwrap_or(0);
    let color = tree
        .attr(el, &doc.qn(part, "color"))
        .and_then(RgbColor::from_hex);
    Some(BorderEdge {
        style,
        size,
        space,
        color,
    })
}

/// Create a `CT_Border` edge element `local` carrying the edge's `w:val`, `w:sz`,
/// `w:space`, and `w:color` (a `None` color writes `w:color="auto"`), and insert it into
/// `parent` at its position in schema `order`. Attribute order is `val, sz, space, color`.
///
/// Shared by paragraph, table, and cell border writes — every `CT_Border` writes
/// identically; only the parent element and its `order` differ.
fn insert_border_edge(
    doc: &mut Document,
    part: PartId,
    parent: NodeId,
    local: &str,
    edge: BorderEdge,
    order: &[&str],
) {
    let val_attr = doc.qn(part, "val");
    let sz_attr = doc.qn(part, "sz");
    let space_attr = doc.qn(part, "space");
    let color_attr = doc.qn(part, "color");
    let name = doc.qn(part, local);
    let index = ordered_insert_index(doc.tree(part), parent, rank_in(order, local), order);
    let tree = doc.tree_mut(part);
    let el = tree.create_element(name);
    // CT_Border attribute order: val, then sz, space, and color.
    tree.set_attr(el, val_attr, edge.style.to_val());
    tree.set_attr(el, sz_attr, edge.size.to_string());
    tree.set_attr(el, space_attr, edge.space.to_string());
    match edge.color {
        Some(color) => tree.set_attr(el, color_attr, color.to_hex()),
        None => tree.set_attr(el, color_attr, "auto"),
    }
    tree.insert_child(parent, index, el);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn split_qname_handles_prefix_and_bare_names() {
        assert_eq!(split_qname("w:p"), (Some("w"), "p"));
        assert_eq!(split_qname("body"), (None, "body"));
    }

    #[test]
    fn is_wml_matches_transitional_strict_and_default_namespaces() {
        // Transitional prefixed, strict prefixed, and a non-WML namespace.
        let xml = br#"<w:document
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:o="urn:other">
            <w:body/><s:p/><o:p/></w:document>"#;
        let tree = XmlTree::parse(xml).unwrap();
        let root = tree.root();
        assert!(is_wml_element(&tree, root, "document"));

        let body = tree.children_named(root, "w:body").next().unwrap();
        assert!(is_wml_element(&tree, body, "body"));
        assert!(!is_wml_element(&tree, body, "p")); // right namespace, wrong local name

        let strict_p = tree.children_named(root, "s:p").next().unwrap();
        assert!(is_wml_element(&tree, strict_p, "p")); // strict URI still matches

        let other_p = tree.children_named(root, "o:p").next().unwrap();
        assert!(!is_wml_element(&tree, other_p, "p")); // non-WML namespace does not match
    }

    #[test]
    fn is_wml_matches_when_main_is_the_default_namespace() {
        let xml = br#"<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><body/></document>"#;
        let tree = XmlTree::parse(xml).unwrap();
        let root = tree.root();
        assert!(is_wml_element(&tree, root, "document"));
        let body = tree.children(root)[0];
        assert!(is_wml_element(&tree, body, "body"));
    }

    #[test]
    fn space_preserve_only_for_edge_whitespace() {
        assert!(!needs_space_preserve("hello"));
        assert!(!needs_space_preserve("a b c"));
        assert!(needs_space_preserve(" leading"));
        assert!(needs_space_preserve("trailing "));
        assert!(needs_space_preserve("\ttab"));
    }
}
