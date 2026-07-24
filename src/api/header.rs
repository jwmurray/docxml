//! The [`HeaderFooter`] handle and the [`Section`] accessors that resolve a section's
//! header and footer parts.
//!
//! Headers and footers live in their own package parts (`word/header1.xml`,
//! `word/footer1.xml`, …), referenced from a section's `w:sectPr` by relationship id.
//! Resolving one is a three-step hop — `w:headerReference/@r:id` → the part-level
//! relationships (`word/_rels/document.xml.rels`) → the header/footer part — after which
//! the part is parsed lazily and cached in the [`Document`]. Because the returned
//! [`Paragraph`]s carry the header/footer part's id, the ordinary [`Paragraph`] / [`Run`]
//! read-and-edit API works on them unchanged.

use crate::xml::{NodeId, XmlTree};

use super::{Document, Paragraph, PartId, Section, is_wml_element, split_qname};

/// The relationships-namespace URIs — transitional (what Word writes) and strict — used
/// to identify the `r:id` attribute regardless of the prefix a document binds to them.
const REL_URIS: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
];

/// A lightweight handle to a header or footer part (`w:hdr` / `w:ftr`).
///
/// Obtained from [`Section::header`] / [`Section::footer`]. Like the other handles it is
/// `Copy` and borrows nothing — it carries the [`PartId`] of the lazily parsed header or
/// footer part, so its [`paragraphs`](Self::paragraphs) read and edit that part's tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderFooter {
    part: PartId,
}

impl HeaderFooter {
    /// Wrap a header/footer part id.
    pub(crate) fn from_part(part: PartId) -> Self {
        HeaderFooter { part }
    }

    /// The body-level paragraphs of the header/footer: the direct `w:p` children of the
    /// part's root `w:hdr` / `w:ftr`.
    ///
    /// The returned [`Paragraph`]s carry this part's id, so their text reads — and their
    /// runs edit — against the header/footer tree through the ordinary API. Paragraphs
    /// nested inside a header/footer table are not included, matching the body's
    /// `paragraphs` behavior.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let mut doc = docxml::Document::open("basic.docx")?;
    /// let section = doc.sections()[0];
    /// if let Some(header) = section.header(&mut doc) {
    ///     let text: String = header
    ///         .paragraphs(&doc)
    ///         .iter()
    ///         .map(|p| p.text(&doc))
    ///         .collect();
    ///     println!("{text}");
    /// }
    /// # Ok::<(), docxml::Error>(())
    /// ```
    pub fn paragraphs(&self, doc: &Document) -> Vec<Paragraph> {
        let tree = doc.tree(self.part);
        let root = tree.root();
        tree.children(root)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "p"))
            .map(|c| Paragraph::from_node(self.part, c))
            .collect()
    }
}

impl Section {
    /// The section's default header, if it references one.
    ///
    /// Resolves the `w:headerReference` of type `"default"` (the `"first"` and `"even"`
    /// reference types are intentionally ignored at this milestone) through the part-level
    /// relationships to a header part, parses it lazily, and caches it in `doc`. Returns
    /// `None` when the section has no default header reference or the target part cannot be
    /// resolved or parsed.
    ///
    /// Takes `&mut Document` because resolving may parse and cache a new part; parsing
    /// alone does **not** mark the document modified, so a read-only header access leaves
    /// every part byte-identical on save.
    ///
    /// Creating a header where none exists is out of scope for this milestone — it needs
    /// new content-type and relationship entries written — so this only ever *reads* an
    /// existing reference.
    pub fn header(&self, doc: &mut Document) -> Option<HeaderFooter> {
        self.hdr_ftr_ref(doc, "headerReference")
    }

    /// The section's default footer, if it references one. See [`header`](Self::header) for
    /// the resolution rules and the `"default"`-only / out-of-scope-creation caveats.
    pub fn footer(&self, doc: &mut Document) -> Option<HeaderFooter> {
        self.hdr_ftr_ref(doc, "footerReference")
    }

    /// Shared header/footer resolution: find the reference of type `"default"` under the
    /// `w:sectPr`, read its relationship id, resolve it to a part, and lazily parse it.
    fn hdr_ftr_ref(&self, doc: &mut Document, ref_local: &str) -> Option<HeaderFooter> {
        let part = self.part();
        // Find the default reference and read its r:id from the main tree.
        let r_id = {
            let tree = doc.tree(part);
            let type_attr = doc.qn(part, "type");
            let sect = self.node();
            tree.children(sect)
                .iter()
                .copied()
                .filter(|&c| is_wml_element(tree, c, ref_local))
                .find(|&c| tree.attr(c, &type_attr) == Some("default"))
                .and_then(|rf| rel_id_attr(tree, rf))
                .map(str::to_owned)
        }?;

        let source = doc.main_part_name().to_string();
        let target = doc.resolve_rel_target(&source, &r_id)?;
        let hf_part = doc.ensure_part(&target)?;
        Some(HeaderFooter::from_part(hf_part))
    }
}

/// The value of an element's relationship-id attribute — the attribute whose local name is
/// `id` and whose prefix resolves (via the in-scope namespace declarations) to a
/// relationships namespace URI, transitional or strict. This is the namespace-correct way
/// to read `r:id` without assuming the `r` prefix.
fn rel_id_attr(tree: &XmlTree, node: NodeId) -> Option<&str> {
    for (key, value) in tree.attrs(node) {
        let (prefix, local) = split_qname(key);
        if local != "id" {
            continue;
        }
        if let Some(uri) = tree.namespace_uri(node, prefix) {
            if REL_URIS.contains(&uri) {
                return Some(value.as_str());
            }
        }
    }
    None
}
