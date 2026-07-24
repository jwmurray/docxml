//! The [`Section`] handle: page geometry (size and margins) read and written through the
//! `w:sectPr` section-properties element.

use crate::xml::{NodeId, XmlTree};

use super::{
    Document, Length, LineNumberRestart, LineNumbering, PartId, is_wml_element,
    ordered_insert_index, rank_in,
};

/// Canonical `w:sectPr` child order (ECMA-376 §17.6.17, `CT_SectPr` — the
/// `EG_HdrFtrReferences` group followed by the `EG_SectPrContents` sequence), local names
/// only. New properties are inserted to keep `w:sectPr`'s children in this order so the
/// output is schema-valid: header/footer references come first, then `pgSz`, then `pgMar`.
/// Unlisted children rank last and stay after authored properties.
const SECTPR_ORDER: &[&str] = &[
    "headerReference",
    "footerReference",
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "formProt",
    "vAlign",
    "noEndnote",
    "titlePg",
    "textDirection",
    "bidi",
    "rtlGutter",
    "docGrid",
    "printerSettings",
    "sectPrChange",
];

/// A lightweight handle to a `w:sectPr` section-properties element.
///
/// A section groups page-level settings (size, margins, header/footer references). Like
/// the other handles, `Section` is `Copy` and borrows nothing — an arena node id plus its
/// part id. Pass a [`Document`] back to it to read or edit.
///
/// # What counts as a section
///
/// Following python-docx, [`Document::sections`] returns every `w:sectPr` in document
/// order: one for each `w:p` whose `w:pPr` carries a `w:sectPr` (which marks the *end* of
/// a section), plus the body-trailing `w:sectPr`. A single-section document has just the
/// body-trailing one.
///
/// # Page geometry
///
/// [`page_width`](Self::page_width) / [`page_height`](Self::page_height) read `w:pgSz`, and
/// the four margin accessors read `w:pgMar`; all are stored in twips and exposed as
/// [`Length`]. A missing element or attribute reads as `None`; the setters create the
/// element (in schema order) and attribute as needed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Section {
    part: PartId,
    node: NodeId,
}

impl Section {
    /// Wrap a known-`w:sectPr` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Section { part, node }
    }

    /// The section's underlying `w:sectPr` tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The part this section lives in (used by header/footer resolution).
    pub(crate) fn part(&self) -> PartId {
        self.part
    }

    /// The page width (`w:pgSz/@w:w`), or `None` when unset.
    pub fn page_width(&self, doc: &Document) -> Option<Length> {
        self.read_len(doc, "pgSz", "w")
    }

    /// The page height (`w:pgSz/@w:h`), or `None` when unset.
    pub fn page_height(&self, doc: &Document) -> Option<Length> {
        self.read_len(doc, "pgSz", "h")
    }

    /// The left page margin (`w:pgMar/@w:left`), or `None` when unset.
    pub fn left_margin(&self, doc: &Document) -> Option<Length> {
        self.read_len(doc, "pgMar", "left")
    }

    /// The right page margin (`w:pgMar/@w:right`), or `None` when unset.
    pub fn right_margin(&self, doc: &Document) -> Option<Length> {
        self.read_len(doc, "pgMar", "right")
    }

    /// The top page margin (`w:pgMar/@w:top`), or `None` when unset.
    pub fn top_margin(&self, doc: &Document) -> Option<Length> {
        self.read_len(doc, "pgMar", "top")
    }

    /// The bottom page margin (`w:pgMar/@w:bottom`), or `None` when unset.
    pub fn bottom_margin(&self, doc: &Document) -> Option<Length> {
        self.read_len(doc, "pgMar", "bottom")
    }

    /// Set the page width (`w:pgSz/@w:w`), creating `w:pgSz` if needed.
    pub fn set_page_width(&self, doc: &mut Document, len: Length) -> Section {
        self.set_len(doc, "pgSz", "w", len);
        *self
    }

    /// Set the page height (`w:pgSz/@w:h`), creating `w:pgSz` if needed.
    pub fn set_page_height(&self, doc: &mut Document, len: Length) -> Section {
        self.set_len(doc, "pgSz", "h", len);
        *self
    }

    /// Set the left page margin (`w:pgMar/@w:left`), creating `w:pgMar` if needed.
    pub fn set_left_margin(&self, doc: &mut Document, len: Length) -> Section {
        self.set_len(doc, "pgMar", "left", len);
        *self
    }

    /// Set the right page margin (`w:pgMar/@w:right`), creating `w:pgMar` if needed.
    pub fn set_right_margin(&self, doc: &mut Document, len: Length) -> Section {
        self.set_len(doc, "pgMar", "right", len);
        *self
    }

    /// Set the top page margin (`w:pgMar/@w:top`), creating `w:pgMar` if needed.
    pub fn set_top_margin(&self, doc: &mut Document, len: Length) -> Section {
        self.set_len(doc, "pgMar", "top", len);
        *self
    }

    /// Set the bottom page margin (`w:pgMar/@w:bottom`), creating `w:pgMar` if needed.
    pub fn set_bottom_margin(&self, doc: &mut Document, len: Length) -> Section {
        self.set_len(doc, "pgMar", "bottom", len);
        *self
    }

    /// Whether this section has a distinct first-page header/footer (`w:sectPr/w:titlePg`).
    ///
    /// `w:titlePg` is an on/off property (`CT_OnOff`, like `w:keepNext`): present and not
    /// explicitly `w:val="0"`/`"false"` means on. When set, Word uses the `"first"`-type
    /// header/footer references on the section's first page instead of the `"default"` ones.
    /// A header of type [`First`](crate::HeaderFooterType::First) is not shown unless this is
    /// also set — see [`create_header`](Self::create_header).
    pub fn different_first_page(&self, doc: &Document) -> bool {
        self.sect_toggle(doc, "titlePg")
    }

    /// Set whether this section has a distinct first-page header/footer
    /// (`w:sectPr/w:titlePg`). `w:titlePg` sits in `CT_SectPr` order (ECMA-376 §17.6.17,
    /// present in [`SECTPR_ORDER`]); on ensures a bare element, off removes it.
    pub fn set_different_first_page(&self, doc: &mut Document, on: bool) -> Section {
        self.set_sect_toggle(doc, "titlePg", on);
        *self
    }

    /// This section's line numbering (`w:sectPr/w:lnNumType`), or `None` when unset.
    ///
    /// Reads `w:countBy`, `w:start`, `w:distance` (as a twips [`Length`]), and `w:restart`.
    /// Schema defaults fill in for absent attributes: `w:start` defaults to `1`, `w:restart`
    /// to [`NewPage`](LineNumberRestart::NewPage), and an absent `w:countBy` reads as `0`.
    pub fn line_numbering(&self, doc: &Document) -> Option<LineNumbering> {
        let tree = doc.tree(self.part);
        let el = self.sect_child(tree, "lnNumType")?;
        let count_by = tree
            .attr(el, &doc.qn(self.part, "countBy"))
            .and_then(|v| v.trim().parse::<u32>().ok())
            .unwrap_or(0);
        let start = tree
            .attr(el, &doc.qn(self.part, "start"))
            .and_then(|v| v.trim().parse::<u32>().ok())
            .unwrap_or(1);
        let distance = tree
            .attr(el, &doc.qn(self.part, "distance"))
            .and_then(Length::from_twips_str);
        let restart = tree
            .attr(el, &doc.qn(self.part, "restart"))
            .and_then(LineNumberRestart::from_val)
            .unwrap_or(LineNumberRestart::NewPage);
        Some(LineNumbering {
            count_by,
            start,
            distance,
            restart,
        })
    }

    /// Set this section's line numbering (`w:sectPr/w:lnNumType`), creating the element in
    /// `CT_SectPr` order if absent.
    ///
    /// `w:lnNumType` sits in `EG_SectPrContents` after `w:pgBorders` and before
    /// `w:pgNumType` (ECMA-376 §17.6.17, `CT_SectPr`; present in [`SECTPR_ORDER`]). The
    /// attributes are written in `CT_LineNumber` order (§17.6.8: `w:countBy`, `w:start`,
    /// `w:distance`, `w:restart`). `w:countBy` is always written; attributes matching their
    /// schema default are omitted — `w:start` when it is `1`, `w:distance` when `None`, and
    /// `w:restart` when [`NewPage`](LineNumberRestart::NewPage). Any previously written values
    /// are cleared first so a re-set never leaves a stale attribute.
    pub fn set_line_numbering(&self, doc: &mut Document, ln: LineNumbering) -> Section {
        let countby_attr = doc.qn(self.part, "countBy");
        let start_attr = doc.qn(self.part, "start");
        let distance_attr = doc.qn(self.part, "distance");
        let restart_attr = doc.qn(self.part, "restart");
        let el = self.ensure_sect_child(doc, "lnNumType");
        let tree = doc.tree_mut(self.part);
        // Clear the managed attributes so a re-set never leaves a stale value behind.
        tree.remove_attr(el, &countby_attr);
        tree.remove_attr(el, &start_attr);
        tree.remove_attr(el, &distance_attr);
        tree.remove_attr(el, &restart_attr);
        // CT_LineNumber attribute order: countBy, start, distance, restart.
        tree.set_attr(el, countby_attr, ln.count_by.to_string());
        if ln.start != 1 {
            tree.set_attr(el, start_attr, ln.start.to_string());
        }
        if let Some(distance) = ln.distance {
            tree.set_attr(el, distance_attr, distance.to_twips_string());
        }
        if ln.restart != LineNumberRestart::NewPage {
            tree.set_attr(el, restart_attr, ln.restart.to_val());
        }
        *self
    }

    /// Remove this section's line numbering, deleting `w:sectPr/w:lnNumType`.
    pub fn clear_line_numbering(&self, doc: &mut Document) {
        if let Some(el) = self.sect_child(doc.tree(self.part), "lnNumType") {
            doc.tree_mut(self.part).remove_from_parent(el);
        }
    }

    /// A direct `w:sectPr` child with the given WML local name, if present.
    pub(crate) fn sect_child(&self, tree: &XmlTree, local: &str) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, local))
    }

    /// The index within this `w:sectPr` at which a new child named `local` should be
    /// inserted to keep the children in `CT_SectPr` schema order.
    ///
    /// Unlike [`ensure_sect_child`](Self::ensure_sect_child), this never reuses an existing
    /// child — header/footer references repeat (one per `w:type`), so their creation always
    /// inserts a fresh element. Used by header/footer reference creation in `header.rs`.
    pub(crate) fn sect_insert_index(&self, tree: &XmlTree, local: &str) -> usize {
        ordered_insert_index(tree, self.node, rank_in(SECTPR_ORDER, local), SECTPR_ORDER)
    }

    /// Read an on/off `w:sectPr` toggle child (`w:titlePg`, …): present and not
    /// `w:val="0"/"false"` is on. Mirrors [`Paragraph`](crate::Paragraph)'s `w:pPr` toggles.
    fn sect_toggle(&self, doc: &Document, local: &str) -> bool {
        let tree = doc.tree(self.part);
        let Some(el) = self.sect_child(tree, local) else {
            return false;
        };
        match tree.attr(el, &doc.qn(self.part, "val")) {
            Some(v) => !matches!(v, "0" | "false"),
            None => true,
        }
    }

    /// Set or clear an on/off `w:sectPr` toggle child. On ensures a bare element (clearing an
    /// explicit `w:val="0"/"false"`); off removes it.
    fn set_sect_toggle(&self, doc: &mut Document, local: &str, on: bool) {
        if on {
            let el = self.ensure_sect_child(doc, local);
            let val = doc.qn(self.part, "val");
            let tree = doc.tree_mut(self.part);
            if let Some(v) = tree.attr(el, &val) {
                if matches!(v, "0" | "false") {
                    tree.remove_attr(el, &val);
                }
            }
        } else if let Some(el) = self.sect_child(doc.tree(self.part), local) {
            doc.tree_mut(self.part).remove_from_parent(el);
        }
    }

    /// Read a twips-valued attribute `attr_local` of the direct `w:sectPr` child element
    /// `elem_local`, as a [`Length`].
    fn read_len(&self, doc: &Document, elem_local: &str, attr_local: &str) -> Option<Length> {
        let tree = doc.tree(self.part);
        let el = self.sect_child(tree, elem_local)?;
        let val = tree.attr(el, &doc.qn(self.part, attr_local))?;
        Length::from_twips_str(val)
    }

    /// Set a twips-valued attribute `attr_local` on the direct `w:sectPr` child element
    /// `elem_local`, creating the element in canonical schema order if absent.
    fn set_len(&self, doc: &mut Document, elem_local: &str, attr_local: &str, len: Length) {
        let attr = doc.qn(self.part, attr_local);
        let el = self.ensure_sect_child(doc, elem_local);
        doc.tree_mut(self.part)
            .set_attr(el, attr, len.to_twips_string());
    }

    /// A direct `w:sectPr` child with the given local name, creating it (in canonical
    /// schema order) if absent.
    fn ensure_sect_child(&self, doc: &mut Document, local: &str) -> NodeId {
        if let Some(existing) = self.sect_child(doc.tree(self.part), local) {
            return existing;
        }
        let name = doc.qn(self.part, local);
        let index = ordered_insert_index(
            doc.tree(self.part),
            self.node,
            rank_in(SECTPR_ORDER, local),
            SECTPR_ORDER,
        );
        let el = doc.tree_mut(self.part).create_element(name);
        doc.tree_mut(self.part).insert_child(self.node, index, el);
        el
    }
}

impl Document {
    /// The document's sections, in document order.
    ///
    /// Matches python-docx's `Document.sections`: every `w:sectPr` in the body — one per
    /// `w:p` whose `w:pPr` contains a `w:sectPr` (each marking the end of a section), plus
    /// the body-trailing `w:sectPr`. The blank template and most single-section documents
    /// therefore report exactly one section.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let doc = Document::new();
    /// assert_eq!(doc.sections().len(), 1);
    /// ```
    pub fn sections(&self) -> Vec<Section> {
        let tree = self.tree(PartId::MAIN);
        let root = tree.root();
        let Some(body) = tree
            .children(root)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "body"))
        else {
            return Vec::new();
        };

        let mut sections = Vec::new();
        for &child in tree.children(body) {
            if is_wml_element(tree, child, "p") {
                // A paragraph-nested w:pPr/w:sectPr marks the end of a section.
                if let Some(ppr) = tree
                    .children(child)
                    .iter()
                    .copied()
                    .find(|&c| is_wml_element(tree, c, "pPr"))
                {
                    if let Some(sect) = tree
                        .children(ppr)
                        .iter()
                        .copied()
                        .find(|&c| is_wml_element(tree, c, "sectPr"))
                    {
                        sections.push(Section::from_node(PartId::MAIN, sect));
                    }
                }
            } else if is_wml_element(tree, child, "sectPr") {
                // The body-trailing section properties.
                sections.push(Section::from_node(PartId::MAIN, child));
            }
        }
        sections
    }
}
