//! The [`Run`] handle and character-formatting accessors.

use crate::xml::{NodeId, XmlTree};

use super::paragraph::append_run_text;
use super::{
    BreakType, Document, PartId, Pt, RgbColor, is_wml_element, needs_space_preserve,
    ordered_insert_index, rank_in,
};

/// Canonical `w:rPr` child order (ECMA-376 §17.3.2.28, `CT_RPr` sequence), local names
/// only. New properties are inserted to keep `w:rPr`'s children in this order so the
/// output is schema-valid. Note the practically important subsequence this milestone
/// authors: `rStyle`, `rFonts`, `b`/`bCs`, `i`/`iCs`, then `color`, then `sz`/`szCs`,
/// then `u` — color precedes size, and underline follows size (verified against the
/// python-docx-generated `tests/fixtures/basic.docx`, whose colored/sized run serializes
/// `<w:color .../><w:sz .../>`). Unlisted children rank last and stay after authored
/// properties.
const RPR_ORDER: &[&str] = &[
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
];

/// A lightweight handle to a `w:r` run — a contiguous span of text with uniform
/// character formatting.
///
/// Like [`Paragraph`](super::Paragraph), `Run` is `Copy` and borrows nothing. The
/// formatting setters return the run so calls chain:
/// `run.bold(&mut doc, true).italic(&mut doc, true)`.
///
/// # Direct properties only
///
/// Every getter reads only the run's *direct* `w:rPr` — it does not resolve inheritance
/// from the run's paragraph style, the `docDefaults`, or a linked character style. A run
/// that renders bold in Word purely because its style is bold will report
/// [`is_bold`](Self::is_bold) as `false` here. Style-inheritance resolution is a later
/// milestone.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Run {
    part: PartId,
    node: NodeId,
}

impl Run {
    /// Wrap a known-`w:r` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Run { part, node }
    }

    /// The run's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The run's text: `w:t` verbatim, `w:tab` as a tab, `w:br` / `w:cr` as a newline.
    pub fn text(&self, doc: &Document) -> String {
        let mut out = String::new();
        append_run_text(doc.tree(self.part), self.node, &mut out);
        out
    }

    /// Replace the run's text with a single `w:t`.
    ///
    /// Existing text-bearing children (`w:t`, `w:tab`, `w:br`, `w:cr`) are removed and
    /// one `w:t` carrying `text` is appended after the run properties, with
    /// `xml:space="preserve"` when `text` has leading or trailing whitespace. The run's
    /// `w:rPr` (bold, italic, …) is left untouched.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("");
    /// let r = p.add_run(&mut doc, "old");
    /// r.set_text(&mut doc, "new");
    /// assert_eq!(r.text(&doc), "new");
    /// ```
    pub fn set_text(&self, doc: &mut Document, text: &str) -> Run {
        let t_name = doc.qn(self.part, "t");
        let tree = doc.tree_mut(self.part);

        let content: Vec<NodeId> = tree
            .children(self.node)
            .iter()
            .copied()
            .filter(|&c| is_content_child(tree, c))
            .collect();
        for child in content {
            tree.remove_from_parent(child);
        }

        let t = tree.create_element(t_name);
        let text_node = tree.create_text(text);
        tree.append_child(t, text_node);
        tree.append_child(self.node, t);
        if needs_space_preserve(text) {
            tree.set_attr(t, "xml:space", "preserve");
        }
        *self
    }

    /// Whether the run is bold.
    ///
    /// True when the direct `w:rPr` carries a `w:b` whose `w:val` is not `"0"` /
    /// `"false"` (a bare `w:b`, as Word writes, means on). Reads direct properties only —
    /// see the [type docs](Self#direct-properties-only).
    pub fn is_bold(&self, doc: &Document) -> bool {
        self.has_toggle(doc, "b")
    }

    /// Whether the run is italic (see [`is_bold`](Self::is_bold) for the `w:val` rule and
    /// the direct-properties-only caveat).
    pub fn is_italic(&self, doc: &Document) -> bool {
        self.has_toggle(doc, "i")
    }

    /// Whether the run is underlined.
    ///
    /// True when the direct `w:rPr` carries a `w:u` whose `w:val` is present and not
    /// `"none"` / `"0"` / `"false"` (a bare `w:u` reads as underlined). Direct properties
    /// only — see the [type docs](Self#direct-properties-only).
    pub fn is_underlined(&self, doc: &Document) -> bool {
        let tree = doc.tree(self.part);
        let Some(u) = self.rpr_child(tree, "u") else {
            return false;
        };
        match tree.attr(u, &doc.qn(self.part, "val")) {
            Some(v) => !matches!(v, "none" | "0" | "false"),
            None => true,
        }
    }

    /// The run's font size, if `w:rPr/w:sz` is set (direct properties only).
    pub fn size(&self, doc: &Document) -> Option<Pt> {
        let tree = doc.tree(self.part);
        let sz = self.rpr_child(tree, "sz")?;
        Pt::from_half_points_str(tree.attr(sz, &doc.qn(self.part, "val"))?)
    }

    /// The run's color, if `w:rPr/w:color` is set to a concrete value (`w:val="auto"`
    /// and unparsable values read as `None`; direct properties only).
    pub fn color(&self, doc: &Document) -> Option<RgbColor> {
        let tree = doc.tree(self.part);
        let color = self.rpr_child(tree, "color")?;
        RgbColor::from_hex(tree.attr(color, &doc.qn(self.part, "val"))?)
    }

    /// The run's font (typeface) name, read from `w:rPr/w:rFonts` `w:ascii`
    /// (direct properties only).
    pub fn font(&self, doc: &Document) -> Option<String> {
        let tree = doc.tree(self.part);
        let rfonts = self.rpr_child(tree, "rFonts")?;
        tree.attr(rfonts, &doc.qn(self.part, "ascii"))
            .map(str::to_owned)
    }

    /// Turn bold on or off.
    ///
    /// On adds a bare `w:b` to `w:rPr` (creating `w:rPr` as the run's first child if
    /// needed); off removes any `w:b`. This milestone does not emit
    /// `w:b w:val="false"` — removal is the off state.
    pub fn bold(&self, doc: &mut Document, on: bool) -> Run {
        self.set_toggle(doc, "b", on);
        *self
    }

    /// Turn italic on or off (see [`bold`](Self::bold) for the representation).
    pub fn italic(&self, doc: &mut Document, on: bool) -> Run {
        self.set_toggle(doc, "i", on);
        *self
    }

    /// Turn underline on or off.
    ///
    /// On sets `w:rPr/w:u w:val="single"`; off removes any `w:u`. Removal is the off
    /// state (no `w:u w:val="none"` is emitted).
    pub fn underline(&self, doc: &mut Document, on: bool) -> Run {
        if on {
            let u = self.ensure_rpr_child(doc, "u");
            let val = doc.qn(self.part, "val");
            doc.tree_mut(self.part).set_attr(u, val, "single");
        } else {
            self.remove_rpr_child(doc, "u");
        }
        *self
    }

    /// Set the font size. Writes both `w:sz` and `w:szCs` (complex-script size) with the
    /// value in half-points, matching python-docx.
    pub fn set_size(&self, doc: &mut Document, size: Pt) -> Run {
        let hp = size.to_half_points_string();
        let val = doc.qn(self.part, "val");
        let sz = self.ensure_rpr_child(doc, "sz");
        doc.tree_mut(self.part)
            .set_attr(sz, val.clone(), hp.clone());
        let sz_cs = self.ensure_rpr_child(doc, "szCs");
        doc.tree_mut(self.part).set_attr(sz_cs, val, hp);
        *self
    }

    /// Set the run color (`w:color w:val` as six uppercase hex digits).
    pub fn set_color(&self, doc: &mut Document, color: RgbColor) -> Run {
        let val = doc.qn(self.part, "val");
        let el = self.ensure_rpr_child(doc, "color");
        doc.tree_mut(self.part).set_attr(el, val, color.to_hex());
        *self
    }

    /// Set the run font (typeface). Writes `w:rFonts` `w:ascii` and `w:hAnsi` to `name`,
    /// the two attributes that cover Latin text.
    pub fn set_font(&self, doc: &mut Document, name: &str) -> Run {
        let ascii = doc.qn(self.part, "ascii");
        let hansi = doc.qn(self.part, "hAnsi");
        let el = self.ensure_rpr_child(doc, "rFonts");
        doc.tree_mut(self.part).set_attr(el, ascii, name);
        doc.tree_mut(self.part).set_attr(el, hansi, name);
        *self
    }

    /// Whether the run is small-caps (`w:rPr/w:smallCaps`).
    ///
    /// Same toggle rule as [`is_bold`](Self::is_bold): present and not `w:val="0"/"false"`
    /// is on. Direct properties only — see the [type docs](Self#direct-properties-only).
    pub fn small_caps(&self, doc: &Document) -> bool {
        self.has_toggle(doc, "smallCaps")
    }

    /// Turn small-caps on or off (`w:rPr/w:smallCaps`). On adds a bare `w:smallCaps`; off
    /// removes it — the same representation as [`bold`](Self::bold).
    pub fn set_small_caps(&self, doc: &mut Document, on: bool) -> Run {
        self.set_toggle(doc, "smallCaps", on);
        *self
    }

    /// Whether the run is all-caps (`w:rPr/w:caps`).
    ///
    /// Same toggle rule as [`is_bold`](Self::is_bold). Direct properties only — see the
    /// [type docs](Self#direct-properties-only).
    pub fn all_caps(&self, doc: &Document) -> bool {
        self.has_toggle(doc, "caps")
    }

    /// Turn all-caps on or off (`w:rPr/w:caps`). On adds a bare `w:caps`; off removes it —
    /// the same representation as [`bold`](Self::bold).
    pub fn set_all_caps(&self, doc: &mut Document, on: bool) -> Run {
        self.set_toggle(doc, "caps", on);
        *self
    }

    /// Whether the run is hidden text (`w:rPr/w:vanish`).
    ///
    /// Hidden text is present in the document but not shown (or printed) by default; Word
    /// uses it for table-of-authorities / table-of-contents marker text, index entries, and
    /// the like. Same toggle rule as [`is_bold`](Self::is_bold): present and not
    /// `w:val="0"/"false"` is on. Direct properties only — see the [type docs](Self#direct-properties-only).
    ///
    /// Hidden runs still contribute their text to [`text`](Self::text) and
    /// [`Paragraph::text`](super::Paragraph::text) — the run is hidden, not absent.
    pub fn vanish(&self, doc: &Document) -> bool {
        self.has_toggle(doc, "vanish")
    }

    /// Turn hidden text on or off (`w:rPr/w:vanish`). On adds a bare `w:vanish`; off removes
    /// it — the same representation as [`bold`](Self::bold).
    pub fn set_vanish(&self, doc: &mut Document, on: bool) -> Run {
        self.set_toggle(doc, "vanish", on);
        *self
    }

    /// Append a break (`w:br`) to the run.
    ///
    /// A [`BreakType::Page`] / [`BreakType::Column`] writes `w:br w:type="page"` /
    /// `"column"`; [`BreakType::Line`] writes a bare `w:br`. The break is appended after the
    /// run's existing content, and reads back as a newline in [`text`](Self::text).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, BreakType};
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("");
    /// let r = p.add_run(&mut doc, "line one");
    /// r.add_break(&mut doc, BreakType::Line);
    /// assert!(r.text(&doc).contains('\n'));
    /// ```
    pub fn add_break(&self, doc: &mut Document, kind: BreakType) -> Run {
        let br_name = doc.qn(self.part, "br");
        let type_attr = doc.qn(self.part, "type");
        let tree = doc.tree_mut(self.part);
        let br = tree.create_element(br_name);
        tree.append_child(self.node, br);
        if let Some(t) = kind.type_val() {
            tree.set_attr(br, type_attr, t);
        }
        *self
    }

    /// The run's `w:rPr`, if present.
    fn rpr(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "rPr"))
    }

    /// The run's `w:rPr`, creating it as the first child if absent (schema requires
    /// `w:rPr` before the run content).
    fn ensure_rpr(&self, doc: &mut Document) -> NodeId {
        if let Some(rpr) = self.rpr(doc.tree(self.part)) {
            return rpr;
        }
        let name = doc.qn(self.part, "rPr");
        let tree = doc.tree_mut(self.part);
        let rpr = tree.create_element(name);
        tree.insert_child(self.node, 0, rpr);
        rpr
    }

    /// A direct `w:rPr` child with the given WML local name, if present.
    fn rpr_child(&self, tree: &XmlTree, local: &str) -> Option<NodeId> {
        let rpr = self.rpr(tree)?;
        tree.children(rpr)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, local))
    }

    /// A direct `w:rPr` child with the given local name, creating it (in canonical
    /// schema order) if absent. Creates `w:rPr` first if needed.
    fn ensure_rpr_child(&self, doc: &mut Document, local: &str) -> NodeId {
        let rpr = self.ensure_rpr(doc);
        if let Some(existing) = {
            let tree = doc.tree(self.part);
            tree.children(rpr)
                .iter()
                .copied()
                .find(|&c| is_wml_element(tree, c, local))
        } {
            return existing;
        }
        let name = doc.qn(self.part, local);
        let index = ordered_insert_index(
            doc.tree(self.part),
            rpr,
            rank_in(RPR_ORDER, local),
            RPR_ORDER,
        );
        let el = doc.tree_mut(self.part).create_element(name);
        doc.tree_mut(self.part).insert_child(rpr, index, el);
        el
    }

    /// Remove a direct `w:rPr` child by local name, if present.
    fn remove_rpr_child(&self, doc: &mut Document, local: &str) {
        if let Some(el) = self.rpr_child(doc.tree(self.part), local) {
            doc.tree_mut(self.part).remove_from_parent(el);
        }
    }

    /// Read a boolean toggle property (`w:b`, `w:i`) from `w:rPr`.
    fn has_toggle(&self, doc: &Document, local: &str) -> bool {
        let tree = doc.tree(self.part);
        let Some(el) = self.rpr_child(tree, local) else {
            return false;
        };
        match tree.attr(el, &doc.qn(self.part, "val")) {
            Some(v) => !matches!(v, "0" | "false"),
            None => true,
        }
    }

    /// Set or clear a boolean toggle property (`w:b`, `w:i`) in `w:rPr`.
    fn set_toggle(&self, doc: &mut Document, local: &str, on: bool) {
        if on {
            let el = self.ensure_rpr_child(doc, local);
            // Clear an explicit `w:val="0"/"false"` so a bare element reads on.
            let val_name = doc.qn(self.part, "val");
            let tree = doc.tree_mut(self.part);
            if let Some(v) = tree.attr(el, &val_name) {
                if matches!(v, "0" | "false") {
                    tree.remove_attr(el, &val_name);
                }
            }
        } else {
            self.remove_rpr_child(doc, local);
        }
    }
}

/// Whether a run child carries text content that [`Run::set_text`] should replace.
fn is_content_child(tree: &XmlTree, id: NodeId) -> bool {
    is_wml_element(tree, id, "t")
        || is_wml_element(tree, id, "tab")
        || is_wml_element(tree, id, "br")
        || is_wml_element(tree, id, "cr")
}
