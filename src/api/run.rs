//! The [`Run`] handle and character-formatting accessors.

use std::collections::HashSet;

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
        let rpr = self.ensure_rpr(doc);
        set_size_in(doc, self.part, rpr, size);
        *self
    }

    /// Set the run color (`w:color w:val` as six uppercase hex digits).
    pub fn set_color(&self, doc: &mut Document, color: RgbColor) -> Run {
        let rpr = self.ensure_rpr(doc);
        set_color_in(doc, self.part, rpr, color);
        *self
    }

    /// Set the run font (typeface). Writes `w:rFonts` `w:ascii` and `w:hAnsi` to `name`,
    /// the two attributes that cover Latin text.
    pub fn set_font(&self, doc: &mut Document, name: &str) -> Run {
        let rpr = self.ensure_rpr(doc);
        set_font_in(doc, self.part, rpr, name);
        *self
    }

    /// The run's character style id, read from `w:rPr/w:rStyle` `w:val` (direct properties
    /// only). This is the internal styleId (e.g. `"Hyperlink"`), not the display name.
    pub fn style_id(&self, doc: &Document) -> Option<String> {
        let tree = doc.tree(self.part);
        let rstyle = self.rpr_child(tree, "rStyle")?;
        tree.attr(rstyle, &doc.qn(self.part, "val"))
            .map(str::to_owned)
    }

    /// Set the run's character style id (`w:rPr/w:rStyle w:val`).
    ///
    /// `w:rStyle` is the first `w:rPr` child in canonical order ([`RPR_ORDER`] position 0),
    /// so it slots ahead of any existing character properties. `id` is a character styleId
    /// defined in `word/styles.xml` (see [`Document::create_style`](super::Document::create_style)).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, StyleType};
    ///
    /// let mut doc = Document::new();
    /// doc.create_style("Emphatic", "Emphatic", StyleType::Character).set_bold(&mut doc, true);
    /// let p = doc.add_paragraph("");
    /// let r = p.add_run(&mut doc, "loud");
    /// r.set_style_id(&mut doc, "Emphatic");
    /// assert_eq!(r.style_id(&doc).as_deref(), Some("Emphatic"));
    /// ```
    pub fn set_style_id(&self, doc: &mut Document, id: &str) -> Run {
        let val = doc.qn(self.part, "val");
        let rstyle = self.ensure_rpr_child(doc, "rStyle");
        doc.tree_mut(self.part).set_attr(rstyle, val, id);
        *self
    }

    /// The run's *effective* bold state, resolving inheritance the way Word renders it.
    ///
    /// Unlike [`is_bold`](Self::is_bold) (which reads only the direct `w:rPr`), this walks
    /// the full resolution order and returns the first level that sets `w:b`:
    ///
    /// 1. the run's direct `w:rPr`;
    /// 2. the run's character style (`w:rStyle`) and its `w:basedOn` chain;
    /// 3. the containing paragraph's style (`w:pStyle`) and its `w:basedOn` chain;
    /// 4. `styles.xml`'s `w:docDefaults/w:rPrDefault`.
    ///
    /// # Toggle semantics (simplified)
    ///
    /// `w:b` is an OOXML *toggle* property, whose full ECMA-376 semantics XOR the settings
    /// found at every level (so bold-on a style plus bold-on direct formatting cancels to
    /// *off*). That is deliberately out of scope here: this returns the **first explicit
    /// setting** encountered in the order above (defaulting to `false`), which matches Word
    /// for the overwhelmingly common case where bold is set at exactly one level. `w:basedOn`
    /// walks are cycle-guarded with a seen-set.
    pub fn effective_bold(&self, doc: &mut Document) -> bool {
        self.effective_toggle(doc, "b")
    }

    /// The run's effective italic state — see [`effective_bold`](Self::effective_bold) for
    /// the resolution order and the simplified toggle semantics.
    pub fn effective_italic(&self, doc: &mut Document) -> bool {
        self.effective_toggle(doc, "i")
    }

    /// The run's effective font size, resolving through the same chain as
    /// [`effective_bold`](Self::effective_bold) (direct `w:rPr` → `w:rStyle` chain →
    /// paragraph `w:pStyle` chain → `w:docDefaults`). `None` when no level sets `w:sz`.
    pub fn effective_size(&self, doc: &mut Document) -> Option<Pt> {
        self.effective_value(doc, |tree, rpr, val| {
            let sz = rpr_child_in(tree, rpr, "sz")?;
            Pt::from_half_points_str(tree.attr(sz, val)?)
        })
    }

    /// The run's effective font (typeface) name, resolving through the same chain as
    /// [`effective_bold`](Self::effective_bold), reading `w:rFonts/@w:ascii`. `None` when no
    /// level sets an explicit `w:ascii` font (a theme font such as `w:asciiTheme` is not a
    /// concrete name and reads as absent here).
    pub fn effective_font(&self, doc: &mut Document) -> Option<String> {
        // `w:ascii` is read via the part's qualified attribute name; the extractor is
        // handed the `w:val` name, so it computes `w:ascii` from the same prefix.
        self.effective_value(doc, |tree, rpr, val| {
            // `val` is the part's qualified `w:val`; swap the local name for `ascii` to get
            // the same-prefix `w:ascii` attribute name.
            let ascii = format!("{}ascii", &val[..val.len() - "val".len()]);
            let rfonts = rpr_child_in(tree, rpr, "rFonts")?;
            tree.attr(rfonts, &ascii).map(str::to_owned)
        })
    }

    /// Resolve a toggle property through the inheritance chain (see
    /// [`effective_bold`](Self::effective_bold)).
    fn effective_toggle(&self, doc: &mut Document, local: &str) -> bool {
        self.effective_value(doc, |tree, rpr, val| read_toggle(tree, rpr, val, local))
            .unwrap_or(false)
    }

    /// The core resolution walk shared by every effective read. `extract` pulls the property
    /// out of a `w:rPr` given the tree it lives in and that part's `w:val` attribute name;
    /// the first level that yields `Some` wins.
    fn effective_value<T>(
        &self,
        doc: &mut Document,
        extract: impl Fn(&XmlTree, NodeId, &str) -> Option<T> + Copy,
    ) -> Option<T> {
        // Level 1 (direct) and the style ids are read from the run's own part.
        let val_self = doc.qn(self.part, "val");
        let (direct, rstyle_id, pstyle_id) = {
            let tree = doc.tree(self.part);
            let direct = self.rpr(tree).and_then(|rpr| extract(tree, rpr, &val_self));
            (
                direct,
                self.rstyle_id_in(tree, &val_self),
                self.paragraph_pstyle_id_in(tree, &val_self),
            )
        };
        if let Some(v) = direct {
            return Some(v);
        }

        // Levels 2–4 live in styles.xml.
        let styles = doc.styles_part()?;
        let val = doc.qn(styles, "val");
        let styleid_attr = doc.qn(styles, "styleId");
        let tree = doc.tree(styles);
        let root = tree.root();

        if let Some(sid) = &rstyle_id {
            if let Some(v) = prop_in_style_chain(tree, root, sid, &styleid_attr, &val, extract) {
                return Some(v);
            }
        }
        if let Some(sid) = &pstyle_id {
            if let Some(v) = prop_in_style_chain(tree, root, sid, &styleid_attr, &val, extract) {
                return Some(v);
            }
        }
        docdefaults_prop(tree, root, &val, extract)
    }

    /// The run's direct character-style id (`w:rPr/w:rStyle/@w:val`), read from `tree`.
    fn rstyle_id_in(&self, tree: &XmlTree, val_name: &str) -> Option<String> {
        let rpr = self.rpr(tree)?;
        let rstyle = rpr_child_in(tree, rpr, "rStyle")?;
        tree.attr(rstyle, val_name).map(str::to_owned)
    }

    /// The style id of the paragraph containing this run (`w:pPr/w:pStyle/@w:val`), found by
    /// walking up to the nearest ancestor `w:p`. `None` for a run outside any paragraph or a
    /// paragraph with no `w:pStyle`.
    fn paragraph_pstyle_id_in(&self, tree: &XmlTree, val_name: &str) -> Option<String> {
        let mut cur = tree.parent(self.node);
        while let Some(p) = cur {
            if is_wml_element(tree, p, "p") {
                let ppr = tree
                    .children(p)
                    .iter()
                    .copied()
                    .find(|&c| is_wml_element(tree, c, "pPr"))?;
                let pstyle = rpr_child_in(tree, ppr, "pStyle")?;
                return tree.attr(pstyle, val_name).map(str::to_owned);
            }
            cur = tree.parent(p);
        }
        None
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
    /// schema order) if absent. Creates `w:rPr` first if needed. Thin wrapper over the
    /// part-agnostic [`ensure_rpr_child_in`] that ensures the run's own `w:rPr`.
    fn ensure_rpr_child(&self, doc: &mut Document, local: &str) -> NodeId {
        let rpr = self.ensure_rpr(doc);
        ensure_rpr_child_in(doc, self.part, rpr, local)
    }

    /// Remove a direct `w:rPr` child by local name, if present.
    fn remove_rpr_child(&self, doc: &mut Document, local: &str) {
        if let Some(rpr) = self.rpr(doc.tree(self.part)) {
            remove_child_in(doc, self.part, rpr, local);
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
            let rpr = self.ensure_rpr(doc);
            toggle_on_in(doc, self.part, rpr, local);
        } else {
            self.remove_rpr_child(doc, local);
        }
    }
}

// --- Reusable `w:rPr` property writers ---------------------------------------------
//
// These operate on any `(part, w:rPr)` pair rather than on a run specifically, so both
// [`Run`] and [`Style`](super::Style) share one implementation of the ordered-insert +
// attribute-write logic. The placement of the `w:rPr` *within its parent* differs (a run
// puts it first; a style slots it in `CT_Style` order), so each caller ensures its own
// `w:rPr` and passes the node id here.

/// A direct `w:rPr` child element with WML local name `local`, if present.
pub(super) fn rpr_child_in(tree: &XmlTree, rpr: NodeId, local: &str) -> Option<NodeId> {
    tree.children(rpr)
        .iter()
        .copied()
        .find(|&c| is_wml_element(tree, c, local))
}

/// Ensure a direct `w:rPr` child `local` exists, inserted in canonical `CT_RPr` order
/// ([`RPR_ORDER`]); return it. `rpr` is a `w:rPr` element living in `part`.
pub(super) fn ensure_rpr_child_in(
    doc: &mut Document,
    part: PartId,
    rpr: NodeId,
    local: &str,
) -> NodeId {
    if let Some(existing) = rpr_child_in(doc.tree(part), rpr, local) {
        return existing;
    }
    let name = doc.qn(part, local);
    let index = ordered_insert_index(doc.tree(part), rpr, rank_in(RPR_ORDER, local), RPR_ORDER);
    let el = doc.tree_mut(part).create_element(name);
    doc.tree_mut(part).insert_child(rpr, index, el);
    el
}

/// Turn a boolean toggle (`w:b`, `w:i`, …) on within `rpr`: ensure a bare element,
/// clearing any explicit `w:val="0"/"false"` so it reads on.
pub(super) fn toggle_on_in(doc: &mut Document, part: PartId, rpr: NodeId, local: &str) {
    let el = ensure_rpr_child_in(doc, part, rpr, local);
    let val = doc.qn(part, "val");
    let tree = doc.tree_mut(part);
    if let Some(v) = tree.attr(el, &val) {
        if matches!(v, "0" | "false") {
            tree.remove_attr(el, &val);
        }
    }
}

/// Remove a direct `w:rPr` child `local` from `rpr`, if present.
pub(super) fn remove_child_in(doc: &mut Document, part: PartId, rpr: NodeId, local: &str) {
    if let Some(el) = rpr_child_in(doc.tree(part), rpr, local) {
        doc.tree_mut(part).remove_from_parent(el);
    }
}

/// Set the font size on `rpr`: `w:sz` and `w:szCs` in half-points (python-docx parity).
pub(super) fn set_size_in(doc: &mut Document, part: PartId, rpr: NodeId, size: Pt) {
    let hp = size.to_half_points_string();
    let val = doc.qn(part, "val");
    let sz = ensure_rpr_child_in(doc, part, rpr, "sz");
    doc.tree_mut(part).set_attr(sz, val.clone(), hp.clone());
    let sz_cs = ensure_rpr_child_in(doc, part, rpr, "szCs");
    doc.tree_mut(part).set_attr(sz_cs, val, hp);
}

/// Set the color on `rpr` (`w:color w:val` as six uppercase hex digits).
pub(super) fn set_color_in(doc: &mut Document, part: PartId, rpr: NodeId, color: RgbColor) {
    let val = doc.qn(part, "val");
    let el = ensure_rpr_child_in(doc, part, rpr, "color");
    doc.tree_mut(part).set_attr(el, val, color.to_hex());
}

/// Set the font on `rpr` (`w:rFonts` `w:ascii` and `w:hAnsi`, the two Latin-text attrs).
pub(super) fn set_font_in(doc: &mut Document, part: PartId, rpr: NodeId, name: &str) {
    let ascii = doc.qn(part, "ascii");
    let hansi = doc.qn(part, "hAnsi");
    let el = ensure_rpr_child_in(doc, part, rpr, "rFonts");
    doc.tree_mut(part).set_attr(el, ascii, name);
    doc.tree_mut(part).set_attr(el, hansi, name);
}

// --- Effective-formatting resolution helpers ---------------------------------------

/// Read a toggle property (`w:b`, `w:i`) out of `rpr` as a three-state value: `None` when
/// the element is absent at this level (so resolution continues up the chain), else `Some`
/// of whether it is on (present and not `w:val="0"/"false"`).
fn read_toggle(tree: &XmlTree, rpr: NodeId, val_name: &str, local: &str) -> Option<bool> {
    let el = rpr_child_in(tree, rpr, local)?;
    Some(match tree.attr(el, val_name) {
        Some(v) => !matches!(v, "0" | "false"),
        None => true,
    })
}

/// Walk a `w:style` and its `w:basedOn` chain in the styles tree, returning the first
/// property `extract` yields from a style's `w:rPr`. Cycle-guarded with a seen-set (a
/// malformed `basedOn` loop terminates rather than spinning), like
/// [`Document::style_numbering`](super::Document::style_numbering).
fn prop_in_style_chain<T>(
    tree: &XmlTree,
    root: NodeId,
    start_id: &str,
    styleid_attr: &str,
    val_name: &str,
    extract: impl Fn(&XmlTree, NodeId, &str) -> Option<T> + Copy,
) -> Option<T> {
    let mut seen: HashSet<String> = HashSet::new();
    let mut current = start_id.to_string();
    while seen.insert(current.clone()) {
        let style = tree.children(root).iter().copied().find(|&c| {
            is_wml_element(tree, c, "style") && tree.attr(c, styleid_attr) == Some(current.as_str())
        })?;
        if let Some(rpr) = rpr_child_in(tree, style, "rPr") {
            if let Some(v) = extract(tree, rpr, val_name) {
                return Some(v);
            }
        }
        current = rpr_child_in(tree, style, "basedOn")
            .and_then(|b| tree.attr(b, val_name))
            .map(str::to_owned)?;
    }
    None
}

/// Extract a property from `w:docDefaults/w:rPrDefault/w:rPr`, the lowest-priority level.
fn docdefaults_prop<T>(
    tree: &XmlTree,
    root: NodeId,
    val_name: &str,
    extract: impl Fn(&XmlTree, NodeId, &str) -> Option<T>,
) -> Option<T> {
    let docdefaults = rpr_child_in(tree, root, "docDefaults")?;
    let rpr_default = rpr_child_in(tree, docdefaults, "rPrDefault")?;
    let rpr = rpr_child_in(tree, rpr_default, "rPr")?;
    extract(tree, rpr, val_name)
}

/// Whether a run child carries text content that [`Run::set_text`] should replace.
fn is_content_child(tree: &XmlTree, id: NodeId) -> bool {
    is_wml_element(tree, id, "t")
        || is_wml_element(tree, id, "tab")
        || is_wml_element(tree, id, "br")
        || is_wml_element(tree, id, "cr")
}
