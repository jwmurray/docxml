//! The [`Paragraph`] handle.

use crate::xml::{NodeId, XmlTree};

use super::{
    Alignment, Document, Length, LineSpacing, PartId, Pt, Run, TabAlignment, TabLeader,
    is_wml_element, needs_space_preserve, ordered_insert_index, rank_in,
};

/// Canonical `w:pPr` child order (ECMA-376 §17.3.1.26, `CT_PPr` sequence), local names
/// only. New properties are inserted to keep `w:pPr`'s children in this order so the
/// output is schema-valid: `pStyle` comes first, `jc` sits late (after the spacing /
/// indentation group), and `rPr` / `sectPr` come near the end. Unlisted children rank
/// last and stay after authored properties.
const PPR_ORDER: &[&str] = &[
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
];

/// A lightweight handle to a `w:p` paragraph.
///
/// `Paragraph` is `Copy` and borrows nothing — it is just an arena node id (plus the id of
/// the part it lives in) with phantom typing. Pass a [`Document`] back to it (`&Document`
/// to read, `&mut Document` to edit) to do anything useful. A paragraph read from a header
/// or footer carries that part's id, so its text and runs resolve against the header/footer
/// tree, not the body.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Paragraph {
    part: PartId,
    node: NodeId,
}

impl Paragraph {
    /// Wrap a known-`w:p` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Paragraph { part, node }
    }

    /// The paragraph's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The part this paragraph lives in (used by field-code construction).
    pub(crate) fn part(&self) -> PartId {
        self.part
    }

    /// The paragraph's visible text: its runs' text concatenated.
    ///
    /// Within each run, `w:t` contributes its text, `w:tab` a tab, and `w:br` / `w:cr` a
    /// newline — matching python-docx's `Paragraph.text`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("");
    /// p.add_run(&mut doc, "Hello, ");
    /// p.add_run(&mut doc, "world");
    /// assert_eq!(p.text(&doc), "Hello, world");
    /// ```
    pub fn text(&self, doc: &Document) -> String {
        let tree = doc.tree(self.part);
        let mut out = String::new();
        for run in self.run_nodes(tree) {
            append_run_text(tree, run, &mut out);
        }
        out
    }

    /// The paragraph's runs, in order.
    pub fn runs(&self, doc: &Document) -> Vec<Run> {
        self.run_nodes(doc.tree(self.part))
            .map(|r| Run::from_node(self.part, r))
            .collect()
    }

    /// Append a run carrying `text`, returning a handle to it.
    ///
    /// Builds `<w:r><w:t>text</w:t></w:r>`, setting `xml:space="preserve"` on the `w:t`
    /// when `text` has leading or trailing whitespace.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("");
    /// let r = p.add_run(&mut doc, "bold text");
    /// r.bold(&mut doc, true);
    /// assert!(r.is_bold(&doc));
    /// ```
    pub fn add_run(&self, doc: &mut Document, text: &str) -> Run {
        let r_name = doc.qn(self.part, "r");
        let t_name = doc.qn(self.part, "t");

        let tree = doc.tree_mut(self.part);
        let r = tree.create_element(r_name);
        let t = tree.create_element(t_name);
        let content = tree.create_text(text);
        tree.append_child(t, content);
        tree.append_child(r, t);
        tree.append_child(self.node, r);
        if needs_space_preserve(text) {
            tree.set_attr(t, "xml:space", "preserve");
        }
        Run::from_node(self.part, r)
    }

    /// The paragraph's alignment, if `w:pPr/w:jc` is set to a recognized value
    /// (`w:val="distribute"` and other unrecognized values read as `None`).
    pub fn alignment(&self, doc: &Document) -> Option<Alignment> {
        let tree = doc.tree(self.part);
        let jc = self.ppr_child(tree, "jc")?;
        Alignment::from_val(tree.attr(jc, &doc.qn(self.part, "val"))?)
    }

    /// Set the paragraph alignment (`w:pPr/w:jc`).
    pub fn set_alignment(&self, doc: &mut Document, alignment: Alignment) -> Paragraph {
        let val = doc.qn(self.part, "val");
        let jc = self.ensure_ppr_child(doc, "jc");
        doc.tree_mut(self.part)
            .set_attr(jc, val, alignment.to_val());
        *self
    }

    /// The paragraph's style id, if `w:pPr/w:pStyle` is set.
    ///
    /// This is the *styleId* — the internal key such as `"Heading1"` — not the human
    /// display name (`"heading 1"`). Resolving a styleId to its display name requires
    /// reading `styles.xml`, which is a later milestone.
    pub fn style_id(&self, doc: &Document) -> Option<String> {
        let tree = doc.tree(self.part);
        let pstyle = self.ppr_child(tree, "pStyle")?;
        tree.attr(pstyle, &doc.qn(self.part, "val"))
            .map(str::to_owned)
    }

    /// Set the paragraph's style id (`w:pPr/w:pStyle w:val`).
    ///
    /// `style_id` is the internal styleId (e.g. `"Heading1"`), not the display name; see
    /// [`style_id`](Self::style_id).
    pub fn set_style_id(&self, doc: &mut Document, style_id: &str) -> Paragraph {
        let val = doc.qn(self.part, "val");
        let pstyle = self.ensure_ppr_child(doc, "pStyle");
        doc.tree_mut(self.part).set_attr(pstyle, val, style_id);
        *self
    }

    /// The paragraph's line spacing (`w:pPr/w:spacing`), if set.
    ///
    /// Reads `w:line` together with `w:lineRule` and maps them to a [`LineSpacing`]
    /// (see that type for the auto-multiple vs. exact/at-least rules). A missing
    /// `w:lineRule` alongside a `w:line` value is read as an auto multiple.
    pub fn line_spacing(&self, doc: &Document) -> Option<LineSpacing> {
        let tree = doc.tree(self.part);
        let sp = self.ppr_child(tree, "spacing")?;
        let line = tree.attr(sp, &doc.qn(self.part, "line"))?;
        let rule = tree.attr(sp, &doc.qn(self.part, "lineRule"));
        LineSpacing::from_line_and_rule(line, rule)
    }

    /// Set the paragraph line spacing (`w:pPr/w:spacing` `w:line` + `w:lineRule`).
    ///
    /// `w:spacing` is a single element shared with [`space_before`](Self::space_before) and
    /// [`space_after`](Self::space_after); this touches only `w:line` and `w:lineRule`,
    /// leaving any `w:before` / `w:after` intact.
    pub fn set_line_spacing(&self, doc: &mut Document, spacing: LineSpacing) -> Paragraph {
        let (line, rule) = spacing.to_line_and_rule();
        let line_attr = doc.qn(self.part, "line");
        let rule_attr = doc.qn(self.part, "lineRule");
        let sp = self.ensure_ppr_child(doc, "spacing");
        let tree = doc.tree_mut(self.part);
        tree.set_attr(sp, line_attr, line);
        tree.set_attr(sp, rule_attr, rule);
        *self
    }

    /// The space above the paragraph (`w:pPr/w:spacing/@w:before`), if set.
    pub fn space_before(&self, doc: &Document) -> Option<Pt> {
        self.read_spacing_pt(doc, "before")
    }

    /// The space below the paragraph (`w:pPr/w:spacing/@w:after`), if set.
    pub fn space_after(&self, doc: &Document) -> Option<Pt> {
        self.read_spacing_pt(doc, "after")
    }

    /// Set the space above the paragraph (`w:pPr/w:spacing/@w:before`, in twentieths of a
    /// point). Preserves the other `w:spacing` attributes (`line`, `lineRule`, `after`).
    pub fn set_space_before(&self, doc: &mut Document, space: Pt) -> Paragraph {
        self.set_spacing_pt(doc, "before", space);
        *self
    }

    /// Set the space below the paragraph (`w:pPr/w:spacing/@w:after`, in twentieths of a
    /// point). Preserves the other `w:spacing` attributes (`line`, `lineRule`, `before`).
    pub fn set_space_after(&self, doc: &mut Document, space: Pt) -> Paragraph {
        self.set_spacing_pt(doc, "after", space);
        *self
    }

    /// The paragraph's left indent (`w:pPr/w:ind`), if set. Reads the transitional
    /// `w:left` and the strict `w:start` spelling.
    pub fn left_indent(&self, doc: &Document) -> Option<Length> {
        self.read_ind(doc, &["left", "start"])
    }

    /// The paragraph's right indent (`w:pPr/w:ind`), if set. Reads the transitional
    /// `w:right` and the strict `w:end` spelling.
    pub fn right_indent(&self, doc: &Document) -> Option<Length> {
        self.read_ind(doc, &["right", "end"])
    }

    /// The paragraph's first-line indent (`w:pPr/w:ind`), if set.
    ///
    /// A `w:firstLine` reads as its positive value; a `w:hanging` reads as the negative of
    /// its (positive) value — python-docx's convention, where a hanging indent is a
    /// negative first-line indent. `w:hanging` takes precedence when both are present.
    pub fn first_line_indent(&self, doc: &Document) -> Option<Length> {
        let tree = doc.tree(self.part);
        let ind = self.ppr_child(tree, "ind")?;
        if let Some(hanging) = tree
            .attr(ind, &doc.qn(self.part, "hanging"))
            .and_then(Length::from_twips_str)
        {
            return Some(Length::from_twips(-hanging.twips()));
        }
        tree.attr(ind, &doc.qn(self.part, "firstLine"))
            .and_then(Length::from_twips_str)
    }

    /// Set the left indent (`w:pPr/w:ind/@w:left`), creating `w:ind` if needed.
    pub fn set_left_indent(&self, doc: &mut Document, indent: Length) -> Paragraph {
        self.set_ind(doc, "left", indent);
        *self
    }

    /// Set the right indent (`w:pPr/w:ind/@w:right`), creating `w:ind` if needed.
    pub fn set_right_indent(&self, doc: &mut Document, indent: Length) -> Paragraph {
        self.set_ind(doc, "right", indent);
        *self
    }

    /// Set the first-line indent (`w:pPr/w:ind`), creating `w:ind` if needed.
    ///
    /// A non-negative `indent` writes `w:firstLine`; a negative one writes `w:hanging` with
    /// the magnitude (python-docx's convention). Whichever of the two is not used is
    /// removed, so a paragraph never carries both spellings at once.
    pub fn set_first_line_indent(&self, doc: &mut Document, indent: Length) -> Paragraph {
        let first_attr = doc.qn(self.part, "firstLine");
        let hanging_attr = doc.qn(self.part, "hanging");
        let ind = self.ensure_ppr_child(doc, "ind");
        let twips = indent.twips();
        let tree = doc.tree_mut(self.part);
        if twips < 0 {
            tree.remove_attr(ind, &first_attr);
            tree.set_attr(ind, hanging_attr, (-twips).to_string());
        } else {
            tree.remove_attr(ind, &hanging_attr);
            tree.set_attr(ind, first_attr, twips.to_string());
        }
        *self
    }

    /// Whether the paragraph's lines are kept together on one page (`w:pPr/w:keepLines`).
    pub fn keep_together(&self, doc: &Document) -> bool {
        self.ppr_toggle(doc, "keepLines")
    }

    /// Set whether the paragraph's lines are kept together (`w:pPr/w:keepLines`).
    pub fn set_keep_together(&self, doc: &mut Document, on: bool) -> Paragraph {
        self.set_ppr_toggle(doc, "keepLines", on);
        *self
    }

    /// Whether the paragraph is kept on the same page as the next (`w:pPr/w:keepNext`).
    pub fn keep_with_next(&self, doc: &Document) -> bool {
        self.ppr_toggle(doc, "keepNext")
    }

    /// Set whether the paragraph is kept with the next (`w:pPr/w:keepNext`).
    pub fn set_keep_with_next(&self, doc: &mut Document, on: bool) -> Paragraph {
        self.set_ppr_toggle(doc, "keepNext", on);
        *self
    }

    /// Whether a page break is forced before the paragraph (`w:pPr/w:pageBreakBefore`).
    pub fn page_break_before(&self, doc: &Document) -> bool {
        self.ppr_toggle(doc, "pageBreakBefore")
    }

    /// Set whether a page break is forced before the paragraph (`w:pPr/w:pageBreakBefore`).
    pub fn set_page_break_before(&self, doc: &mut Document, on: bool) -> Paragraph {
        self.set_ppr_toggle(doc, "pageBreakBefore", on);
        *self
    }

    /// The paragraph's tab stops (`w:pPr/w:tabs`), in document order.
    ///
    /// Each entry is the stop's position, alignment, and leader. A `w:tab` with an
    /// unrecognized `w:val` defaults to [`TabAlignment::Left`]; a missing or `"none"`
    /// leader is [`TabLeader::None`]; a `w:tab` without a parsable `w:pos` (e.g. a
    /// `w:val="clear"` stop) is skipped.
    pub fn tab_stops(&self, doc: &Document) -> Vec<(Length, TabAlignment, TabLeader)> {
        let tree = doc.tree(self.part);
        let Some(tabs) = self.ppr_child(tree, "tabs") else {
            return Vec::new();
        };
        let pos_attr = doc.qn(self.part, "pos");
        let val_attr = doc.qn(self.part, "val");
        let leader_attr = doc.qn(self.part, "leader");
        tree.children(tabs)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "tab"))
            .filter_map(|c| {
                let pos = tree.attr(c, &pos_attr).and_then(Length::from_twips_str)?;
                let align = tree
                    .attr(c, &val_attr)
                    .and_then(TabAlignment::from_val)
                    .unwrap_or(TabAlignment::Left);
                let leader = tree
                    .attr(c, &leader_attr)
                    .map(TabLeader::from_val)
                    .unwrap_or(TabLeader::None);
                Some((pos, align, leader))
            })
            .collect()
    }

    /// Add a tab stop (`w:pPr/w:tabs/w:tab`), creating `w:tabs` if needed.
    ///
    /// The new `w:tab` carries `w:pos` (in twips), `w:val` (alignment), and `w:leader`
    /// (omitted for [`TabLeader::None`]). Stops are kept sorted ascending by position — the
    /// new one is inserted before the first existing stop that sits farther right — matching
    /// python-docx's position-ordered `TabStops`.
    pub fn add_tab_stop(
        &self,
        doc: &mut Document,
        pos: Length,
        alignment: TabAlignment,
        leader: TabLeader,
    ) {
        let val_attr = doc.qn(self.part, "val");
        let pos_attr = doc.qn(self.part, "pos");
        let leader_attr = doc.qn(self.part, "leader");
        let tab_name = doc.qn(self.part, "tab");
        let tabs = self.ensure_ppr_child(doc, "tabs");

        // Insert before the first existing stop positioned farther right (ascending order).
        let pos_twips = pos.twips();
        let index = {
            let tree = doc.tree(self.part);
            tree.children(tabs)
                .iter()
                .position(|&c| {
                    is_wml_element(tree, c, "tab")
                        && tree
                            .attr(c, &pos_attr)
                            .and_then(Length::from_twips_str)
                            .is_some_and(|l| l.twips() > pos_twips)
                })
                .unwrap_or_else(|| tree.children(tabs).len())
        };

        let tab = doc.tree_mut(self.part).create_element(tab_name);
        // Schema order of CT_TabStop attributes: val, leader, pos.
        doc.tree_mut(self.part)
            .set_attr(tab, val_attr, alignment.to_val());
        if let Some(l) = leader.to_val() {
            doc.tree_mut(self.part).set_attr(tab, leader_attr, l);
        }
        doc.tree_mut(self.part)
            .set_attr(tab, pos_attr, pos.to_twips_string());
        doc.tree_mut(self.part).insert_child(tabs, index, tab);
    }

    /// Remove all tab stops, deleting the `w:pPr/w:tabs` element.
    ///
    /// The whole `w:tabs` is removed rather than left empty: `CT_Tabs` requires at least one
    /// `w:tab`, so an empty `w:tabs` would be schema-invalid. After this,
    /// [`tab_stops`](Self::tab_stops) returns an empty vector.
    pub fn clear_tab_stops(&self, doc: &mut Document) {
        if let Some(tabs) = self.ppr_child(doc.tree(self.part), "tabs") {
            doc.tree_mut(self.part).remove_from_parent(tabs);
        }
    }

    /// Read a twentieths-of-a-point `w:spacing` attribute as [`Pt`].
    fn read_spacing_pt(&self, doc: &Document, attr_local: &str) -> Option<Pt> {
        let tree = doc.tree(self.part);
        let sp = self.ppr_child(tree, "spacing")?;
        Pt::from_twentieths_str(tree.attr(sp, &doc.qn(self.part, attr_local))?)
    }

    /// Set a twentieths-of-a-point `w:spacing` attribute, creating `w:spacing` if needed.
    fn set_spacing_pt(&self, doc: &mut Document, attr_local: &str, space: Pt) {
        let attr = doc.qn(self.part, attr_local);
        let sp = self.ensure_ppr_child(doc, "spacing");
        doc.tree_mut(self.part)
            .set_attr(sp, attr, space.to_twentieths_string());
    }

    /// Read the first present of `attr_locals` on `w:ind` as a [`Length`].
    fn read_ind(&self, doc: &Document, attr_locals: &[&str]) -> Option<Length> {
        let tree = doc.tree(self.part);
        let ind = self.ppr_child(tree, "ind")?;
        for local in attr_locals {
            if let Some(v) = tree.attr(ind, &doc.qn(self.part, local)) {
                return Length::from_twips_str(v);
            }
        }
        None
    }

    /// Set a twips-valued `w:ind` attribute, creating `w:ind` if needed.
    fn set_ind(&self, doc: &mut Document, attr_local: &str, indent: Length) {
        let attr = doc.qn(self.part, attr_local);
        let ind = self.ensure_ppr_child(doc, "ind");
        doc.tree_mut(self.part)
            .set_attr(ind, attr, indent.to_twips_string());
    }

    /// Read an on/off `w:pPr` toggle child (`w:keepLines`, `w:keepNext`, …): present and not
    /// `w:val="0"/"false"` is on.
    fn ppr_toggle(&self, doc: &Document, local: &str) -> bool {
        let tree = doc.tree(self.part);
        let Some(el) = self.ppr_child(tree, local) else {
            return false;
        };
        match tree.attr(el, &doc.qn(self.part, "val")) {
            Some(v) => !matches!(v, "0" | "false"),
            None => true,
        }
    }

    /// Set or clear an on/off `w:pPr` toggle child. On ensures a bare element (clearing an
    /// explicit `w:val="0"/"false"`); off removes it.
    fn set_ppr_toggle(&self, doc: &mut Document, local: &str, on: bool) {
        if on {
            let el = self.ensure_ppr_child(doc, local);
            let val = doc.qn(self.part, "val");
            let tree = doc.tree_mut(self.part);
            if let Some(v) = tree.attr(el, &val) {
                if matches!(v, "0" | "false") {
                    tree.remove_attr(el, &val);
                }
            }
        } else if let Some(el) = self.ppr_child(doc.tree(self.part), local) {
            doc.tree_mut(self.part).remove_from_parent(el);
        }
    }

    /// The paragraph's direct `w:r` children as node ids.
    fn run_nodes<'a>(&self, tree: &'a XmlTree) -> impl Iterator<Item = NodeId> + 'a {
        tree.children(self.node)
            .iter()
            .copied()
            .filter(move |&c| is_wml_element(tree, c, "r"))
    }

    /// The paragraph's `w:pPr`, if present.
    fn ppr(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "pPr"))
    }

    /// The paragraph's `w:pPr`, creating it as the first child if absent (schema
    /// requires `w:pPr` before the paragraph content).
    fn ensure_ppr(&self, doc: &mut Document) -> NodeId {
        if let Some(ppr) = self.ppr(doc.tree(self.part)) {
            return ppr;
        }
        let name = doc.qn(self.part, "pPr");
        let tree = doc.tree_mut(self.part);
        let ppr = tree.create_element(name);
        tree.insert_child(self.node, 0, ppr);
        ppr
    }

    /// A direct `w:pPr` child with the given WML local name, if present.
    fn ppr_child(&self, tree: &XmlTree, local: &str) -> Option<NodeId> {
        let ppr = self.ppr(tree)?;
        tree.children(ppr)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, local))
    }

    /// A direct `w:pPr` child with the given local name, creating it (in canonical
    /// schema order) if absent. Creates `w:pPr` first if needed.
    fn ensure_ppr_child(&self, doc: &mut Document, local: &str) -> NodeId {
        let ppr = self.ensure_ppr(doc);
        if let Some(existing) = {
            let tree = doc.tree(self.part);
            tree.children(ppr)
                .iter()
                .copied()
                .find(|&c| is_wml_element(tree, c, local))
        } {
            return existing;
        }
        let name = doc.qn(self.part, local);
        let index = ordered_insert_index(
            doc.tree(self.part),
            ppr,
            rank_in(PPR_ORDER, local),
            PPR_ORDER,
        );
        let el = doc.tree_mut(self.part).create_element(name);
        doc.tree_mut(self.part).insert_child(ppr, index, el);
        el
    }
}

/// Append a run's text (python-docx semantics) to `out`: `w:t` text verbatim, `w:tab`
/// as a tab, `w:br` / `w:cr` as a newline.
pub(super) fn append_run_text(tree: &XmlTree, run: NodeId, out: &mut String) {
    for descendant in tree.descendants(run) {
        if is_wml_element(tree, descendant, "t") {
            out.push_str(&tree.text_content(descendant));
        } else if is_wml_element(tree, descendant, "tab") {
            out.push('\t');
        } else if is_wml_element(tree, descendant, "br") || is_wml_element(tree, descendant, "cr") {
            out.push('\n');
        }
    }
}
