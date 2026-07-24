//! The [`Paragraph`] handle.

use crate::xml::{NodeId, XmlTree};

use super::{
    Alignment, Document, PartId, Run, is_wml_element, needs_space_preserve, ordered_insert_index,
    rank_in,
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
