//! The [`Run`] handle and character-formatting accessors.

use crate::xml::{NodeId, XmlTree};

use super::paragraph::append_run_text;
use super::{Document, is_wml_element, needs_space_preserve, split_qname};

/// A lightweight handle to a `w:r` run — a contiguous span of text with uniform
/// character formatting.
///
/// Like [`Paragraph`](super::Paragraph), `Run` is `Copy` and borrows nothing. The
/// formatting setters return the run so calls chain:
/// `run.bold(&mut doc, true).italic(&mut doc, true)`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Run {
    node: NodeId,
}

impl Run {
    /// Wrap a known-`w:r` node id.
    pub(crate) fn from_node(node: NodeId) -> Self {
        Run { node }
    }

    /// The run's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The run's text: `w:t` verbatim, `w:tab` as a tab, `w:br` / `w:cr` as a newline.
    pub fn text(&self, doc: &Document) -> String {
        let mut out = String::new();
        append_run_text(doc.tree(), self.node, &mut out);
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
        let t_name = doc.qn("t");
        let tree = doc.tree_mut();

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
    /// True when `w:rPr` carries a `w:b` whose `w:val` is not `"0"` / `"false"` (a bare
    /// `w:b`, as Word writes, means on).
    pub fn is_bold(&self, doc: &Document) -> bool {
        self.has_toggle(doc, "b")
    }

    /// Whether the run is italic (see [`is_bold`](Self::is_bold) for the `w:val` rule).
    pub fn is_italic(&self, doc: &Document) -> bool {
        self.has_toggle(doc, "i")
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
        if let Some(rpr) = self.rpr(doc.tree()) {
            return rpr;
        }
        let name = doc.qn("rPr");
        let tree = doc.tree_mut();
        let rpr = tree.create_element(name);
        tree.insert_child(self.node, 0, rpr);
        rpr
    }

    /// Read a boolean toggle property (`w:b`, `w:i`) from `w:rPr`.
    fn has_toggle(&self, doc: &Document, local: &str) -> bool {
        let tree = doc.tree();
        let Some(rpr) = self.rpr(tree) else {
            return false;
        };
        let Some(el) = tree
            .children(rpr)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, local))
        else {
            return false;
        };
        match tree.attr(el, &doc.qn("val")) {
            Some(v) => !matches!(v, "0" | "false"),
            None => true,
        }
    }

    /// Set or clear a boolean toggle property (`w:b`, `w:i`) in `w:rPr`.
    fn set_toggle(&self, doc: &mut Document, local: &str, on: bool) {
        if on {
            let rpr = self.ensure_rpr(doc);
            let existing = {
                let tree = doc.tree();
                tree.children(rpr)
                    .iter()
                    .copied()
                    .find(|&c| is_wml_element(tree, c, local))
            };
            match existing {
                // Already present: clear an explicit `w:val="0"/"false"` so it reads on.
                Some(el) => {
                    let val_name = doc.qn("val");
                    let tree = doc.tree_mut();
                    if let Some(v) = tree.attr(el, &val_name) {
                        if matches!(v, "0" | "false") {
                            tree.remove_attr(el, &val_name);
                        }
                    }
                }
                None => {
                    let name = doc.qn(local);
                    let index = insertion_index(doc.tree(), rpr, rpr_rank(local));
                    let el = doc.tree_mut().create_element(name);
                    doc.tree_mut().insert_child(rpr, index, el);
                }
            }
        } else {
            let existing = {
                let tree = doc.tree();
                self.rpr(tree).and_then(|rpr| {
                    tree.children(rpr)
                        .iter()
                        .copied()
                        .find(|&c| is_wml_element(tree, c, local))
                })
            };
            if let Some(el) = existing {
                doc.tree_mut().remove_from_parent(el);
            }
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

/// Schema-order rank of the `w:rPr` children this module authors (lower comes first).
/// Unlisted children rank last so authored properties slot in ahead of them.
fn rpr_rank(local: &str) -> u32 {
    match local {
        "rStyle" => 0,
        "b" => 1,
        "bCs" => 2,
        "i" => 3,
        "iCs" => 4,
        _ => u32::MAX,
    }
}

/// The rank of an existing `w:rPr` child (unknown / non-WML children rank last).
fn child_rank(tree: &XmlTree, id: NodeId) -> u32 {
    match tree.name(id) {
        Some(name) => rpr_rank(split_qname(name).1),
        None => u32::MAX,
    }
}

/// Index at which to insert a new `w:rPr` child of the given rank so ranks stay
/// ascending: before the first existing child that ranks after it.
fn insertion_index(tree: &XmlTree, rpr: NodeId, rank: u32) -> usize {
    let children = tree.children(rpr);
    children
        .iter()
        .position(|&c| child_rank(tree, c) > rank)
        .unwrap_or(children.len())
}
