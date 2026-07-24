//! The [`Paragraph`] handle.

use crate::xml::{NodeId, XmlTree};

use super::{Document, Run, is_wml_element, needs_space_preserve};

/// A lightweight handle to a `w:p` paragraph.
///
/// `Paragraph` is `Copy` and borrows nothing — it is just an arena node id with phantom
/// typing. Pass a [`Document`] back to it (`&Document` to read, `&mut Document` to edit)
/// to do anything useful.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Paragraph {
    node: NodeId,
}

impl Paragraph {
    /// Wrap a known-`w:p` node id.
    pub(crate) fn from_node(node: NodeId) -> Self {
        Paragraph { node }
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
        let tree = doc.tree();
        let mut out = String::new();
        for run in self.run_nodes(tree) {
            append_run_text(tree, run, &mut out);
        }
        out
    }

    /// The paragraph's runs, in order.
    pub fn runs(&self, doc: &Document) -> Vec<Run> {
        self.run_nodes(doc.tree()).map(Run::from_node).collect()
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
        let r_name = doc.qn("r");
        let t_name = doc.qn("t");

        let tree = doc.tree_mut();
        let r = tree.create_element(r_name);
        let t = tree.create_element(t_name);
        let content = tree.create_text(text);
        tree.append_child(t, content);
        tree.append_child(r, t);
        tree.append_child(self.node, r);
        if needs_space_preserve(text) {
            tree.set_attr(t, "xml:space", "preserve");
        }
        Run::from_node(r)
    }

    /// The paragraph's direct `w:r` children as node ids.
    fn run_nodes<'a>(&self, tree: &'a XmlTree) -> impl Iterator<Item = NodeId> + 'a {
        tree.children(self.node)
            .iter()
            .copied()
            .filter(move |&c| is_wml_element(tree, c, "r"))
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
