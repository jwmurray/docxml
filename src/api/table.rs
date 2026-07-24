//! The [`Table`], [`Row`], and [`Cell`] handles.
//!
//! These are `Copy` arena-node handles in the same spirit as [`Paragraph`](super::Paragraph)
//! and [`Run`](super::Run): a [`NodeId`] plus phantom typing, borrowing nothing. Reads go
//! through `&Document`, mutations through `&mut Document`.
//!
//! # Physical, not grid, cells
//!
//! [`Row::cells`] returns the row's *physical* `w:tc` elements: a horizontally merged
//! cell (`w:gridSpan`) appears once, and a vertically merged continuation (`w:vMerge`
//! without `w:val="restart"`) is a distinct `w:tc` in its own row. This deliberately
//! differs from python-docx, whose `row.cells` is grid-based: it expands spans so every
//! grid position yields a cell (the same underlying `<w:tc>` repeated for a span). Full
//! grid/virtual-cell semantics need a merge-resolution pass over the whole table and are
//! deferred; [`Cell::grid_span`] and [`Cell::v_merge`] expose the raw merge markers so a
//! caller can reason about merges today.

use crate::xml::{NodeId, XmlTree};

use super::{Document, Paragraph, PartId, is_wml_element};

/// A lightweight handle to a `w:tbl` table.
///
/// `Table` is `Copy` and borrows nothing — just an arena node id (plus its part id) with
/// phantom typing. Pass a [`Document`] back to it (`&Document` to read, `&mut Document` to
/// edit).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Table {
    part: PartId,
    node: NodeId,
}

/// A lightweight handle to a `w:tr` table row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Row {
    part: PartId,
    node: NodeId,
}

/// A lightweight handle to a `w:tc` table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Cell {
    part: PartId,
    node: NodeId,
}

/// A cell's vertical-merge role (`w:tcPr/w:vMerge`).
///
/// `w:vMerge w:val="restart"` starts a vertically merged region ([`Restart`](VMerge::Restart));
/// a `w:vMerge` with no `w:val`, or `w:val="continue"`, continues the region above it
/// ([`Continue`](VMerge::Continue)). A cell with no `w:vMerge` is unmerged and reads as
/// `None` from [`Cell::v_merge`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VMerge {
    /// The top cell of a vertical merge (`w:val="restart"`).
    Restart,
    /// A continuation cell merged into the one above (`w:val="continue"` or bare).
    Continue,
}

impl Table {
    /// Wrap a known-`w:tbl` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Table { part, node }
    }

    /// The table's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The table's rows, in order (its direct `w:tr` children).
    pub fn rows(&self, doc: &Document) -> Vec<Row> {
        let tree = doc.tree(self.part);
        tree.children(self.node)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "tr"))
            .map(|c| Row::from_node(self.part, c))
            .collect()
    }

    /// The number of columns declared by the table grid (its `w:tblGrid/w:gridCol` count).
    fn grid_col_count(&self, tree: &XmlTree) -> usize {
        let Some(grid) = tree
            .children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "tblGrid"))
        else {
            return 0;
        };
        tree.children(grid)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "gridCol"))
            .count()
    }

    /// Append a new row to the table, returning a handle to it.
    ///
    /// The row is given one empty cell per grid column (the column count comes from the
    /// table's `w:tblGrid`), matching python-docx's `Table.add_row`. Each cell is a
    /// `w:tc` with a minimal `w:tcPr` (auto width) and one empty `w:p`.
    pub fn add_row(&self, doc: &mut Document) -> Row {
        let cols = self.grid_col_count(doc.tree(self.part));
        let tr_name = doc.qn(self.part, "tr");
        let tr = doc.tree_mut(self.part).create_element(tr_name);
        for _ in 0..cols {
            let tc = build_cell(doc, self.part);
            doc.tree_mut(self.part).append_child(tr, tc);
        }
        doc.tree_mut(self.part).append_child(self.node, tr);
        Row::from_node(self.part, tr)
    }
}

impl Row {
    /// Wrap a known-`w:tr` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Row { part, node }
    }

    /// The row's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The row's cells, in order — its *physical* `w:tc` children.
    ///
    /// A horizontally merged cell (`w:gridSpan`) appears once here, not once per spanned
    /// grid column; see the [module docs](self) for how this differs from python-docx's
    /// grid-based cells.
    pub fn cells(&self, doc: &Document) -> Vec<Cell> {
        let tree = doc.tree(self.part);
        tree.children(self.node)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "tc"))
            .map(|c| Cell::from_node(self.part, c))
            .collect()
    }
}

impl Cell {
    /// Wrap a known-`w:tc` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Cell { part, node }
    }

    /// The cell's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The cell's paragraphs, in order — its *direct* `w:p` children.
    ///
    /// Paragraphs inside a nested table are not included (they belong to that table's
    /// cells), matching python-docx's `_Cell.paragraphs`.
    pub fn paragraphs(&self, doc: &Document) -> Vec<Paragraph> {
        let tree = doc.tree(self.part);
        tree.children(self.node)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "p"))
            .map(|c| Paragraph::from_node(self.part, c))
            .collect()
    }

    /// The cell's text: its direct paragraphs' text joined with `'\n'`.
    ///
    /// This matches python-docx's `_Cell.text` getter (one newline between paragraphs).
    pub fn text(&self, doc: &Document) -> String {
        self.paragraphs(doc)
            .iter()
            .map(|p| p.text(doc))
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Append a new paragraph to the cell carrying `text`, returning a handle to it.
    ///
    /// The paragraph is added as the cell's last child; when `text` is non-empty a single
    /// run carrying it is added. Reuses the [`Paragraph`] handle, so all of its
    /// formatting methods apply.
    pub fn add_paragraph(&self, doc: &mut Document, text: &str) -> Paragraph {
        let name = doc.qn(self.part, "p");
        let p = doc.tree_mut(self.part).create_element(name);
        doc.tree_mut(self.part).append_child(self.node, p);
        let para = Paragraph::from_node(self.part, p);
        if !text.is_empty() {
            para.add_run(doc, text);
        }
        para
    }

    /// Replace the cell's content with a single paragraph containing `text`.
    ///
    /// This is python-docx's `_Cell.text` *setter*: every existing block item (all
    /// paragraphs and any nested table) is removed and one `w:p` with a single run
    /// carrying `text` is put in their place. The cell's `w:tcPr` is preserved. A cell
    /// must contain at least one `w:p` (the schema requires block content), so this always
    /// leaves exactly one paragraph.
    pub fn set_text(&self, doc: &mut Document, text: &str) {
        // Remove every child except the properties element.
        let to_remove: Vec<NodeId> = {
            let tree = doc.tree(self.part);
            tree.children(self.node)
                .iter()
                .copied()
                .filter(|&c| !is_wml_element(tree, c, "tcPr"))
                .collect()
        };
        for child in to_remove {
            doc.tree_mut(self.part).remove_from_parent(child);
        }
        self.add_paragraph(doc, text);
    }

    /// The cell's horizontal span in grid columns (`w:tcPr/w:gridSpan`), defaulting to `1`.
    ///
    /// A value of `2` or more means this physical cell was horizontally merged across that
    /// many grid columns. An absent or unparsable `w:gridSpan` reads as `1`.
    pub fn grid_span(&self, doc: &Document) -> u32 {
        let tree = doc.tree(self.part);
        let Some(gs) = self.tc_pr_child(tree, "gridSpan") else {
            return 1;
        };
        match tree.attr(gs, &doc.qn(self.part, "val")) {
            Some(v) => v.trim().parse().unwrap_or(1),
            None => 1,
        }
    }

    /// The cell's vertical-merge role (`w:tcPr/w:vMerge`), or `None` when unmerged.
    ///
    /// `w:val="restart"` → [`VMerge::Restart`]; a bare `w:vMerge` or `w:val="continue"` →
    /// [`VMerge::Continue`].
    pub fn v_merge(&self, doc: &Document) -> Option<VMerge> {
        let tree = doc.tree(self.part);
        let vmerge = self.tc_pr_child(tree, "vMerge")?;
        match tree.attr(vmerge, &doc.qn(self.part, "val")) {
            Some("restart") => Some(VMerge::Restart),
            _ => Some(VMerge::Continue),
        }
    }

    /// The cell's `w:tcPr`, if present.
    fn tc_pr(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "tcPr"))
    }

    /// A direct `w:tcPr` child with the given WML local name, if present.
    fn tc_pr_child(&self, tree: &XmlTree, local: &str) -> Option<NodeId> {
        let tcpr = self.tc_pr(tree)?;
        tree.children(tcpr)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, local))
    }
}

/// Build a detached empty `w:tc` in `part`: a `w:tcPr` (with `w:tcW w:w="0" w:type="auto"`)
/// followed by one empty `w:p`, in schema order (`w:tcPr` first).
fn build_cell(doc: &mut Document, part: PartId) -> NodeId {
    let tc_name = doc.qn(part, "tc");
    let tcpr_name = doc.qn(part, "tcPr");
    let tcw_name = doc.qn(part, "tcW");
    let p_name = doc.qn(part, "p");
    let w_attr = doc.qn(part, "w");
    let type_attr = doc.qn(part, "type");

    let tree = doc.tree_mut(part);
    let tc = tree.create_element(tc_name);
    let tcpr = tree.create_element(tcpr_name);
    let tcw = tree.create_element(tcw_name);
    tree.set_attr(tcw, w_attr, "0");
    tree.set_attr(tcw, type_attr, "auto");
    tree.append_child(tcpr, tcw);
    tree.append_child(tc, tcpr);
    let p = tree.create_element(p_name);
    tree.append_child(tc, p);
    tc
}

/// Build a detached `w:tbl` skeleton of `rows` × `cols` empty cells, returning its node id.
///
/// Emits `w:tblPr` (with `w:tblW w:type="auto" w:w="0"` and a `w:tblLook`) first, then a
/// `w:tblGrid` of `cols` bare `w:gridCol`, then `rows` `w:tr`, each with `cols` cells from
/// [`build_cell`]. This mirrors the skeleton python-docx's `Document.add_table` writes,
/// minus the table style (python-docx's `add_table` applies no style when none is passed).
pub(super) fn build_table(doc: &mut Document, part: PartId, rows: usize, cols: usize) -> NodeId {
    let tbl_name = doc.qn(part, "tbl");
    let tblpr_name = doc.qn(part, "tblPr");
    let tblw_name = doc.qn(part, "tblW");
    let tbllook_name = doc.qn(part, "tblLook");
    let tblgrid_name = doc.qn(part, "tblGrid");
    let gridcol_name = doc.qn(part, "gridCol");
    let tr_name = doc.qn(part, "tr");
    let type_attr = doc.qn(part, "type");
    let w_attr = doc.qn(part, "w");
    // w:tblLook attribute names, all in the WML (`w:`) prefix.
    let look_attrs: [(String, &str); 7] = [
        (doc.qn(part, "val"), "04A0"),
        (doc.qn(part, "firstRow"), "1"),
        (doc.qn(part, "lastRow"), "0"),
        (doc.qn(part, "firstColumn"), "1"),
        (doc.qn(part, "lastColumn"), "0"),
        (doc.qn(part, "noHBand"), "0"),
        (doc.qn(part, "noVBand"), "1"),
    ];

    // w:tblPr with w:tblW (auto) and a conventional w:tblLook.
    let tree = doc.tree_mut(part);
    let tblpr = tree.create_element(tblpr_name);
    let tblw = tree.create_element(tblw_name);
    tree.set_attr(tblw, type_attr, "auto");
    tree.set_attr(tblw, w_attr, "0");
    tree.append_child(tblpr, tblw);
    let tbllook = tree.create_element(tbllook_name);
    for (name, value) in look_attrs {
        tree.set_attr(tbllook, name, value);
    }
    tree.append_child(tblpr, tbllook);

    // w:tblGrid with `cols` bare gridCol.
    let tblgrid = tree.create_element(tblgrid_name);
    for _ in 0..cols {
        let gc = tree.create_element(gridcol_name.clone());
        tree.append_child(tblgrid, gc);
    }

    let tbl = tree.create_element(tbl_name);
    tree.append_child(tbl, tblpr);
    tree.append_child(tbl, tblgrid);

    for _ in 0..rows {
        let tr = doc.tree_mut(part).create_element(tr_name.clone());
        for _ in 0..cols {
            let tc = build_cell(doc, part);
            doc.tree_mut(part).append_child(tr, tc);
        }
        doc.tree_mut(part).append_child(tbl, tr);
    }
    tbl
}
