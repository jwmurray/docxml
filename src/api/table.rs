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
//!
//! # Grid-based addressing
//!
//! [`Table::cell`] complements the physical [`Row::cells`] with python-docx's *grid*
//! addressing: `cell(row, col)` walks a row's physical `w:tc`, accumulating `w:gridSpan`,
//! to find the physical cell covering grid column `col`, and resolves a `w:vMerge`
//! continuation up to its `w:val="restart"` origin — so every grid coordinate covered by a
//! span yields the one underlying cell object, exactly as python-docx's `Table.cell` does.
//! [`Table::merge`] creates merges over a grid-coordinate region.

use crate::error::{Error, Result};
use crate::xml::{NodeId, XmlTree};

use super::{Document, Length, Paragraph, PartId, is_wml_element, ordered_insert_index, rank_in};

/// Canonical `w:tcPr` child order (ECMA-376 §17.4.70, `CT_TcPr` sequence), local names
/// only. New properties are inserted to keep `w:tcPr`'s children in this order so the
/// output is schema-valid — of interest to this milestone: `tcW`, then `gridSpan`, then
/// `vMerge`. Unlisted children rank last and stay after authored properties.
const TCPR_ORDER: &[&str] = &[
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFit",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
];

/// Canonical `w:tblPr` child order (ECMA-376 §17.4.60, `CT_TblPr` sequence), local names
/// only. Used to slot `w:tblLayout` (fixed/autofit) into schema position. Unlisted children
/// rank last.
const TBLPR_ORDER: &[&str] = &[
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
];

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

    /// The number of columns declared by the table grid — its `w:tblGrid/w:gridCol` count.
    ///
    /// This is the *grid* column count, the coordinate space [`cell`](Self::cell) and
    /// [`merge`](Self::merge) address. A horizontally merged (`w:gridSpan`) physical cell
    /// still occupies its full span of grid columns here, so a row's physical
    /// [`Row::cells`] count can be smaller than this.
    pub fn column_count(&self, doc: &Document) -> usize {
        self.grid_col_count(doc.tree(self.part))
    }

    /// The `w:tblGrid` element, if present.
    fn tbl_grid(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "tblGrid"))
    }

    /// The `w:gridCol` children of `w:tblGrid`, in order.
    fn grid_cols(&self, tree: &XmlTree) -> Vec<NodeId> {
        let Some(grid) = self.tbl_grid(tree) else {
            return Vec::new();
        };
        tree.children(grid)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "gridCol"))
            .collect()
    }

    /// The `w:tblPr` element, creating it (as the table's first child, per `CT_Tbl`) if
    /// absent.
    fn ensure_tbl_pr(&self, doc: &mut Document) -> NodeId {
        let tree = doc.tree(self.part);
        if let Some(existing) = tree
            .children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "tblPr"))
        {
            return existing;
        }
        let name = doc.qn(self.part, "tblPr");
        let tree = doc.tree_mut(self.part);
        let el = tree.create_element(name);
        tree.insert_child(self.node, 0, el);
        el
    }

    /// The declared grid-column widths, read from each `w:tblGrid/w:gridCol`'s `w:w`
    /// (twips). A `w:gridCol` with a missing or non-integer `w:w` reads as `None`; the
    /// returned vector has one entry per grid column, in order.
    pub fn column_widths(&self, doc: &Document) -> Vec<Option<Length>> {
        let tree = doc.tree(self.part);
        let w_attr = doc.qn(self.part, "w");
        self.grid_cols(tree)
            .into_iter()
            .map(|gc| tree.attr(gc, &w_attr).and_then(Length::from_twips_str))
            .collect()
    }

    /// Set the width of every grid column, matching python-docx's `Column.width` behavior:
    /// each `w:gridCol`'s `w:w` *and* every physical cell's `w:tcPr/w:tcW` (`w:type="dxa"`)
    /// in that grid column are set to the corresponding width in twips.
    ///
    /// Widths are applied grid-column by grid-column, so a horizontally spanned cell that
    /// covers several grid columns has its `w:tcW` set once per column it covers — the last
    /// (right-most) column's width wins, as in python-docx.
    ///
    /// Note that Word may recompute visual widths under its default *autofit* table layout;
    /// call [`set_fixed_layout(true)`](Self::set_fixed_layout) to pin the layout so the set
    /// widths are honored. python-docx does not set fixed layout on a width change, and
    /// neither does this method.
    ///
    /// # Panics
    /// Panics if `widths.len()` is not equal to [`column_count`](Self::column_count) — the
    /// same assert-on-bad-input convention as [`Document::add_heading`](crate::Document::add_heading).
    pub fn set_column_widths(&self, doc: &mut Document, widths: &[Length]) {
        let count = self.column_count(doc);
        assert!(
            widths.len() == count,
            "set_column_widths expects one width per grid column ({count}), got {}",
            widths.len()
        );

        let grid_cols = self.grid_cols(doc.tree(self.part));
        let w_attr = doc.qn(self.part, "w");
        let rows = self.rows(doc);
        for (i, width) in widths.iter().enumerate() {
            let twips = width.to_twips_string();
            doc.tree_mut(self.part)
                .set_attr(grid_cols[i], w_attr.clone(), twips);
            // Set the covering physical cell's tcW in each row for this grid column.
            for &row in &rows {
                if let Some((cell, _, _)) = physical_cell_covering(doc, row, i as u32) {
                    cell.set_width(doc, *width);
                }
            }
        }
    }

    /// Set the table layout algorithm (`w:tblPr/w:tblLayout`): `true` writes
    /// `w:type="fixed"` (Word honors the declared column widths verbatim), `false` writes
    /// `w:type="autofit"` (Word may resize columns to fit content).
    ///
    /// `w:tblLayout` sits in `CT_TblPr` schema order (present in [`TBLPR_ORDER`]).
    pub fn set_fixed_layout(&self, doc: &mut Document, fixed: bool) {
        let tblpr = self.ensure_tbl_pr(doc);
        let layout = ensure_ordered_child(doc, self.part, tblpr, "tblLayout", TBLPR_ORDER);
        let type_attr = doc.qn(self.part, "type");
        let val = if fixed { "fixed" } else { "autofit" };
        doc.tree_mut(self.part).set_attr(layout, type_attr, val);
    }

    /// The cell covering grid coordinate (`row`, `col`), or `None` when out of range.
    ///
    /// These are *grid* coordinates, not physical-cell indices. The row's physical `w:tc`
    /// are walked accumulating `w:gridSpan` to find the one covering grid column `col`; if
    /// that cell is a `w:vMerge` continuation, the merge is resolved upward and the
    /// `w:val="restart"` origin cell (in the same grid column) is returned. Every grid
    /// coordinate a merged cell spans therefore yields the single origin [`Cell`], mirroring
    /// python-docx's `Table.cell`.
    ///
    /// Contrast [`Row::cells`], which returns the row's raw physical cells (a span appears
    /// once, a continuation is a distinct cell). Out-of-range `row` or `col` → `None`.
    pub fn cell(&self, doc: &Document, row: usize, col: usize) -> Option<Cell> {
        let rows = self.rows(doc);
        let target = *rows.get(row)?;
        let col = u32::try_from(col).ok()?;
        let (cell, _start, _span) = physical_cell_covering(doc, target, col)?;

        // A continuation resolves up to its restart origin in the same grid column.
        if cell.v_merge(doc) != Some(VMerge::Continue) {
            return Some(cell);
        }
        for r in (0..row).rev() {
            let up = physical_cell_covering(doc, rows[r], col)?;
            if up.0.v_merge(doc) != Some(VMerge::Continue) {
                return Some(up.0);
            }
        }
        // Malformed table (a continuation with no restart above); return the cell itself.
        Some(cell)
    }

    /// Merge the rectangular grid region from `top_left` to `bottom_right` (inclusive,
    /// `(row, col)` grid coordinates) into a single cell, and return that origin [`Cell`].
    ///
    /// This is python-docx's `cell.merge(other_cell)`: content is preserved (the block
    /// content of every spanned cell moves into the origin), horizontal spans collapse to a
    /// single `w:tc` carrying `w:gridSpan`, and vertical spans become a `w:vMerge`
    /// `w:val="restart"` origin over bare `w:vMerge` continuation cells (each left with one
    /// empty `w:p`). When every merged grid column has a known width, the origin's `w:tcW`
    /// is set to their sum.
    ///
    /// # Errors
    /// Returns [`Error::InvalidMerge`] when the region is out of range, or when it would
    /// *partially* overlap a pre-existing merge (an existing `w:gridSpan`/`w:vMerge` region
    /// straddling the requested boundary) — the region must be rectangular in grid space.
    pub fn merge(
        &self,
        doc: &mut Document,
        top_left: (usize, usize),
        bottom_right: (usize, usize),
    ) -> Result<Cell> {
        let rows = self.rows(doc);
        let nrows = rows.len();
        let ncols = self.column_count(doc);

        let r0 = top_left.0.min(bottom_right.0);
        let r1 = top_left.0.max(bottom_right.0);
        let c0 = top_left.1.min(bottom_right.1);
        let c1 = top_left.1.max(bottom_right.1);
        if r1 >= nrows || c1 >= ncols {
            return Err(Error::InvalidMerge(format!(
                "region ({r0},{c0})..=({r1},{c1}) is out of range for a {nrows}x{ncols} grid"
            )));
        }
        self.check_no_partial_overlap(doc, &rows, r0, r1, c0 as u32, c1 as u32)?;

        let width = (c1 - c0 + 1) as u32;
        let (col0, col1) = (c0 as u32, c1 as u32);

        // Sum of the merged grid columns' widths, if all are known.
        let summed_width: Option<Length> = {
            let widths = self.column_widths(doc);
            let slice = &widths[c0..=c1];
            if slice.iter().all(Option::is_some) {
                Some(Length::from_twips(
                    slice.iter().map(|w| w.unwrap().twips()).sum(),
                ))
            } else {
                None
            }
        };

        // Horizontal pass: on each row of the region, collapse the cells covering grid
        // columns col0..=col1 into their leading cell (moving content), then set gridSpan.
        let mut leading_per_row: Vec<Cell> = Vec::with_capacity(r1 - r0 + 1);
        for &row in &rows[r0..=r1] {
            let covered = physical_cells_in_span(doc, row, col0, col1);
            let leading = covered[0].0;
            for &(other, _, _) in &covered[1..] {
                move_cell_content(doc, self.part, other.node, leading.node);
                doc.tree_mut(self.part).remove_from_parent(other.node);
            }
            if width > 1 {
                set_grid_span(doc, self.part, leading, width);
            }
            if let Some(w) = summed_width {
                leading.set_width(doc, w);
            }
            leading_per_row.push(leading);
        }

        // Vertical pass: fold lower rows' leading cells into the top (origin) cell.
        let origin = leading_per_row[0];
        if r1 > r0 {
            set_v_merge(doc, self.part, origin, true);
            for &lower in &leading_per_row[1..] {
                move_cell_content(doc, self.part, lower.node, origin.node);
                // A cell must keep at least one w:p; leave the continuation with one empty.
                add_empty_paragraph(doc, self.part, lower.node);
                set_v_merge(doc, self.part, lower, false);
            }
        }
        Ok(origin)
    }

    /// Error unless the region `[r0..=r1] x [c0..=c1]` (grid coordinates) contains only
    /// whole cells — no existing merged cell straddles its boundary.
    fn check_no_partial_overlap(
        &self,
        doc: &Document,
        rows: &[Row],
        r0: usize,
        r1: usize,
        c0: u32,
        c1: u32,
    ) -> Result<()> {
        for (ri, &row) in rows.iter().enumerate() {
            for (cell, start, span) in physical_cells_of_row(doc, row) {
                // Skip vMerge continuations — they belong to their restart's rectangle.
                if cell.v_merge(doc) == Some(VMerge::Continue) {
                    continue;
                }
                let min_col = start;
                let max_col = start + span - 1;
                let min_row = ri;
                let max_row = if cell.v_merge(doc) == Some(VMerge::Restart) {
                    vertical_extent(doc, rows, ri, start)
                } else {
                    ri
                };

                let intersects = !(max_row < r0 || min_row > r1 || max_col < c0 || min_col > c1);
                let contained = min_row >= r0 && max_row <= r1 && min_col >= c0 && max_col <= c1;
                if intersects && !contained {
                    return Err(Error::InvalidMerge(format!(
                        "region ({r0},{c0})..=({r1},{c1}) partially overlaps an existing \
                         merged cell spanning rows {min_row}..={max_row}, columns \
                         {min_col}..={max_col}"
                    )));
                }
            }
        }
        Ok(())
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

    /// The cell's width from `w:tcPr/w:tcW`, or `None` when unset.
    ///
    /// Only an absolute width (`w:type="dxa"`, in twips) is returned as a [`Length`]. An
    /// `"auto"` type, a missing `w:tcW`, and — because [`Length`] models an absolute
    /// measure — a percentage (`w:type="pct"`) all read as `None`.
    pub fn width(&self, doc: &Document) -> Option<Length> {
        let tree = doc.tree(self.part);
        let tcw = self.tc_pr_child(tree, "tcW")?;
        match tree.attr(tcw, &doc.qn(self.part, "type")) {
            Some("dxa") => tree
                .attr(tcw, &doc.qn(self.part, "w"))
                .and_then(Length::from_twips_str),
            _ => None,
        }
    }

    /// Set the cell's absolute width, writing `w:tcPr/w:tcW` with `w:type="dxa"` and the
    /// width in twips (creating `w:tcPr`/`w:tcW` in schema order if absent).
    pub fn set_width(&self, doc: &mut Document, width: Length) {
        let tcw = self.ensure_tc_pr_child(doc, "tcW");
        let w_attr = doc.qn(self.part, "w");
        let type_attr = doc.qn(self.part, "type");
        let tree = doc.tree_mut(self.part);
        tree.set_attr(tcw, w_attr, width.to_twips_string());
        tree.set_attr(tcw, type_attr, "dxa");
    }

    /// The cell's `w:tcPr`, if present.
    fn tc_pr(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "tcPr"))
    }

    /// The cell's `w:tcPr`, creating it (as the cell's first child, per `CT_Tc`) if absent.
    fn ensure_tc_pr(&self, doc: &mut Document) -> NodeId {
        if let Some(existing) = self.tc_pr(doc.tree(self.part)) {
            return existing;
        }
        let name = doc.qn(self.part, "tcPr");
        let tree = doc.tree_mut(self.part);
        let el = tree.create_element(name);
        tree.insert_child(self.node, 0, el);
        el
    }

    /// A direct `w:tcPr` child with the given local name, creating it in canonical
    /// `CT_TcPr` schema order if absent.
    fn ensure_tc_pr_child(&self, doc: &mut Document, local: &str) -> NodeId {
        let tcpr = self.ensure_tc_pr(doc);
        ensure_ordered_child(doc, self.part, tcpr, local, TCPR_ORDER)
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

/// A direct child of `parent` with WML local name `local`, creating it in canonical schema
/// `order` if absent. Shared by `w:tcPr` and `w:tblPr` property insertion.
fn ensure_ordered_child(
    doc: &mut Document,
    part: PartId,
    parent: NodeId,
    local: &str,
    order: &[&str],
) -> NodeId {
    let tree = doc.tree(part);
    if let Some(existing) = tree
        .children(parent)
        .iter()
        .copied()
        .find(|&c| is_wml_element(tree, c, local))
    {
        return existing;
    }
    let name = doc.qn(part, local);
    let index = ordered_insert_index(doc.tree(part), parent, rank_in(order, local), order);
    let tree = doc.tree_mut(part);
    let el = tree.create_element(name);
    tree.insert_child(parent, index, el);
    el
}

/// A row's physical cells paired with the grid column each *starts* at and its grid span:
/// `(cell, grid_start, grid_span)`, left to right.
fn physical_cells_of_row(doc: &Document, row: Row) -> Vec<(Cell, u32, u32)> {
    let mut pos = 0u32;
    let mut out = Vec::new();
    for cell in row.cells(doc) {
        let span = cell.grid_span(doc);
        out.push((cell, pos, span));
        pos += span;
    }
    out
}

/// The physical cell of `row` covering grid column `col`, with its grid start and span.
fn physical_cell_covering(doc: &Document, row: Row, col: u32) -> Option<(Cell, u32, u32)> {
    physical_cells_of_row(doc, row)
        .into_iter()
        .find(|&(_, start, span)| col >= start && col < start + span)
}

/// The physical cells of `row` that overlap the grid-column span `c0..=c1`, left to right;
/// with a rectangular (non-partial) region the leading entry starts exactly at `c0`.
fn physical_cells_in_span(doc: &Document, row: Row, c0: u32, c1: u32) -> Vec<(Cell, u32, u32)> {
    physical_cells_of_row(doc, row)
        .into_iter()
        .filter(|&(_, start, span)| start <= c1 && start + span > c0)
        .collect()
}

/// The last row index a vertical merge restarting at (`restart_row`, `col`) extends to,
/// following bare `w:vMerge` continuation cells anchored at the same grid column.
fn vertical_extent(doc: &Document, rows: &[Row], restart_row: usize, col: u32) -> usize {
    let mut max_row = restart_row;
    for (r, &row) in rows.iter().enumerate().skip(restart_row + 1) {
        match physical_cell_covering(doc, row, col) {
            Some((cell, start, _))
                if start == col && cell.v_merge(doc) == Some(VMerge::Continue) =>
            {
                max_row = r;
            }
            _ => break,
        }
    }
    max_row
}

/// Move `src`'s block-level content (every child except `w:tcPr`) to the end of `dst`,
/// preserving order. Used by [`Table::merge`] to keep spanned cells' content.
fn move_cell_content(doc: &mut Document, part: PartId, src: NodeId, dst: NodeId) {
    let tree = doc.tree(part);
    let to_move: Vec<NodeId> = tree
        .children(src)
        .iter()
        .copied()
        .filter(|&c| !is_wml_element(tree, c, "tcPr"))
        .collect();
    for child in to_move {
        doc.tree_mut(part).remove_from_parent(child);
        doc.tree_mut(part).append_child(dst, child);
    }
}

/// Append one empty `w:p` to `cell_node` (a cell must always keep at least one paragraph).
fn add_empty_paragraph(doc: &mut Document, part: PartId, cell_node: NodeId) {
    let name = doc.qn(part, "p");
    let tree = doc.tree_mut(part);
    let p = tree.create_element(name);
    tree.append_child(cell_node, p);
}

/// Set `cell`'s `w:tcPr/w:gridSpan` to `span` (creating the element in schema order).
fn set_grid_span(doc: &mut Document, part: PartId, cell: Cell, span: u32) {
    let gs = cell.ensure_tc_pr_child(doc, "gridSpan");
    let val = doc.qn(part, "val");
    doc.tree_mut(part).set_attr(gs, val, span.to_string());
}

/// Set `cell`'s `w:tcPr/w:vMerge`: a restart origin (`w:val="restart"`) or a bare
/// continuation (`w:val` removed), creating the element in schema order.
fn set_v_merge(doc: &mut Document, part: PartId, cell: Cell, restart: bool) {
    let vm = cell.ensure_tc_pr_child(doc, "vMerge");
    let val = doc.qn(part, "val");
    if restart {
        doc.tree_mut(part).set_attr(vm, val, "restart");
    } else {
        doc.tree_mut(part).remove_attr(vm, &val);
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
