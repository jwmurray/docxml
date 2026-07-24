//! Milestone 5: tables — reading a body table and its cell texts, merge awareness plus
//! nested-table exclusion on the merged fixture, table creation and `add_row` growth, the
//! `set_text` single-paragraph replacement, and the untouched-save fidelity regression
//! across every fixture.

use std::path::{Path, PathBuf};

use docxml::opc::Package;
use docxml::xml::XmlTree;
use docxml::{Document, VMerge};

const BASIC: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");
const MERGED: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/tables_merged.docx"
);

/// Every `.docx` fixture in `tests/fixtures/`, sorted for stable test output.
fn all_fixtures() -> Vec<PathBuf> {
    let dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut fixtures: Vec<PathBuf> = std::fs::read_dir(&dir)
        .unwrap()
        .filter_map(|e| e.ok().map(|e| e.path()))
        .filter(|p| p.extension().is_some_and(|e| e == "docx"))
        .collect();
    fixtures.sort();
    fixtures
}

/// The main document part's bytes from a `.docx`.
fn document_xml(path: &Path) -> Vec<u8> {
    let pkg = Package::open(path).unwrap();
    pkg.part("word/document.xml").unwrap().data.clone()
}

/// Count all `w:tbl` elements anywhere in the tree (body-level *and* nested).
fn total_tbl_count(xml: &[u8]) -> usize {
    let tree = XmlTree::parse(xml).unwrap();
    tree.descendants(tree.root())
        .filter(|&id| tree.name(id) == Some("w:tbl"))
        .count()
}

// 1. Read a body table: one 3x3 table with the expected cell texts.
#[test]
fn reads_body_table_from_fixture() {
    let doc = Document::open(BASIC).unwrap();

    assert_eq!(doc.tables().len(), 1, "one body-level table expected");
    let table = doc.tables()[0];

    let rows = table.rows(&doc);
    assert_eq!(rows.len(), 3, "3 rows");
    for (r, row) in rows.iter().enumerate() {
        let cells = row.cells(&doc);
        assert_eq!(cells.len(), 3, "3 cells in row {r}");
        for (c, cell) in cells.iter().enumerate() {
            assert_eq!(cell.text(&doc), format!("r{r}c{c}"));
        }
    }
}

// 2. Merge awareness + nested-table exclusion on the merged fixture.
#[test]
fn merged_fixture_merges_and_nesting() {
    let doc = Document::open(MERGED).unwrap();

    // tables() returns only body-level tables; the fixture nests a table in a cell, so the
    // total w:tbl count strictly exceeds the body-level count.
    let total = total_tbl_count(&document_xml(Path::new(MERGED)));
    assert_eq!(
        total, 3,
        "two body tables + one nested table in the fixture"
    );
    assert_eq!(
        doc.tables().len(),
        2,
        "tables() excludes the nested table (body-level only)"
    );

    // First body table: a horizontally merged cell (gridSpan >= 2) and vertical
    // merge Restart/Continue cells.
    let first = doc.tables()[0];
    let all_cells: Vec<_> = first
        .rows(&doc)
        .iter()
        .flat_map(|row| row.cells(&doc))
        .collect();

    assert!(
        all_cells.iter().any(|c| c.grid_span(&doc) >= 2),
        "a horizontally merged cell (gridSpan >= 2) is present"
    );
    assert!(
        all_cells
            .iter()
            .any(|c| c.v_merge(&doc) == Some(VMerge::Restart)),
        "a vertical-merge Restart cell is present"
    );
    assert!(
        all_cells
            .iter()
            .any(|c| c.v_merge(&doc) == Some(VMerge::Continue)),
        "a vertical-merge Continue cell is present"
    );
    // Unmerged cells report grid_span 1 and no v_merge.
    let plain = all_cells
        .iter()
        .find(|c| c.text(&doc) == "ID")
        .expect("the 'ID' header cell");
    assert_eq!(plain.grid_span(&doc), 1);
    assert_eq!(plain.v_merge(&doc), None);

    // The second body table's cell that contains the nested table still reports its own
    // direct paragraphs without panicking.
    let second = doc.tables()[1];
    let nesting_cell = second
        .rows(&doc)
        .iter()
        .flat_map(|row| row.cells(&doc))
        .find(|c| c.text(&doc).contains("Nested table"))
        .expect("the cell introducing the nested table");
    let paras = nesting_cell.paragraphs(&doc);
    assert!(
        paras.iter().any(|p| p.text(&doc) == "Nested table:"),
        "the nesting cell's own paragraph reads back"
    );
}

// 3. Create a table, fill it, round-trip, then grow it with add_row.
#[test]
fn create_fill_roundtrip_and_add_row() {
    let mut doc = Document::new();
    let table = doc.add_table(2, 3);

    for r in 0..2 {
        let row = table.rows(&doc)[r];
        let cells = row.cells(&doc);
        assert_eq!(cells.len(), 3);
        for (c, cell) in cells.into_iter().enumerate() {
            cell.set_text(&mut doc, &format!("r{r}c{c}"));
        }
    }

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("created.docx");
    doc.save(&saved).unwrap();

    // The saved main part parses and its root is the document element.
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    assert_eq!(tree.name(tree.root()), Some("w:document"));

    let reopened = Document::open(&saved).unwrap();
    assert_eq!(reopened.tables().len(), 1);
    let table = reopened.tables()[0];
    let rows = table.rows(&reopened);
    assert_eq!(rows.len(), 2);
    for (r, row) in rows.iter().enumerate() {
        let cells = row.cells(&reopened);
        assert_eq!(cells.len(), 3);
        for (c, cell) in cells.iter().enumerate() {
            assert_eq!(cell.text(&reopened), format!("r{r}c{c}"));
        }
    }

    // add_row grows to 3 rows, each with 3 cells cloned from the grid column count.
    let mut doc2 = reopened;
    let table = doc2.tables()[0];
    table.add_row(&mut doc2);
    let table = doc2.tables()[0];
    assert_eq!(table.rows(&doc2).len(), 3);
    assert_eq!(table.rows(&doc2)[2].cells(&doc2).len(), 3);
}

// 4. set_text collapses multi-paragraph cell content to exactly one paragraph.
#[test]
fn set_text_replaces_multiple_paragraphs() {
    let mut doc = Document::new();
    let table = doc.add_table(1, 1);
    let cell = table.rows(&doc)[0].cells(&doc)[0];

    // Build a cell with two paragraphs of content.
    cell.set_text(&mut doc, "first");
    cell.add_paragraph(&mut doc, "second");
    assert_eq!(cell.paragraphs(&doc).len(), 2);
    assert_eq!(cell.text(&doc), "first\nsecond");

    // set_text replaces it all with one paragraph.
    cell.set_text(&mut doc, "replaced");
    assert_eq!(cell.paragraphs(&doc).len(), 1);
    assert_eq!(cell.text(&doc), "replaced");
}

// 5. Fidelity regression: untouched open -> save is byte-identical for every fixture.
#[test]
fn untouched_save_is_byte_identical_all_fixtures() {
    let dir = tempfile::tempdir().unwrap();
    for fixture in all_fixtures() {
        let doc = Document::open(&fixture).unwrap();
        let saved = dir.path().join("untouched.docx");
        doc.save(&saved).unwrap();

        let original = document_xml(&fixture);
        let round_tripped = document_xml(&saved);
        assert_eq!(
            original,
            round_tripped,
            "document.xml changed on untouched save of {}",
            fixture.display()
        );
    }
}
