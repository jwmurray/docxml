//! Milestone 15: table borders (`w:tblPr/w:tblBorders`) and cell borders
//! (`w:tcPr/w:tcBorders`).
//!
//! Each test drives the public API and, where the exact serialization is load-bearing (the
//! `CT_TblBorders`/`CT_TcBorders` child order), reopens the saved `.docx` and walks its
//! `word/document.xml` tree. The final test re-establishes the fidelity contract: an
//! untouched open→save is byte-identical across every fixture.

use std::path::{Path, PathBuf};

use docxml::opc::Package;
use docxml::xml::{NodeId, XmlTree};
use docxml::{BorderEdge, BorderStyle, Document};

/// The main document part's bytes from a `.docx`.
fn document_xml(path: &Path) -> Vec<u8> {
    let pkg = Package::open(path).unwrap();
    pkg.part("word/document.xml").unwrap().data.clone()
}

/// First element in the tree (pre-order) with the given qualified name.
fn find_first_named(tree: &XmlTree, qname: &str) -> Option<NodeId> {
    tree.descendants(tree.root())
        .find(|&n| tree.name(n) == Some(qname))
}

/// The local names of an element's element children, in document order (`w:top` → `top`).
fn child_locals(tree: &XmlTree, id: NodeId) -> Vec<String> {
    tree.children(id)
        .iter()
        .copied()
        .filter_map(|c| tree.name(c))
        .map(|n| n.rsplit(':').next().unwrap().to_string())
        .collect()
}

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

// 1. Full grid: set_borders(edge, inside) emits all six w:tblBorders children in
//    CT_TblBorders schema order and round-trips through save/reopen.
#[test]
fn table_full_grid_borders_order_and_roundtrip() {
    let mut doc = Document::new();
    let table = doc.add_table(2, 2);
    let edge = BorderEdge {
        style: BorderStyle::Single,
        size: 4,
        space: 0,
        color: None,
    };
    table.set_borders(&mut doc, Some(edge), Some(edge));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("grid.docx");
    doc.save(&saved).unwrap();

    // w:tblBorders children in CT_TblBorders order: top, left, bottom, right, insideH,
    // insideV (ECMA-376 §17.4.39; transitional left/right names).
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let borders = find_first_named(&tree, "w:tblBorders").unwrap();
    assert_eq!(
        child_locals(&tree, borders),
        vec!["top", "left", "bottom", "right", "insideH", "insideV"]
    );

    // Round-trip read-back: edge from top, inside from insideH.
    let reopened = Document::open(&saved).unwrap();
    let table = reopened.tables()[0];
    let (read_edge, read_inside) = table.borders(&reopened);
    assert_eq!(read_edge, Some(edge));
    assert_eq!(read_inside, Some(edge));
}

// 2. Cell override: one cell's set_borders(top only, Double) emits w:tcBorders with just
//    a w:top edge; sibling cells are untouched; round-trips.
#[test]
fn cell_border_override_order_and_roundtrip() {
    let mut doc = Document::new();
    let table = doc.add_table(2, 2);
    let top = BorderEdge {
        style: BorderStyle::Double,
        size: 12,
        space: 0,
        color: None,
    };
    let target = table.cell(&doc, 0, 0).unwrap();
    target.set_borders(&mut doc, Some(top), None, None, None);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("cell.docx");
    doc.save(&saved).unwrap();

    // Exactly one w:tcBorders, holding only w:top (CT_TcBorders §17.4.66).
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let all_tc_borders: Vec<NodeId> = tree
        .descendants(tree.root())
        .filter(|&n| tree.name(n) == Some("w:tcBorders"))
        .collect();
    assert_eq!(all_tc_borders.len(), 1, "only the one overridden cell");
    assert_eq!(child_locals(&tree, all_tc_borders[0]), vec!["top"]);

    // Round-trip: the overridden cell reads back its top edge and nothing else; the other
    // cells have no borders.
    let reopened = Document::open(&saved).unwrap();
    let table = reopened.tables()[0];
    assert_eq!(
        table.cell(&reopened, 0, 0).unwrap().borders(&reopened),
        (Some(top), None, None, None)
    );
    for (r, c) in [(0, 1), (1, 0), (1, 1)] {
        assert_eq!(
            table.cell(&reopened, r, c).unwrap().borders(&reopened),
            (None, None, None, None),
            "cell ({r},{c}) untouched"
        );
    }
}

// 2b. Cell set_borders rebuilds the four managed edges in CT_TcBorders order regardless of
//     argument order, and clearing (all None) removes w:tcBorders.
#[test]
fn cell_all_edges_order_and_removal() {
    let mut doc = Document::new();
    let table = doc.add_table(1, 1);
    let edge = BorderEdge {
        style: BorderStyle::Single,
        size: 4,
        space: 0,
        color: None,
    };
    let cell = table.cell(&doc, 0, 0).unwrap();
    cell.set_borders(&mut doc, Some(edge), Some(edge), Some(edge), Some(edge));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("cellfull.docx");
    doc.save(&saved).unwrap();
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let borders = find_first_named(&tree, "w:tcBorders").unwrap();
    assert_eq!(
        child_locals(&tree, borders),
        vec!["top", "left", "bottom", "right"]
    );

    // Clear: all None removes w:tcBorders entirely.
    cell.set_borders(&mut doc, None, None, None, None);
    assert_eq!(cell.borders(&doc), (None, None, None, None));
    let saved2 = dir.path().join("cellcleared.docx");
    doc.save(&saved2).unwrap();
    let tree2 = XmlTree::parse(&document_xml(&saved2)).unwrap();
    assert!(find_first_named(&tree2, "w:tcBorders").is_none());
}

// 3. Removal: Table::set_borders(None, None) removes w:tblBorders; and the untouched-save
//    fidelity regression stays byte-identical across every fixture.
#[test]
fn table_border_removal_and_fixture_fidelity() {
    let mut doc = Document::new();
    let table = doc.add_table(2, 2);
    let edge = BorderEdge {
        style: BorderStyle::Single,
        size: 4,
        space: 0,
        color: None,
    };
    table.set_borders(&mut doc, Some(edge), Some(edge));
    // Now remove.
    table.set_borders(&mut doc, None, None);
    assert_eq!(table.borders(&doc), (None, None));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("removed.docx");
    doc.save(&saved).unwrap();
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    assert!(find_first_named(&tree, "w:tblBorders").is_none());

    // Fidelity regression: an untouched open -> save stays byte-identical for every fixture.
    for fixture in all_fixtures() {
        let doc = Document::open(&fixture).unwrap();
        let out = dir.path().join("untouched.docx");
        doc.save(&out).unwrap();
        assert_eq!(
            document_xml(&fixture),
            document_xml(&out),
            "document.xml changed on untouched save of {}",
            fixture.display()
        );
    }
}
