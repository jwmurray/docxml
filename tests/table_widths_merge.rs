//! Milestone 11: table column widths, grid-based cell addressing, and merge creation.
//!
//! Covers `set_column_widths`/`column_widths`/`Cell::width`/`set_fixed_layout`, the
//! grid-coordinate `Table::cell` (span and vMerge-continuation resolution) against the
//! merged fixture, horizontal / vertical / rectangular `Table::merge` with content
//! preservation and round-trip, a partial-overlap error, and the untouched-save fidelity
//! regression across every fixture.

use std::path::{Path, PathBuf};

use docxml::opc::Package;
use docxml::xml::XmlTree;
use docxml::{Document, Length, VMerge};

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

/// Save `doc` to a fresh file under `dir` and reopen it.
fn round_trip(doc: &Document, dir: &Path) -> Document {
    let saved = dir.join("rt.docx");
    doc.save(&saved).unwrap();
    Document::open(&saved).unwrap()
}

/// True when the saved main part contains a `w:tblLayout` with `w:type="fixed"`.
fn has_fixed_layout(xml: &[u8]) -> bool {
    let tree = XmlTree::parse(xml).unwrap();
    tree.descendants(tree.root())
        .any(|id| tree.name(id) == Some("w:tblLayout") && tree.attr(id, "w:type") == Some("fixed"))
}

// 1. Column widths: gridCol w:w and per-cell tcW are set; read back; fixed layout emits.
#[test]
fn set_column_widths_writes_grid_and_cells_and_roundtrips() {
    let mut doc = Document::new();
    let table = doc.add_table(2, 3);
    assert_eq!(table.column_count(&doc), 3);

    let widths = [
        Length::from_inches(1.5), // 2160 twips
        Length::from_inches(2.0), // 2880 twips
        Length::from_inches(3.0), // 4320 twips
    ];
    table.set_column_widths(&mut doc, &widths);

    // gridCol w:w values.
    let cols = table.column_widths(&doc);
    assert_eq!(cols, widths.iter().copied().map(Some).collect::<Vec<_>>());
    assert_eq!(
        cols.iter().map(|w| w.unwrap().twips()).collect::<Vec<_>>(),
        [2160, 2880, 4320]
    );

    // Every physical cell's tcW (dxa) matches its grid column width.
    for row in table.rows(&doc) {
        for (c, cell) in row.cells(&doc).into_iter().enumerate() {
            assert_eq!(
                cell.width(&doc),
                Some(widths[c]),
                "cell width in column {c}"
            );
        }
    }

    // Round-trip: the widths survive save + reopen.
    let dir = tempfile::tempdir().unwrap();
    let reopened = round_trip(&doc, dir.path());
    let table = reopened.tables()[0];
    assert_eq!(
        table.column_widths(&reopened),
        widths.iter().copied().map(Some).collect::<Vec<_>>()
    );
    assert_eq!(
        table.rows(&reopened)[0].cells(&reopened)[1].width(&reopened),
        Some(widths[1])
    );

    // set_fixed_layout(true) emits w:tblLayout w:type="fixed".
    let mut doc2 = reopened;
    let table = doc2.tables()[0];
    table.set_fixed_layout(&mut doc2, true);
    let saved = dir.path().join("fixed.docx");
    doc2.save(&saved).unwrap();
    assert!(has_fixed_layout(&document_xml(&saved)));
}

// 2. Grid-based addressing on tables_merged.docx: span and vMerge resolution, out of range.
#[test]
fn grid_addressing_resolves_spans_and_vmerge() {
    let doc = Document::open(MERGED).unwrap();
    let table = doc.tables()[0];
    assert_eq!(table.column_count(&doc), 4);

    // Horizontal merge in row 1 (gridSpan=2 over grid columns 1 and 2): both grid columns
    // resolve to the single leading (col-1) cell.
    let c11 = table.cell(&doc, 1, 1).expect("grid (1,1)");
    let c12 = table.cell(&doc, 1, 2).expect("grid (1,2)");
    assert_eq!(c11, c12, "both spanned grid columns are the same cell node");
    assert!(c11.text(&doc).contains("Widget"));
    assert_eq!(c11.grid_span(&doc), 2);

    // A distinct grid column (col 0) is a different cell.
    assert_ne!(table.cell(&doc, 1, 0).unwrap(), c11);

    // Vertical merge in grid column 0: the continuation row resolves up to the restart.
    let restart = table.cell(&doc, 2, 0).expect("grid (2,0) restart");
    let continued = table.cell(&doc, 3, 0).expect("grid (3,0) continuation");
    assert_eq!(
        restart, continued,
        "continuation resolves to its restart origin"
    );
    assert_eq!(restart.v_merge(&doc), Some(VMerge::Restart));
    assert!(restart.text(&doc).contains("2"));

    // Out of range → None.
    assert_eq!(table.cell(&doc, 0, 4), None, "column past the grid");
    assert_eq!(table.cell(&doc, 4, 0), None, "row past the table");
}

// 3. Horizontal merge: three cells collapse to one gridSpan-3 cell keeping all content.
#[test]
fn horizontal_merge_collapses_row_and_preserves_content() {
    let mut doc = Document::new();
    let table = doc.add_table(3, 3);

    for c in 0..3 {
        table
            .cell(&doc, 0, c)
            .unwrap()
            .set_text(&mut doc, &format!("h{c}"));
    }

    let origin = table.merge(&mut doc, (0, 0), (0, 2)).unwrap();
    assert_eq!(origin.grid_span(&doc), 3);

    // Row 0 now has one physical cell carrying all three texts.
    let row0 = table.rows(&doc)[0].cells(&doc);
    assert_eq!(row0.len(), 1);
    let text = row0[0].text(&doc);
    for c in 0..3 {
        assert!(
            text.contains(&format!("h{c}")),
            "merged text keeps h{c}: {text:?}"
        );
    }

    // Rows 1 and 2 are untouched (still three physical cells).
    assert_eq!(table.rows(&doc)[1].cells(&doc).len(), 3);
    assert_eq!(table.rows(&doc)[2].cells(&doc).len(), 3);

    // Round-trip, then grid addressing still resolves the spanned columns to the origin.
    let dir = tempfile::tempdir().unwrap();
    let reopened = round_trip(&doc, dir.path());
    let table = reopened.tables()[0];
    let o = table.cell(&reopened, 0, 0).unwrap();
    assert_eq!(table.cell(&reopened, 0, 1).unwrap(), o);
    assert_eq!(table.cell(&reopened, 0, 2).unwrap(), o);
    assert_eq!(o.grid_span(&reopened), 3);
}

// 4. Vertical merge: a column becomes a restart origin over bare-vMerge continuations.
#[test]
fn vertical_merge_stacks_column_and_empties_continuations() {
    let mut doc = Document::new();
    let table = doc.add_table(3, 2);

    for r in 0..3 {
        table
            .cell(&doc, r, 0)
            .unwrap()
            .set_text(&mut doc, &format!("v{r}"));
    }

    let origin = table.merge(&mut doc, (0, 0), (2, 0)).unwrap();
    assert_eq!(origin.v_merge(&doc), Some(VMerge::Restart));

    // Origin carries every row's content.
    let text = origin.text(&doc);
    for r in 0..3 {
        assert!(
            text.contains(&format!("v{r}")),
            "origin keeps v{r}: {text:?}"
        );
    }

    // The lower rows' column-0 cells are bare-vMerge continuations, each with one empty w:p.
    for r in 1..3 {
        let cont = table.rows(&doc)[r].cells(&doc)[0];
        assert_eq!(cont.v_merge(&doc), Some(VMerge::Continue));
        assert_eq!(
            cont.paragraphs(&doc).len(),
            1,
            "continuation keeps one paragraph"
        );
        assert_eq!(cont.text(&doc), "", "continuation paragraph is empty");
    }

    // Round-trip, then the continuation still resolves up to the origin.
    let dir = tempfile::tempdir().unwrap();
    let reopened = round_trip(&doc, dir.path());
    let table = reopened.tables()[0];
    assert_eq!(
        table.cell(&reopened, 2, 0).unwrap(),
        table.cell(&reopened, 0, 0).unwrap()
    );
}

// 5. Rectangular 2x2 merge (horizontal + vertical), and a clean partial-overlap error.
#[test]
fn rectangular_merge_and_partial_overlap_error() {
    let mut doc = Document::new();
    let table = doc.add_table(3, 3);
    for r in 0..2 {
        for c in 0..2 {
            table
                .cell(&doc, r, c)
                .unwrap()
                .set_text(&mut doc, &format!("r{r}c{c}"));
        }
    }

    let origin = table.merge(&mut doc, (0, 0), (1, 1)).unwrap();
    // Origin spans two columns and restarts a vertical merge.
    assert_eq!(origin.grid_span(&doc), 2);
    assert_eq!(origin.v_merge(&doc), Some(VMerge::Restart));
    // The lower row's leading cell is a 2-wide continuation.
    let lower = table.rows(&doc)[1].cells(&doc)[0];
    assert_eq!(lower.grid_span(&doc), 2);
    assert_eq!(lower.v_merge(&doc), Some(VMerge::Continue));
    // Origin gathered all four source texts.
    let text = origin.text(&doc);
    for r in 0..2 {
        for c in 0..2 {
            assert!(text.contains(&format!("r{r}c{c}")), "origin keeps r{r}c{c}");
        }
    }
    // Round-trips cleanly.
    let dir = tempfile::tempdir().unwrap();
    let _ = round_trip(&doc, dir.path());

    // Partial overlap: merge a horizontal pair, then a region straddling its boundary.
    let mut doc2 = Document::new();
    let table2 = doc2.add_table(3, 3);
    table2.merge(&mut doc2, (0, 0), (0, 1)).unwrap();
    let err = table2.merge(&mut doc2, (0, 1), (1, 2)).unwrap_err();
    assert!(
        matches!(err, docxml::Error::InvalidMerge(_)),
        "partial overlap is an InvalidMerge, got {err:?}"
    );
    // Out-of-range region is also an InvalidMerge.
    assert!(matches!(
        table2.merge(&mut doc2, (0, 0), (0, 9)).unwrap_err(),
        docxml::Error::InvalidMerge(_)
    ));
}

// 6. Fidelity regression: untouched open -> save is byte-identical for every fixture.
#[test]
fn untouched_save_is_byte_identical_all_fixtures() {
    let dir = tempfile::tempdir().unwrap();
    for fixture in all_fixtures() {
        let doc = Document::open(&fixture).unwrap();
        let saved = dir.path().join("untouched.docx");
        doc.save(&saved).unwrap();
        assert_eq!(
            document_xml(&fixture),
            document_xml(&saved),
            "document.xml changed on untouched save of {}",
            fixture.display()
        );
    }
}
