//! Typed `w:docGrid` and `w:pgNumType`: authoring, read-back, schema order, and the
//! author → save → reopen round trip.

use std::io::Cursor;

use docxml::{DocGrid, DocGridType, Document, LineNumbering, NumberFormat, PageNumbering, Pt};

fn reopen(doc: &Document) -> Document {
    let mut buf = Cursor::new(Vec::new());
    doc.write(&mut buf).expect("document writes");
    Document::read(Cursor::new(buf.into_inner())).expect("written document reopens")
}

#[test]
fn doc_grid_round_trips_with_line_pitch() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    // The blank template already carries Word's default grid — a bare
    // `<w:docGrid w:linePitch="360"/>` (18pt pitch, type defaulted). The typed
    // read sees it, and the setter below replaces it in place.
    let template_grid = section.doc_grid(&doc).expect("template ships a docGrid");
    assert_eq!(template_grid.grid_type, DocGridType::Default);
    assert_eq!(template_grid.line_pitch, Some(Pt(18.0)));

    section.set_doc_grid(
        &mut doc,
        DocGrid {
            grid_type: DocGridType::Lines,
            line_pitch: Some(Pt(23.1)),
            char_space: None,
        },
    );

    let reopened = reopen(&doc);
    let grid = reopened.sections()[0]
        .doc_grid(&reopened)
        .expect("doc grid persists");
    assert_eq!(grid.grid_type, DocGridType::Lines);
    assert_eq!(grid.line_pitch, Some(Pt(23.1)));
    assert_eq!(grid.char_space, None);

    // The stored value is exactly 462 twentieths of a point — the pitch that fits
    // 28 lines into a nine-inch body where plain double spacing (480) stops at 27.
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    assert!(
        xml.contains(r#"<w:docGrid w:type="lines" w:linePitch="462"/>"#),
        "docGrid element shape: {}",
        &xml[xml.find("docGrid").unwrap_or(0)
            ..(xml.find("docGrid").unwrap_or(0) + 80).min(xml.len())]
    );
}

#[test]
fn doc_grid_reset_clears_stale_attributes() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    section.set_doc_grid(
        &mut doc,
        DocGrid {
            grid_type: DocGridType::LinesAndChars,
            line_pitch: Some(Pt(24.0)),
            char_space: Some(4320),
        },
    );
    // Re-set with fewer fields: the old linePitch/charSpace must not survive.
    section.set_doc_grid(
        &mut doc,
        DocGrid {
            grid_type: DocGridType::Default,
            line_pitch: None,
            char_space: None,
        },
    );
    let grid = section.doc_grid(&doc).expect("doc grid present");
    assert_eq!(grid.grid_type, DocGridType::Default);
    assert_eq!(grid.line_pitch, None);
    assert_eq!(grid.char_space, None);
}

#[test]
fn doc_grid_clears() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    section.set_doc_grid(
        &mut doc,
        DocGrid {
            grid_type: DocGridType::Lines,
            line_pitch: Some(Pt(23.1)),
            char_space: None,
        },
    );
    section.clear_doc_grid(&mut doc);
    assert!(section.doc_grid(&doc).is_none());
}

#[test]
fn page_numbering_round_trips() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    assert!(section.page_numbering(&doc).is_none());

    section.set_page_numbering(
        &mut doc,
        PageNumbering {
            format: Some(NumberFormat::LowerRoman),
            start: Some(1),
        },
    );

    let reopened = reopen(&doc);
    let pn = reopened.sections()[0]
        .page_numbering(&reopened)
        .expect("pgNumType persists");
    assert_eq!(pn.format, Some(NumberFormat::LowerRoman));
    assert_eq!(pn.start, Some(1));
}

#[test]
fn page_numbering_defaults_write_no_attributes() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    section.set_page_numbering(&mut doc, PageNumbering::default());

    let pn = section.page_numbering(&doc).expect("element exists");
    assert_eq!(pn.format, None);
    assert_eq!(pn.start, None);

    let reopened = reopen(&doc);
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    assert!(
        xml.contains("<w:pgNumType/>"),
        "bare element, no attributes"
    );
}

#[test]
fn page_numbering_clears() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    section.set_page_numbering(
        &mut doc,
        PageNumbering {
            format: Some(NumberFormat::Decimal),
            start: Some(5),
        },
    );
    section.clear_page_numbering(&mut doc);
    assert!(section.page_numbering(&doc).is_none());
}

/// `CT_SectPr` orders `lnNumType` before `pgNumType` and `docGrid` near the end
/// (ECMA-376 §17.6.17). Author them in the wrong order and verify the serialized
/// element order is the schema's, not the call order's.
#[test]
fn sect_pr_children_stay_in_schema_order() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    section.set_doc_grid(
        &mut doc,
        DocGrid {
            grid_type: DocGridType::Lines,
            line_pitch: Some(Pt(23.1)),
            char_space: None,
        },
    );
    section.set_page_numbering(
        &mut doc,
        PageNumbering {
            format: None,
            start: Some(1),
        },
    );
    section.set_line_numbering(
        &mut doc,
        LineNumbering {
            count_by: 1,
            start: 1,
            distance: None,
            restart: docxml::LineNumberRestart::NewPage,
        },
    );

    let reopened = reopen(&doc);
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    let ln = xml.find("<w:lnNumType").expect("lnNumType present");
    let pg = xml.find("<w:pgNumType").expect("pgNumType present");
    let grid = xml.find("<w:docGrid").expect("docGrid present");
    assert!(ln < pg, "lnNumType before pgNumType");
    assert!(pg < grid, "pgNumType before docGrid");
}
