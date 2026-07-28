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

/// `w:snapToGrid` is the paragraph-side half of the document grid: it says whether
/// *this* paragraph's lines occupy whole grid cells. All three authored states
/// round-trip, and they are three states, not two — an absent element inherits,
/// a bare element states "snap", and `w:val="0"` states "do not".
#[test]
fn snap_to_grid_round_trips_all_three_states() {
    let mut doc = Document::new();
    let inherit = doc.add_paragraph("inherits");
    let bound = doc.add_paragraph("grid-bound");
    let free = doc.add_paragraph("free");

    assert_eq!(inherit.snap_to_grid(&doc), None, "absent by default");
    bound.set_snap_to_grid(&mut doc, Some(true));
    free.set_snap_to_grid(&mut doc, Some(false));

    let reopened = reopen(&doc);
    let paras = reopened.paragraphs();
    let by_text = |t: &str| {
        *paras
            .iter()
            .find(|p| p.text(&reopened) == t)
            .unwrap_or_else(|| panic!("no paragraph {t}"))
    };
    assert_eq!(by_text("inherits").snap_to_grid(&reopened), None);
    assert_eq!(by_text("grid-bound").snap_to_grid(&reopened), Some(true));
    assert_eq!(by_text("free").snap_to_grid(&reopened), Some(false));

    // The serialization is the one Word writes: bare element for on, an explicit
    // zero for off. A bare element for "off" would read back as "on" everywhere.
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    assert!(
        xml.contains("<w:snapToGrid/>"),
        "on is a bare element: {xml}"
    );
    assert!(
        xml.contains(r#"<w:snapToGrid w:val="0"/>"#),
        "off is an explicit zero: {xml}"
    );

    // Clearing removes the element rather than writing the default.
    let target = by_text("free");
    let mut doc2 = reopened;
    target.set_snap_to_grid(&mut doc2, None);
    let reopened2 = reopen(&doc2);
    let cleared = *reopened2
        .paragraphs()
        .iter()
        .find(|p| p.text(&reopened2) == "free")
        .expect("paragraph survives");
    assert_eq!(cleared.snap_to_grid(&reopened2), None);
}

/// `w:snapToGrid` sits between `w:adjustRightInd` and `w:spacing` in `CT_PPr`
/// (ECMA-376 §17.3.1.32). A paragraph that carries both must serialize in that
/// order or Word rejects the part.
#[test]
fn snap_to_grid_precedes_spacing_in_ppr() {
    use docxml::LineSpacing;

    let mut doc = Document::new();
    let p = doc.add_paragraph("both");
    // Deliberately set spacing FIRST: order is the writer's job, not the caller's.
    p.set_line_spacing(&mut doc, LineSpacing::Exactly(Pt(23.1)));
    p.set_snap_to_grid(&mut doc, Some(true));

    let reopened = reopen(&doc);
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    let snap = xml.find("<w:snapToGrid").expect("snapToGrid present");
    let spacing = xml.find("<w:spacing").expect("spacing present");
    assert!(snap < spacing, "snapToGrid must precede spacing: {xml}");
}
