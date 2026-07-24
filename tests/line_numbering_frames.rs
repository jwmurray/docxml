//! Milestone 13: section line numbering (`w:lnNumType`), paragraph frames (`w:framePr`) and
//! borders (`w:pBdr`), and hidden text (`w:vanish`).
//!
//! Every test drives the public API and, where the exact serialization is load-bearing (the
//! `w:lnNumType` attributes, the `w:pBdr` child order), reopens the saved `.docx` and walks
//! its `word/document.xml` tree. The final test re-establishes the fidelity contract: an
//! untouched open→save is byte-identical across all fixtures.

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use docxml::opc::Package;
use docxml::xml::{NodeId, XmlTree};
use docxml::{
    BorderEdge, BorderStyle, Document, FrameAnchor, FrameOptions, FrameWrap, Length,
    LineNumberRestart, LineNumbering, LineSpacing,
};

/// The main document part's bytes from a saved `.docx`.
fn document_xml(path: &Path) -> Vec<u8> {
    let pkg = Package::open(path).unwrap();
    pkg.part("word/document.xml").unwrap().data.clone()
}

/// Every part of a `.docx`, keyed by name, for byte-level comparison.
fn parts_map(path: &Path) -> BTreeMap<String, Vec<u8>> {
    Package::open(path)
        .unwrap()
        .parts()
        .iter()
        .map(|p| (p.name.clone(), p.data.clone()))
        .collect()
}

/// First element in the tree (pre-order) with the given qualified name.
fn find_first_named(tree: &XmlTree, qname: &str) -> Option<NodeId> {
    tree.descendants(tree.root())
        .find(|&n| tree.name(n) == Some(qname))
}

/// Local name of a node (`w:pBdr` → `pBdr`).
fn local(tree: &XmlTree, id: NodeId) -> String {
    tree.name(id)
        .unwrap()
        .rsplit(':')
        .next()
        .unwrap()
        .to_string()
}

/// Every `.docx` fixture in `tests/fixtures/`, sorted for stable output.
fn all_fixtures() -> Vec<PathBuf> {
    let dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut fixtures: Vec<PathBuf> = std::fs::read_dir(&dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().is_some_and(|ext| ext == "docx"))
        .collect();
    fixtures.sort();
    fixtures
}

// 1. Section line numbering: set on a fresh document, assert the emitted w:lnNumType
//    attributes, and round-trip through save/reopen. clear removes the element.
#[test]
fn line_numbering_roundtrips_and_clears() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    section.set_line_numbering(
        &mut doc,
        LineNumbering {
            count_by: 1,
            start: 1,
            distance: Some(Length::from_inches(0.25)),
            restart: LineNumberRestart::NewPage,
        },
    );

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("lnnum.docx");
    doc.save(&saved).unwrap();

    // Emitted attributes: countBy always written; start omitted (default 1); distance in
    // twips (0.25in = 360); restart omitted (default newPage).
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let el = find_first_named(&tree, "w:lnNumType").expect("w:lnNumType emitted");
    assert_eq!(tree.attr(el, "w:countBy"), Some("1"));
    assert_eq!(tree.attr(el, "w:start"), None, "start default 1 is omitted");
    assert_eq!(tree.attr(el, "w:distance"), Some("360"));
    assert_eq!(
        tree.attr(el, "w:restart"),
        None,
        "restart default newPage is omitted"
    );

    // Read-back through the API.
    let reopened = Document::open(&saved).unwrap();
    let ln = reopened.sections()[0]
        .line_numbering(&reopened)
        .expect("line numbering reads back");
    assert_eq!(ln.count_by, 1);
    assert_eq!(ln.start, 1, "absent start defaults to 1");
    assert_eq!(ln.distance.map(|d| d.twips()), Some(360));
    assert_eq!(ln.restart, LineNumberRestart::NewPage);

    // Clear removes the element entirely.
    let mut doc2 = Document::open(&saved).unwrap();
    doc2.sections()[0].clear_line_numbering(&mut doc2);
    let saved2 = dir.path().join("lnnum-cleared.docx");
    doc2.save(&saved2).unwrap();
    let reopened2 = Document::open(&saved2).unwrap();
    assert_eq!(reopened2.sections()[0].line_numbering(&reopened2), None);
    let tree2 = XmlTree::parse(&document_xml(&saved2)).unwrap();
    assert!(find_first_named(&tree2, "w:lnNumType").is_none());
}

// 1b. Non-default line numbering: count_by 2, start 3, restart Continuous writes all three
//     attributes and round-trips.
#[test]
fn line_numbering_non_default_attrs_roundtrip() {
    let mut doc = Document::new();
    doc.sections()[0].set_line_numbering(
        &mut doc,
        LineNumbering {
            count_by: 2,
            start: 3,
            distance: None,
            restart: LineNumberRestart::Continuous,
        },
    );

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("lnnum2.docx");
    doc.save(&saved).unwrap();

    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let el = find_first_named(&tree, "w:lnNumType").unwrap();
    assert_eq!(tree.attr(el, "w:countBy"), Some("2"));
    assert_eq!(tree.attr(el, "w:start"), Some("3"));
    assert_eq!(tree.attr(el, "w:distance"), None);
    assert_eq!(tree.attr(el, "w:restart"), Some("continuous"));

    let reopened = Document::open(&saved).unwrap();
    let ln = reopened.sections()[0].line_numbering(&reopened).unwrap();
    assert_eq!(ln.count_by, 2);
    assert_eq!(ln.start, 3);
    assert_eq!(ln.distance, None);
    assert_eq!(ln.restart, LineNumberRestart::Continuous);
}

// 1c. Pleading-shaped smoke test: a 28-line double-spaced page — line numbering plus the
//     milestone-8 double spacing on the body paragraph — survives round-trip.
#[test]
fn pleading_paper_setup_roundtrips() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    section.set_line_numbering(&mut doc, LineNumbering::default());
    let p = doc.add_paragraph("1. Plaintiff alleges as follows:");
    p.set_line_spacing(&mut doc, LineSpacing::Double);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("pleading.docx");
    doc.save(&saved).unwrap();

    let reopened = Document::open(&saved).unwrap();
    let ln = reopened.sections()[0].line_numbering(&reopened).unwrap();
    assert_eq!(ln, LineNumbering::default());
    // The body paragraph is the last one (the template's blank paragraph precedes it).
    let body = reopened.paragraphs();
    let last = body.last().unwrap();
    assert_eq!(last.line_spacing(&reopened), Some(LineSpacing::Double));
    assert_eq!(last.text(&reopened), "1. Plaintiff alleges as follows:");
}

// 2. suppress_line_numbers: off by default, round-trips when set.
#[test]
fn suppress_line_numbers_roundtrips() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Caption — not numbered");
    assert!(!p.suppress_line_numbers(&doc), "off by default");
    p.set_suppress_line_numbers(&mut doc, true);
    assert!(p.suppress_line_numbers(&doc));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("suppress.docx");
    doc.save(&saved).unwrap();

    let reopened = Document::open(&saved).unwrap();
    let p = *reopened.paragraphs().last().unwrap();
    assert!(p.suppress_line_numbers(&reopened));

    // Turning it off removes the element.
    let mut doc2 = Document::open(&saved).unwrap();
    let p2 = *doc2.paragraphs().last().unwrap();
    p2.set_suppress_line_numbers(&mut doc2, false);
    assert!(!p2.suppress_line_numbers(&doc2));
    let saved2 = dir.path().join("suppress-off.docx");
    doc2.save(&saved2).unwrap();
    let tree = XmlTree::parse(&document_xml(&saved2)).unwrap();
    assert!(find_first_named(&tree, "w:suppressLineNumbers").is_none());
}

// 3. Frame + borders warning box: a framed, bordered paragraph round-trips exactly, and the
//    w:pBdr children are emitted in CT_PBdr schema order (top, left, bottom, right).
#[test]
fn frame_and_borders_warning_box_roundtrips() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("WARNING: this is a boxed notice.");
    p.set_frame(
        &mut doc,
        FrameOptions {
            width: Some(Length::from_inches(4.0)),
            height: Some(Length::from_inches(1.0)),
            h_anchor: FrameAnchor::Margin,
            v_anchor: FrameAnchor::Text,
            wrap: Some(FrameWrap::Around),
            ..FrameOptions::default()
        },
    );
    let edge = BorderEdge {
        style: BorderStyle::Single,
        size: 4,
        space: 4,
        color: None,
    };
    p.set_borders(&mut doc, Some(edge), Some(edge), Some(edge), Some(edge));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("warnbox.docx");
    doc.save(&saved).unwrap();

    // Read the frame and borders back exactly.
    let reopened = Document::open(&saved).unwrap();
    let p = *reopened.paragraphs().last().unwrap();
    let frame = p.frame(&reopened).expect("frame reads back");
    assert_eq!(frame.width.map(|l| l.twips()), Some(5760)); // 4in
    assert_eq!(frame.height.map(|l| l.twips()), Some(1440)); // 1in
    assert_eq!(frame.h_anchor, FrameAnchor::Margin);
    assert_eq!(frame.v_anchor, FrameAnchor::Text);
    assert_eq!(frame.wrap, Some(FrameWrap::Around));
    assert_eq!(frame.x, None);
    assert_eq!(frame.y, None);

    let (top, bottom, left, right) = p.borders(&reopened);
    assert_eq!(top, Some(edge));
    assert_eq!(bottom, Some(edge));
    assert_eq!(left, Some(edge));
    assert_eq!(right, Some(edge));

    // Height set implies w:hRule="atLeast".
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let fp = find_first_named(&tree, "w:framePr").unwrap();
    assert_eq!(tree.attr(fp, "w:hRule"), Some("atLeast"));

    // w:pBdr children in schema order: top, left, bottom, right.
    let pbdr = find_first_named(&tree, "w:pBdr").unwrap();
    let edges: Vec<String> = tree
        .children(pbdr)
        .iter()
        .copied()
        .map(|c| local(&tree, c))
        .collect();
    assert_eq!(edges, vec!["top", "left", "bottom", "right"]);
}

// 3b. A colored, partial border set: only top and bottom, with an explicit color. The
//     unspecified edges are absent and read back as None.
#[test]
fn partial_colored_borders_roundtrip() {
    use docxml::RgbColor;
    let mut doc = Document::new();
    let p = doc.add_paragraph("Ruled above and below");
    let edge = BorderEdge {
        style: BorderStyle::Double,
        size: 8,
        space: 2,
        color: Some(RgbColor(0x1F, 0x4E, 0x79)),
    };
    p.set_borders(&mut doc, Some(edge), Some(edge), None, None);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("ruled.docx");
    doc.save(&saved).unwrap();

    let reopened = Document::open(&saved).unwrap();
    let p = *reopened.paragraphs().last().unwrap();
    let (top, bottom, left, right) = p.borders(&reopened);
    assert_eq!(top, Some(edge));
    assert_eq!(bottom, Some(edge));
    assert_eq!(left, None);
    assert_eq!(right, None);

    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let pbdr = find_first_named(&tree, "w:pBdr").unwrap();
    let edges: Vec<String> = tree
        .children(pbdr)
        .iter()
        .copied()
        .map(|c| local(&tree, c))
        .collect();
    assert_eq!(edges, vec!["top", "bottom"]);
}

// 4. clear_frame removes w:framePr; set_borders(None x4) removes w:pBdr.
#[test]
fn clear_frame_and_clear_borders_remove_elements() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("boxed");
    p.set_frame(
        &mut doc,
        FrameOptions {
            width: Some(Length::from_inches(3.0)),
            ..FrameOptions::default()
        },
    );
    let edge = BorderEdge {
        style: BorderStyle::Single,
        size: 4,
        space: 4,
        color: None,
    };
    p.set_borders(&mut doc, Some(edge), Some(edge), Some(edge), Some(edge));

    // Now clear both.
    p.clear_frame(&mut doc);
    p.set_borders(&mut doc, None, None, None, None);
    assert!(p.frame(&doc).is_none());
    assert_eq!(p.borders(&doc), (None, None, None, None));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("cleared.docx");
    doc.save(&saved).unwrap();
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    assert!(find_first_named(&tree, "w:framePr").is_none());
    assert!(find_first_named(&tree, "w:pBdr").is_none());
}

// 5. vanish (hidden text) round-trips; the hidden run's text is still included in text().
#[test]
fn vanish_roundtrips_and_text_is_still_present() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("");
    let visible = p.add_run(&mut doc, "Visible ");
    let hidden = p.add_run(&mut doc, "TA-entry");
    assert!(!hidden.vanish(&doc), "off by default");
    hidden.set_vanish(&mut doc, true);
    assert!(hidden.vanish(&doc));
    assert!(!visible.vanish(&doc));

    // Hidden, not absent: full text still reads both runs.
    assert_eq!(p.text(&doc), "Visible TA-entry");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("vanish.docx");
    doc.save(&saved).unwrap();

    let reopened = Document::open(&saved).unwrap();
    let p = *reopened.paragraphs().last().unwrap();
    let runs = p.runs(&reopened);
    assert!(!runs[0].vanish(&reopened));
    assert!(runs[1].vanish(&reopened));
    assert_eq!(p.text(&reopened), "Visible TA-entry");

    // Turning it off removes the w:vanish element.
    let mut doc2 = Document::open(&saved).unwrap();
    let p2 = *doc2.paragraphs().last().unwrap();
    p2.runs(&doc2)[1].set_vanish(&mut doc2, false);
    let saved2 = dir.path().join("vanish-off.docx");
    doc2.save(&saved2).unwrap();
    let reopened2 = Document::open(&saved2).unwrap();
    assert!(!reopened2.paragraphs().last().unwrap().runs(&reopened2)[1].vanish(&reopened2));
}

// 6. Fidelity regression: an untouched open→save is byte-identical across all fixtures.
#[test]
fn untouched_open_save_is_byte_identical() {
    for fixture in all_fixtures() {
        let before = parts_map(&fixture);
        let doc = Document::open(&fixture).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let saved = dir.path().join("saved.docx");
        doc.save(&saved).unwrap();
        let after = parts_map(&saved);
        assert_eq!(
            before,
            after,
            "untouched open→save changed a part for {}",
            fixture.display()
        );
    }
}
