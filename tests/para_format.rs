//! Milestone 8: paragraph formatting (spacing, indent, line spacing, tab stops), breaks,
//! and field codes.
//!
//! Every test drives the public API and, where the exact serialization is load-bearing
//! (shared `w:spacing`, the hanging-indent convention, the field run sequence), walks the
//! saved `word/document.xml` tree to assert the emitted XML. The final test re-establishes
//! the fidelity contract: an untouched open→save is byte-identical across all fixtures.

use std::collections::BTreeMap;
use std::path::Path;

use docxml::opc::Package;
use docxml::xml::{NodeId, XmlTree};
use docxml::{BreakType, Document, Length, LineSpacing, Pt, TabAlignment, TabLeader};

const BASIC: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");
const STYLES_TOC: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/styles_toc.docx"
);
const TABLES_MERGED: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/tables_merged.docx"
);
const HYPERLINKS_IMAGES: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/hyperlinks_images.docx"
);

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

/// Local name of a node (`w:spacing` → `spacing`).
fn local(tree: &XmlTree, id: NodeId) -> String {
    tree.name(id)
        .unwrap()
        .rsplit(':')
        .next()
        .unwrap()
        .to_string()
}

// 1. Double-spacing round-trips, carries line=480 lineRule=auto, and survives a later
//    set_space_after on the same (shared) w:spacing element.
#[test]
fn double_spacing_roundtrips_and_survives_shared_spacing_edit() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Body text");
    p.set_line_spacing(&mut doc, LineSpacing::Double);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("double.docx");
    doc.save(&saved).unwrap();

    // Reopen: Double reads back as Double (not Multiple(2.0)).
    let reopened = Document::open(&saved).unwrap();
    assert_eq!(
        reopened.paragraphs()[0].line_spacing(&reopened),
        Some(LineSpacing::Double)
    );

    // The w:spacing element carries line=480 lineRule=auto.
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let sp = find_first_named(&tree, "w:spacing").expect("a w:spacing");
    assert_eq!(tree.attr(sp, "w:line"), Some("480"));
    assert_eq!(tree.attr(sp, "w:lineRule"), Some("auto"));

    // Now set space-after on the SAME paragraph: the shared w:spacing must keep its line.
    let mut doc2 = Document::open(&saved).unwrap();
    let p2 = doc2.paragraphs()[0];
    p2.set_space_after(&mut doc2, Pt(12.0));
    let saved2 = dir.path().join("double-after.docx");
    doc2.save(&saved2).unwrap();

    let reopened2 = Document::open(&saved2).unwrap();
    let p2 = reopened2.paragraphs()[0];
    assert_eq!(
        p2.line_spacing(&reopened2),
        Some(LineSpacing::Double),
        "line spacing must survive a space_after edit"
    );
    assert_eq!(p2.space_after(&reopened2), Some(Pt(12.0)));

    // And exactly one w:spacing element carries both line and after.
    let tree2 = XmlTree::parse(&document_xml(&saved2)).unwrap();
    let sp2 = find_first_named(&tree2, "w:spacing").expect("a w:spacing");
    assert_eq!(tree2.attr(sp2, "w:line"), Some("480"));
    assert_eq!(tree2.attr(sp2, "w:lineRule"), Some("auto"));
    assert_eq!(tree2.attr(sp2, "w:after"), Some("240")); // 12pt * 20
}

// 2. Left 0.5in + first-line -0.25in writes w:hanging="360" and reads back with the
//    negative first-line convention.
#[test]
fn indents_hanging_first_line_convention() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Indented");
    p.set_left_indent(&mut doc, Length::from_inches(0.5));
    p.set_first_line_indent(&mut doc, Length::from_inches(-0.25));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("indents.docx");
    doc.save(&saved).unwrap();

    // w:ind carries w:left=720 and w:hanging=360 (0.25in), no w:firstLine.
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let ind = find_first_named(&tree, "w:ind").expect("a w:ind");
    assert_eq!(tree.attr(ind, "w:left"), Some("720"));
    assert_eq!(tree.attr(ind, "w:hanging"), Some("360"));
    assert_eq!(tree.attr(ind, "w:firstLine"), None);

    let reopened = Document::open(&saved).unwrap();
    let p = reopened.paragraphs()[0];
    assert_eq!(p.left_indent(&reopened), Some(Length::from_inches(0.5)));
    assert_eq!(
        p.first_line_indent(&reopened),
        Some(Length::from_inches(-0.25)),
        "negative first-line indent reads back from w:hanging"
    );
}

// 3. Right-aligned tab at 6.5in with a dot leader round-trips; clear_tab_stops empties it.
#[test]
fn tab_stops_roundtrip_and_clear() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Name\tPage");
    p.add_tab_stop(
        &mut doc,
        Length::from_inches(6.5),
        TabAlignment::Right,
        TabLeader::Dots,
    );

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("tabs.docx");
    doc.save(&saved).unwrap();

    // XML: one w:tab with w:val=right, w:leader=dot, w:pos=9360 (6.5in * 1440).
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let tab = find_first_named(&tree, "w:tab").expect("a w:tab");
    assert_eq!(tree.attr(tab, "w:val"), Some("right"));
    assert_eq!(tree.attr(tab, "w:leader"), Some("dot"));
    assert_eq!(tree.attr(tab, "w:pos"), Some("9360"));

    let reopened = Document::open(&saved).unwrap();
    let stops = reopened.paragraphs()[0].tab_stops(&reopened);
    assert_eq!(
        stops,
        vec![(
            Length::from_inches(6.5),
            TabAlignment::Right,
            TabLeader::Dots
        )]
    );

    // clear_tab_stops empties the tab stops.
    let mut doc2 = Document::open(&saved).unwrap();
    let p2 = doc2.paragraphs()[0];
    p2.clear_tab_stops(&mut doc2);
    assert!(p2.tab_stops(&doc2).is_empty());

    let saved2 = dir.path().join("tabs-cleared.docx");
    doc2.save(&saved2).unwrap();
    let reopened2 = Document::open(&saved2).unwrap();
    assert!(reopened2.paragraphs()[0].tab_stops(&reopened2).is_empty());
    // The now-empty w:tabs element is removed entirely (CT_Tabs requires >=1 w:tab).
    let tree2 = XmlTree::parse(&document_xml(&saved2)).unwrap();
    assert!(find_first_named(&tree2, "w:tabs").is_none());
}

// 4. add_page_break creates a paragraph + run + w:br type=page; Run::add_break(Line) reads
//    back as a newline in text().
#[test]
fn breaks_page_and_line() {
    let mut doc = Document::new();
    let page_para = doc.add_page_break();
    assert_eq!(page_para.runs(&doc).len(), 1);

    let p = doc.add_paragraph("");
    let r = p.add_run(&mut doc, "before");
    r.add_break(&mut doc, BreakType::Line);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("breaks.docx");
    doc.save(&saved).unwrap();

    let reopened = Document::open(&saved).unwrap();
    // Page-break paragraph: its single run's w:br carries type=page.
    let page_para = reopened.paragraphs()[0];
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let br = find_first_named(&tree, "w:br").expect("a w:br");
    assert_eq!(tree.attr(br, "w:type"), Some("page"));
    assert_eq!(page_para.text(&reopened), "\n");

    // Line break reads back as a newline.
    let line_para = reopened.paragraphs()[1];
    assert_eq!(line_para.text(&reopened), "before\n");
    // The line break is a bare w:br (no w:type).
    let brs: Vec<NodeId> = tree
        .descendants(tree.root())
        .filter(|&n| tree.name(n) == Some("w:br"))
        .collect();
    assert_eq!(brs.len(), 2);
    assert_eq!(
        tree.attr(brs[1], "w:type"),
        None,
        "line break is a bare w:br"
    );
}

// 5. add_page_number_field emits begin / instrText / separate / cached / end in order, the
//    instrText carries xml:space=preserve, the sequence survives a reopen, and the
//    styles_toc footer PAGE field still round-trips byte-identically when untouched.
#[test]
fn page_number_field_sequence_and_footer_fidelity() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Page ");
    p.add_page_number_field(&mut doc);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("pagefield.docx");
    doc.save(&saved).unwrap();

    // Walk the runs of the paragraph and collect the field structure.
    let tree = XmlTree::parse(&document_xml(&saved)).unwrap();
    let para = find_first_named(&tree, "w:p").expect("a w:p");
    // Gather fldCharType values and whether an instrText / result w:t is present, in order.
    let mut sequence: Vec<String> = Vec::new();
    let mut instr_preserve = false;
    for &run in tree.children(para) {
        if local(&tree, run) != "r" {
            continue;
        }
        for &child in tree.children(run) {
            match local(&tree, child).as_str() {
                "fldChar" => sequence.push(format!(
                    "fldChar:{}",
                    tree.attr(child, "w:fldCharType").unwrap()
                )),
                "instrText" => {
                    sequence.push(format!("instrText:{}", tree.text_content(child)));
                    if tree.attr(child, "xml:space") == Some("preserve") {
                        instr_preserve = true;
                    }
                }
                "t" => sequence.push(format!("t:{}", tree.text_content(child))),
                _ => {}
            }
        }
    }
    assert_eq!(
        sequence,
        vec![
            "t:Page ".to_string(),
            "fldChar:begin".to_string(),
            "instrText: PAGE ".to_string(),
            "fldChar:separate".to_string(),
            "t:1".to_string(),
            "fldChar:end".to_string(),
        ],
        "field run sequence out of order"
    );
    assert!(instr_preserve, "instrText must carry xml:space=preserve");

    // The sequence survives a reopen.
    let reopened = Document::open(&saved).unwrap();
    let reopened_xml = XmlTree::parse(&document_xml(&saved)).unwrap();
    let fld_types: Vec<String> = reopened_xml
        .descendants(reopened_xml.root())
        .filter(|&n| reopened_xml.name(n) == Some("w:fldChar"))
        .map(|n| reopened_xml.attr(n, "w:fldCharType").unwrap().to_string())
        .collect();
    assert_eq!(fld_types, ["begin", "separate", "end"]);
    // Paragraph text is the literal prefix plus the cached result.
    assert_eq!(reopened.paragraphs()[0].text(&reopened), "Page 1");

    // The styles_toc footer PAGE field round-trips byte-identically when untouched.
    let original = parts_map(Path::new(STYLES_TOC));
    let doc = Document::open(STYLES_TOC).unwrap();
    let untouched = dir.path().join("styles-untouched.docx");
    doc.save(&untouched).unwrap();
    let saved_parts = parts_map(&untouched);
    assert_eq!(
        original.get("word/footer1.xml"),
        saved_parts.get("word/footer1.xml"),
        "the footer PAGE field part must be byte-identical when untouched"
    );
}

// 6. small_caps / all_caps round-trip; keep_with_next / page_break_before round-trip.
#[test]
fn caps_and_pagination_toggles_roundtrip() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("");
    p.set_keep_with_next(&mut doc, true);
    p.set_page_break_before(&mut doc, true);
    p.set_keep_together(&mut doc, true);
    let r = p.add_run(&mut doc, "SMALL");
    r.set_small_caps(&mut doc, true);
    let r2 = p.add_run(&mut doc, "BIG");
    r2.set_all_caps(&mut doc, true);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("caps.docx");
    doc.save(&saved).unwrap();

    let reopened = Document::open(&saved).unwrap();
    let p = reopened.paragraphs()[0];
    assert!(p.keep_with_next(&reopened));
    assert!(p.page_break_before(&reopened));
    assert!(p.keep_together(&reopened));

    let runs = p.runs(&reopened);
    assert!(runs[0].small_caps(&reopened));
    assert!(!runs[0].all_caps(&reopened));
    assert!(runs[1].all_caps(&reopened));
    assert!(!runs[1].small_caps(&reopened));
}

// 7. Fidelity regression: an untouched open→save is byte-identical across every fixture.
#[test]
fn untouched_save_is_byte_identical_across_fixtures() {
    for fixture in [BASIC, STYLES_TOC, TABLES_MERGED, HYPERLINKS_IMAGES] {
        let doc = Document::open(fixture).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let saved = dir.path().join("untouched.docx");
        doc.save(&saved).unwrap();

        let original = Package::open(fixture).unwrap();
        let out = Package::open(&saved).unwrap();
        for (a, b) in original.parts().iter().zip(out.parts()) {
            assert_eq!(a.name, b.name, "part order changed in {fixture}");
            assert_eq!(
                a.data, b.data,
                "part {} changed on untouched save of {fixture}",
                a.name
            );
        }
    }
}
