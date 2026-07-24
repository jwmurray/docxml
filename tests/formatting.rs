//! Milestone 4: character and paragraph formatting — reading known formatting from the
//! fixture, a full write/roundtrip of every setter, schema-order enforcement, the
//! untouched-save fidelity regression, and the underline toggle-off case.

use docxml::opc::Package;
use docxml::xml::{NodeId, XmlTree};
use docxml::{Alignment, Document, Pt, RgbColor};

const FIXTURE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

/// The main document part's bytes from a saved `.docx`.
fn document_xml(path: &std::path::Path) -> Vec<u8> {
    let pkg = Package::open(path).unwrap();
    pkg.part("word/document.xml").unwrap().data.clone()
}

// 1. Read known formatting from the python-docx-generated fixture.
#[test]
fn reads_size_color_and_heading_styles_from_fixture() {
    let doc = Document::open(FIXTURE).unwrap();

    // The "colored 14pt" run: w:sz="28" (→ 14pt), w:color="1F4E79".
    let run = doc
        .paragraphs()
        .iter()
        .flat_map(|p| p.runs(&doc))
        .find(|r| r.text(&doc) == "colored 14pt")
        .expect("a run whose text is \"colored 14pt\"");
    assert_eq!(run.size(&doc), Some(Pt(14.0)));
    assert_eq!(run.color(&doc), Some(RgbColor(0x1F, 0x4E, 0x79)));

    // Heading paragraphs report their styleId (not display name).
    let by_text = |want: &str| {
        doc.paragraphs()
            .into_iter()
            .find(|p| p.text(&doc) == want)
            .unwrap_or_else(|| panic!("paragraph {want:?}"))
    };
    assert_eq!(
        by_text("docxml round-trip fixture")
            .style_id(&doc)
            .as_deref(),
        Some("Title")
    );
    assert_eq!(
        by_text("Formatting").style_id(&doc).as_deref(),
        Some("Heading1")
    );

    // A plain (Normal) body paragraph carries no explicit w:pStyle.
    let formatting_para = "Plain text, then bold, italic, colored 14pt, and unicode: naïve façade — “quotes” • ±5µm ✓";
    assert_eq!(by_text(formatting_para).style_id(&doc), None);
}

// 2. Write every setter, save, reopen, and confirm every getter round-trips.
#[test]
fn write_every_setter_roundtrips() {
    let mut doc = Document::new();

    doc.add_heading("The Title", 0);
    doc.add_heading("Chapter", 1);
    doc.add_heading("Section", 2);

    let p = doc.add_paragraph("");
    p.set_alignment(&mut doc, Alignment::Center);
    p.add_run(&mut doc, "styled")
        .bold(&mut doc, true)
        .italic(&mut doc, true)
        .underline(&mut doc, true)
        .set_size(&mut doc, Pt(18.0))
        .set_color(&mut doc, RgbColor(0xCC, 0x00, 0x33))
        .set_font(&mut doc, "Georgia");

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("formatted.docx");
    doc.save(&saved_path).unwrap();

    let reopened = Document::open(&saved_path).unwrap();
    let paras = reopened.paragraphs();

    assert_eq!(paras[0].style_id(&reopened).as_deref(), Some("Title"));
    assert_eq!(paras[1].style_id(&reopened).as_deref(), Some("Heading1"));
    assert_eq!(paras[2].style_id(&reopened).as_deref(), Some("Heading2"));

    let styled_para = paras[3];
    assert_eq!(styled_para.alignment(&reopened), Some(Alignment::Center));

    let r = styled_para.runs(&reopened)[0];
    assert!(r.is_bold(&reopened));
    assert!(r.is_italic(&reopened));
    assert!(r.is_underlined(&reopened));
    assert_eq!(r.size(&reopened), Some(Pt(18.0)));
    assert_eq!(r.color(&reopened), Some(RgbColor(0xCC, 0x00, 0x33)));
    assert_eq!(r.font(&reopened).as_deref(), Some("Georgia"));
}

// 3. Properties set in a scrambled order still serialize in canonical CT_RPr order.
#[test]
fn rpr_children_are_written_in_schema_order() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("");
    let r = p.add_run(&mut doc, "x");

    // Deliberately scrambled call order: color, font, size, underline, bold.
    r.set_color(&mut doc, RgbColor(0x11, 0x22, 0x33));
    r.set_font(&mut doc, "Arial");
    r.set_size(&mut doc, Pt(12.0));
    r.underline(&mut doc, true);
    r.bold(&mut doc, true);

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("ordered.docx");
    doc.save(&saved_path).unwrap();

    let tree = XmlTree::parse(&document_xml(&saved_path)).unwrap();
    let rpr = find_first_named(&tree, "w:rPr").expect("a w:rPr");
    let locals: Vec<String> = tree
        .children(rpr)
        .iter()
        .filter_map(|&c| tree.name(c))
        .map(|n| n.rsplit(':').next().unwrap().to_string())
        .collect();

    // Canonical CT_RPr subsequence for the properties set: rFonts, b, color, sz, szCs, u.
    assert_eq!(
        locals,
        ["rFonts", "b", "color", "sz", "szCs", "u"],
        "w:rPr children out of schema order"
    );
}

// The pPr order list is likewise enforced: pStyle before jc.
#[test]
fn ppr_children_are_written_in_schema_order() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("x");

    // Set jc first, then pStyle: pStyle must still land first in the serialized pPr.
    p.set_alignment(&mut doc, Alignment::Right);
    p.set_style_id(&mut doc, "Heading3");

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("ppr-ordered.docx");
    doc.save(&saved_path).unwrap();

    let tree = XmlTree::parse(&document_xml(&saved_path)).unwrap();
    let ppr = find_first_named(&tree, "w:pPr").expect("a w:pPr");
    let locals: Vec<String> = tree
        .children(ppr)
        .iter()
        .filter_map(|&c| tree.name(c))
        .map(|n| n.rsplit(':').next().unwrap().to_string())
        .collect();
    assert_eq!(
        locals,
        ["pStyle", "jc"],
        "w:pPr children out of schema order"
    );
}

// 4. Fidelity regression: an untouched open→save keeps every part byte-identical.
#[test]
fn untouched_save_is_byte_identical() {
    let doc = Document::open(FIXTURE).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("untouched.docx");
    doc.save(&saved_path).unwrap();

    let original = Package::open(FIXTURE).unwrap();
    let saved = Package::open(&saved_path).unwrap();
    for (a, b) in original.parts().iter().zip(saved.parts()) {
        assert_eq!(a.name, b.name);
        assert_eq!(a.data, b.data, "part {} changed on untouched save", a.name);
    }
}

// 5. underline(true) then underline(false) leaves no w:u element.
#[test]
fn underline_toggle_off_removes_the_element() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("");
    let r = p.add_run(&mut doc, "x");

    r.underline(&mut doc, true);
    assert!(r.is_underlined(&doc));
    r.underline(&mut doc, false);
    assert!(!r.is_underlined(&doc));

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("no-underline.docx");
    doc.save(&saved_path).unwrap();

    let doc_xml = String::from_utf8(document_xml(&saved_path)).unwrap();
    assert!(
        !doc_xml.contains("<w:u "),
        "w:u element should be gone: {doc_xml}"
    );
    assert!(!doc_xml.contains("<w:u/>"), "bare w:u should be gone");

    let reopened = Document::open(&saved_path).unwrap();
    assert!(!reopened.paragraphs()[0].runs(&reopened)[0].is_underlined(&reopened));
}

/// First element in the tree (pre-order) with the given qualified name.
fn find_first_named(tree: &XmlTree, qname: &str) -> Option<NodeId> {
    tree.descendants(tree.root())
        .find(|&n| tree.name(n) == Some(qname))
}
