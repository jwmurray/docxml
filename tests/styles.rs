//! Milestone 14: styles catalog, style authoring, docDefaults, and style-aware
//! (effective) formatting reads.
//!
//! Covers reading the style catalog out of `basic.docx` (`styles`, `style_by_id`,
//! `style_by_name`, `Paragraph::style_name`); the effective-formatting reads that resolve
//! through the `w:rStyle` / `w:pStyle` `w:basedOn` chains and `w:docDefaults`; authoring new
//! paragraph and character styles (`create_style`, the `Style` setters, CT_Style child
//! order); `set_default_font` writing `w:docDefaults`; and the fidelity contract —
//! modifying only a style dirties only `styles.xml`, and an untouched open→save is
//! byte-identical across every fixture.

use std::path::{Path, PathBuf};

use docxml::opc::Package;
use docxml::xml::{NodeId, XmlTree};
use docxml::{Alignment, Document, Pt, StyleType};

const BASIC: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

/// Parse a named part of a saved `.docx` into a tree.
fn part_tree(path: &Path, part: &str) -> XmlTree {
    let pkg = Package::open(path).unwrap();
    let bytes = &pkg.part(part).unwrap().data;
    XmlTree::parse(bytes).unwrap()
}

/// Local name of an element (`w:style` → `style`).
fn local(tree: &XmlTree, id: NodeId) -> Option<String> {
    tree.name(id)
        .map(|n| n.rsplit(':').next().unwrap().to_string())
}

/// The `w:style` element with the given `w:styleId` in a styles tree, if present.
fn style_node(tree: &XmlTree, style_id: &str) -> Option<NodeId> {
    let root = tree.root();
    tree.children(root).iter().copied().find(|&c| {
        local(tree, c).as_deref() == Some("style") && tree.attr(c, "w:styleId") == Some(style_id)
    })
}

/// Every `.docx` fixture in `tests/fixtures/`, sorted.
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

/// The first body paragraph whose style id equals `style_id`.
fn paragraph_with_style(doc: &Document, style_id: &str) -> docxml::Paragraph {
    doc.paragraphs()
        .into_iter()
        .find(|p| p.style_id(doc).as_deref() == Some(style_id))
        .expect("a paragraph with that style")
}

// 1. The style catalog read out of the python-docx-generated fixture.
#[test]
fn reads_style_catalog_from_fixture() {
    let mut doc = Document::open(BASIC).unwrap();

    // Non-empty catalog with the built-in styles present.
    assert!(!doc.styles().is_empty());

    // Heading1's display name is python-docx's lowercased builtin name.
    let h1 = doc.style_by_id("Heading1").expect("Heading1 defined");
    assert_eq!(h1.display_name(&doc).as_deref(), Some("heading 1"));
    assert_eq!(h1.style_type(&doc), StyleType::Paragraph);
    assert_eq!(h1.based_on(&doc).as_deref(), Some("Normal"));

    // style_by_name round-trips back to the same styleId (display-name lookup).
    let by_name = doc.style_by_name("heading 1").expect("found by name");
    assert_eq!(by_name.style_id(&doc), "Heading1");

    // Case-insensitive fallback also finds it.
    assert_eq!(
        doc.style_by_name("HEADING 1").map(|s| s.style_id(&doc)),
        Some("Heading1".to_string())
    );

    // Paragraph::style_name resolves styleId -> display name on a heading paragraph.
    let heading = paragraph_with_style(&doc, "Heading1");
    assert_eq!(heading.style_name(&mut doc).as_deref(), Some("heading 1"));
}

// 2. Effective (style-aware) formatting reads against the fixture.
#[test]
fn effective_reads_resolve_through_chain_and_defaults() {
    let mut doc = Document::open(BASIC).unwrap();

    // The Heading1 paragraph's run carries no direct rPr bold, but the Heading1 style sets
    // w:b and w:sz=28 (14pt) — the effective reads must see those through the pStyle chain.
    let heading = paragraph_with_style(&doc, "Heading1");
    let run = heading.runs(&doc)[0];
    assert!(!run.is_bold(&doc), "no direct bold on the run");
    assert!(run.effective_bold(&mut doc), "Heading1 style is bold");
    assert_eq!(run.effective_size(&mut doc), Some(Pt(14.0)));

    // A plain body run (no direct formatting, no paragraph style) falls all the way through
    // to w:docDefaults, whose rPrDefault sets w:sz=22 (11pt).
    let plain = doc
        .paragraphs()
        .into_iter()
        .find(|p| p.style_id(&doc).is_none() && p.text(&doc).starts_with("Plain text"))
        .expect("the plain formatting paragraph");
    let first = plain.runs(&doc)[0];
    assert!(!first.effective_bold(&mut doc));
    assert_eq!(first.effective_size(&mut doc), Some(Pt(11.0)));

    // Direct formatting still wins: the "colored 14pt" run sets w:sz=28 directly.
    let colored = plain
        .runs(&doc)
        .into_iter()
        .find(|r| r.text(&doc).contains("colored"))
        .expect("the colored run");
    assert_eq!(colored.effective_size(&mut doc), Some(Pt(14.0)));
}

// 3. Authoring a new paragraph style, applying it, and round-tripping.
#[test]
fn authors_paragraph_style_and_round_trips() {
    let mut doc = Document::new();
    let style = doc.create_style("FirmTitle", "Firm Title", StyleType::Paragraph);
    style
        .set_based_on(&mut doc, "Normal")
        .set_bold(&mut doc, true)
        .set_size(&mut doc, Pt(14.0))
        .set_alignment(&mut doc, Alignment::Center);

    let p = doc.add_paragraph("Hepworth Legal");
    p.set_style_id(&mut doc, "FirmTitle");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("styled.docx");
    doc.save(&saved).unwrap();

    // Inspect the serialized style: children in CT_Style order.
    let styles = part_tree(&saved, "word/styles.xml");
    let style_el = style_node(&styles, "FirmTitle").expect("FirmTitle serialized");
    assert_eq!(styles.attr(style_el, "w:type"), Some("paragraph"));
    let child_locals: Vec<String> = styles
        .children(style_el)
        .iter()
        .filter_map(|&c| local(&styles, c))
        .collect();
    assert_eq!(child_locals, ["name", "basedOn", "qFormat", "pPr", "rPr"]);

    // Reopen and confirm the paragraph reports its style and inherits formatting.
    let mut reopened = Document::open(&saved).unwrap();
    let para = paragraph_with_style(&reopened, "FirmTitle");
    assert_eq!(para.style_id(&reopened).as_deref(), Some("FirmTitle"));
    assert_eq!(
        para.style_name(&mut reopened).as_deref(),
        Some("Firm Title")
    );
    let run = para.runs(&reopened)[0];
    assert!(run.effective_bold(&mut reopened));
    assert_eq!(run.effective_size(&mut reopened), Some(Pt(14.0)));
}

// 4. A character style resolved through the w:rStyle chain.
#[test]
fn character_style_resolves_via_rstyle_chain() {
    let mut doc = Document::new();
    doc.create_style("Emphatic", "Emphatic", StyleType::Character)
        .set_bold(&mut doc, true)
        .set_size(&mut doc, Pt(20.0));

    let p = doc.add_paragraph("");
    let run = p.add_run(&mut doc, "loud");
    run.set_style_id(&mut doc, "Emphatic");

    assert_eq!(run.style_id(&doc).as_deref(), Some("Emphatic"));
    assert!(!run.is_bold(&doc), "no direct bold");
    assert!(run.effective_bold(&mut doc), "bold via rStyle");
    assert_eq!(run.effective_size(&mut doc), Some(Pt(20.0)));
}

// 5. docDefaults: set_default_font writes the run defaults, first child of w:styles.
#[test]
fn sets_default_font_in_docdefaults() {
    let mut doc = Document::new();
    doc.set_default_font("Century Schoolbook", Pt(13.0));

    // A run with no other formatting inherits the defaults.
    let p = doc.add_paragraph("");
    let run = p.add_run(&mut doc, "body");
    assert_eq!(
        run.effective_font(&mut doc).as_deref(),
        Some("Century Schoolbook")
    );
    assert_eq!(run.effective_size(&mut doc), Some(Pt(13.0)));

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("defaults.docx");
    doc.save(&saved).unwrap();

    let styles = part_tree(&saved, "word/styles.xml");
    let root = styles.root();
    // docDefaults is the FIRST child of w:styles.
    assert_eq!(
        local(&styles, styles.children(root)[0]).as_deref(),
        Some("docDefaults")
    );
    let docdefaults = styles.children(root)[0];
    let rpr_default = styles
        .children(docdefaults)
        .iter()
        .copied()
        .find(|&c| local(&styles, c).as_deref() == Some("rPrDefault"))
        .unwrap();
    let rpr = styles.children(rpr_default)[0];
    let rfonts = styles
        .children(rpr)
        .iter()
        .copied()
        .find(|&c| local(&styles, c).as_deref() == Some("rFonts"))
        .unwrap();
    assert_eq!(styles.attr(rfonts, "w:ascii"), Some("Century Schoolbook"));
    assert_eq!(styles.attr(rfonts, "w:hAnsi"), Some("Century Schoolbook"));
    let sz = styles
        .children(rpr)
        .iter()
        .copied()
        .find(|&c| local(&styles, c).as_deref() == Some("sz"))
        .unwrap();
    assert_eq!(styles.attr(sz, "w:val"), Some("26")); // 13pt in half-points

    // Round-trip: the reopened default still reads back.
    let mut reopened = Document::open(&saved).unwrap();
    let r = reopened.add_paragraph("").add_run(&mut reopened, "x");
    assert_eq!(r.effective_size(&mut reopened), Some(Pt(13.0)));
}

// 6. Dirty tracking: modifying only a style re-serializes only styles.xml.
#[test]
fn modifying_a_style_dirties_only_styles_xml() {
    let original = Package::open(BASIC).unwrap();

    let mut doc = Document::open(BASIC).unwrap();
    // Modify an existing style — no new part, so content-types/rels are untouched.
    doc.style_by_id("Heading1")
        .unwrap()
        .set_italic(&mut doc, true);

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("saved.docx");
    doc.save(&saved_path).unwrap();
    let saved = Package::open(&saved_path).unwrap();

    // Same set of parts, same order.
    let orig_names: Vec<_> = original.parts().iter().map(|p| p.name.as_str()).collect();
    let saved_names: Vec<_> = saved.parts().iter().map(|p| p.name.as_str()).collect();
    assert_eq!(orig_names, saved_names);

    // Every part except styles.xml is byte-identical.
    for (a, b) in original.parts().iter().zip(saved.parts()) {
        if a.name == "word/styles.xml" {
            assert_ne!(a.data, b.data, "styles.xml should have changed");
        } else {
            assert_eq!(a.data, b.data, "{} must be byte-identical", a.name);
        }
    }
}

// 7. create_style is idempotent, and untouched open->save stays byte-identical everywhere.
#[test]
fn create_style_is_idempotent() {
    let mut doc = Document::new();
    let first = doc.create_style("Custom", "Custom", StyleType::Paragraph);
    first.set_bold(&mut doc, true);
    // A second call with the same id returns the existing style, not a duplicate.
    let second = doc.create_style("Custom", "Ignored Name", StyleType::Character);
    assert_eq!(first.node(), second.node());
    assert_eq!(second.style_id(&doc), "Custom");
    // The original display name and type are preserved (idempotent = no overwrite).
    assert_eq!(second.display_name(&doc).as_deref(), Some("Custom"));
    assert_eq!(second.style_type(&doc), StyleType::Paragraph);

    // Exactly one w:style with that id exists.
    let count = doc
        .styles()
        .into_iter()
        .filter(|s| s.style_id(&doc) == "Custom")
        .count();
    assert_eq!(count, 1);
}

#[test]
fn untouched_open_save_is_byte_identical_across_fixtures() {
    for fixture in all_fixtures() {
        let original = Package::open(&fixture).unwrap();
        let doc = Document::open(&fixture).unwrap();

        let dir = tempfile::tempdir().unwrap();
        let saved_path = dir.path().join("saved.docx");
        doc.save(&saved_path).unwrap();
        let saved = Package::open(&saved_path).unwrap();

        for (a, b) in original.parts().iter().zip(saved.parts()) {
            assert_eq!(
                a.data,
                b.data,
                "{} changed on untouched round trip of {}",
                a.name,
                fixture.display()
            );
        }
    }
}
