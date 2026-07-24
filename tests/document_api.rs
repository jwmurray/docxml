//! Typed document API: reading, creating, editing, and the fidelity + modified-tracking
//! guarantees that layer on top of the OPC and XML round-trip contracts.

use docxml::Document;
use docxml::opc::Package;
use docxml::xml::XmlTree;

const FIXTURE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

/// The main document part's bytes from a saved `.docx`.
fn document_xml(path: &std::path::Path) -> Vec<u8> {
    let pkg = Package::open(path).unwrap();
    pkg.part("word/document.xml").unwrap().data.clone()
}

// 1. Fidelity: open then save without touching anything → every part byte-identical.
#[test]
fn untouched_save_preserves_every_part() {
    let doc = Document::open(FIXTURE).unwrap();

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("untouched.docx");
    doc.save(&saved_path).unwrap();

    let original = Package::open(FIXTURE).unwrap();
    let saved = Package::open(&saved_path).unwrap();

    let original_names: Vec<_> = original.parts().iter().map(|p| p.name.as_str()).collect();
    let saved_names: Vec<_> = saved.parts().iter().map(|p| p.name.as_str()).collect();
    assert_eq!(original_names, saved_names, "part names/order changed");

    for (a, b) in original.parts().iter().zip(saved.parts()) {
        assert_eq!(a.data, b.data, "part {} changed on untouched save", a.name);
    }
}

// 2. Read: the fixture's known paragraph texts and the known bold run.
#[test]
fn reads_fixture_paragraphs_and_bold_run() {
    let doc = Document::open(FIXTURE).unwrap();
    let paragraphs = doc.paragraphs();

    // Body-level paragraphs only (table cells excluded); the fixture has a dozen.
    assert!(
        paragraphs.len() >= 10,
        "expected >= 10 body paragraphs, got {}",
        paragraphs.len()
    );

    let texts: Vec<String> = paragraphs.iter().map(|p| p.text(&doc)).collect();
    for expected in [
        "docxml round-trip fixture",
        "Plain text, then bold, italic, colored 14pt, and unicode: naïve façade — “quotes” • ±5µm ✓",
        "first",
        "one",
        "Second page after an explicit break.",
    ] {
        assert!(
            texts.iter().any(|t| t == expected),
            "missing expected paragraph text: {expected:?}\ngot: {texts:?}"
        );
    }

    // The run reading "bold" in the formatting paragraph is bold.
    let bold_run = paragraphs
        .iter()
        .flat_map(|p| p.runs(&doc))
        .find(|r| r.text(&doc) == "bold")
        .expect("a run whose text is \"bold\"");
    assert!(bold_run.is_bold(&doc));
    assert!(!bold_run.is_italic(&doc));

    // The "italic" run is italic but not bold.
    let italic_run = paragraphs
        .iter()
        .flat_map(|p| p.runs(&doc))
        .find(|r| r.text(&doc) == "italic")
        .expect("a run whose text is \"italic\"");
    assert!(italic_run.is_italic(&doc));
    assert!(!italic_run.is_bold(&doc));
}

// 3. Create: new document, add content with formatting, save, reopen, verify round trip.
#[test]
fn create_add_content_roundtrips() {
    let mut doc = Document::new();
    assert!(
        doc.paragraphs().is_empty(),
        "blank template has no body paragraphs"
    );

    doc.add_paragraph("Plain heading");
    let p = doc.add_paragraph("");
    p.add_run(&mut doc, "bold bit").bold(&mut doc, true);
    p.add_run(&mut doc, " and ");
    p.add_run(&mut doc, "italic bit").italic(&mut doc, true);

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("created.docx");
    doc.save(&saved_path).unwrap();

    // The saved file is a valid package whose document part parses as an XmlTree.
    let pkg = Package::open(&saved_path).unwrap();
    let doc_part = pkg.main_document_part().unwrap();
    let tree = XmlTree::parse(&doc_part.data).unwrap();
    assert_eq!(tree.name(tree.root()), Some("w:document"));

    // Reopen through the typed API: texts and formatting flags survived.
    let reopened = Document::open(&saved_path).unwrap();
    let texts: Vec<String> = reopened
        .paragraphs()
        .iter()
        .map(|p| p.text(&reopened))
        .collect();
    assert_eq!(texts, ["Plain heading", "bold bit and italic bit"]);

    let runs = reopened.paragraphs()[1].runs(&reopened);
    assert_eq!(runs.len(), 3);
    assert!(runs[0].is_bold(&reopened));
    assert!(!runs[0].is_italic(&reopened));
    assert!(!runs[1].is_bold(&reopened));
    assert!(!runs[1].is_italic(&reopened));
    assert!(runs[2].is_italic(&reopened));
    assert!(!runs[2].is_bold(&reopened));
}

// 4. Edit: change an existing run's text; only word/document.xml changes; it stays valid.
#[test]
fn edit_existing_run_touches_only_document_xml() {
    let mut doc = Document::open(FIXTURE).unwrap();

    // Retarget the run that currently reads "bold".
    let target = doc
        .paragraphs()
        .iter()
        .flat_map(|p| p.runs(&doc))
        .find(|r| r.text(&doc) == "bold")
        .expect("a run whose text is \"bold\"");
    target.set_text(&mut doc, "STRONG");

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("edited.docx");
    doc.save(&saved_path).unwrap();

    let reopened = Document::open(&saved_path).unwrap();
    let all_text: String = reopened
        .paragraphs()
        .iter()
        .map(|p| p.text(&reopened))
        .collect();
    assert!(all_text.contains("STRONG"), "new text missing");
    assert!(
        !all_text.contains("bold"),
        "old text still present: {all_text}"
    );

    // Formatting on the edited run is intact (set_text leaves w:rPr alone).
    let edited_run = reopened
        .paragraphs()
        .iter()
        .flat_map(|p| p.runs(&reopened))
        .find(|r| r.text(&reopened) == "STRONG")
        .expect("edited run");
    assert!(edited_run.is_bold(&reopened));

    // Every part other than word/document.xml is byte-identical to the original.
    let original = Package::open(FIXTURE).unwrap();
    let saved = Package::open(&saved_path).unwrap();
    for (a, b) in original.parts().iter().zip(saved.parts()) {
        assert_eq!(a.name, b.name);
        if a.name == "word/document.xml" {
            assert_ne!(a.data, b.data, "document.xml should have changed");
        } else {
            assert_eq!(a.data, b.data, "part {} changed unexpectedly", a.name);
        }
    }

    // document.xml is still semantically valid: parses, root is w:document.
    let doc_xml = document_xml(&saved_path);
    let tree = XmlTree::parse(&doc_xml).unwrap();
    assert_eq!(tree.name(tree.root()), Some("w:document"));
}

// Whitespace-bearing runs get xml:space="preserve" and survive a round trip intact.
#[test]
fn whitespace_runs_are_preserved() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("");
    p.add_run(&mut doc, "leading");
    p.add_run(&mut doc, " and trailing ");

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("spaces.docx");
    doc.save(&saved_path).unwrap();

    let doc_xml = String::from_utf8(document_xml(&saved_path)).unwrap();
    assert!(
        doc_xml.contains("xml:space=\"preserve\""),
        "expected xml:space=preserve for the whitespace-edged run"
    );

    let reopened = Document::open(&saved_path).unwrap();
    assert_eq!(
        reopened.paragraphs()[0].text(&reopened),
        "leading and trailing "
    );
}

// Toggling bold/italic off removes the property; turning it back on restores it.
#[test]
fn formatting_toggles_off_and_on() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("");
    let r = p.add_run(&mut doc, "text");

    r.bold(&mut doc, true).italic(&mut doc, true);
    assert!(r.is_bold(&doc) && r.is_italic(&doc));

    r.bold(&mut doc, false);
    assert!(!r.is_bold(&doc));
    assert!(r.is_italic(&doc), "clearing bold must not clear italic");

    // Round-trips through save/reopen.
    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("toggles.docx");
    doc.save(&saved_path).unwrap();
    let reopened = Document::open(&saved_path).unwrap();
    let rr = reopened.paragraphs()[0].runs(&reopened)[0];
    assert!(!rr.is_bold(&reopened));
    assert!(rr.is_italic(&reopened));
}

// 5. Modified-tracking: reads alone must not mark the document modified.
#[test]
fn reads_do_not_modify() {
    let doc = Document::open(FIXTURE).unwrap();

    // Exercise every read accessor.
    for p in doc.paragraphs() {
        let _ = p.text(&doc);
        for r in p.runs(&doc) {
            let _ = r.text(&doc);
            let _ = r.is_bold(&doc);
            let _ = r.is_italic(&doc);
        }
    }

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("after-reads.docx");
    doc.save(&saved_path).unwrap();

    // document.xml is byte-identical: reads never triggered re-serialization.
    let original = document_xml(std::path::Path::new(FIXTURE));
    let saved = document_xml(&saved_path);
    assert_eq!(original, saved, "reads must not re-serialize document.xml");
}
