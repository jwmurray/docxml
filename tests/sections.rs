//! Milestone 6: sections and headers/footers.
//!
//! Reading section page geometry (size and margins as `Length`), reading header/footer
//! text through the ordinary `Paragraph`/`Run` API, editing a header and proving per-part
//! dirty tracking (only the header part changes; `word/document.xml` and every other part
//! stay byte-identical), the read-only-access-does-not-dirty guarantee, round-tripping
//! page geometry set on a fresh document, and reading a field-bearing footer.

use std::collections::BTreeMap;
use std::path::Path;

use docxml::opc::Package;
use docxml::{Document, Length};

const BASIC: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");
const STYLES_TOC: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/styles_toc.docx"
);

/// Every part of a `.docx`, keyed by name, for byte-level comparison.
fn parts_map(path: &Path) -> BTreeMap<String, Vec<u8>> {
    Package::open(path)
        .unwrap()
        .parts()
        .iter()
        .map(|p| (p.name.clone(), p.data.clone()))
        .collect()
}

// 1. basic.docx: one section, Letter page geometry, 1.25in left margin.
#[test]
fn reads_single_section_and_page_geometry() {
    let doc = Document::open(BASIC).unwrap();

    assert_eq!(doc.sections().len(), 1, "one section expected");
    let section = doc.sections()[0];

    // Left margin is 1.25in = 1800 twips; compare in twips to avoid float wobble.
    assert_eq!(
        section.left_margin(&doc).unwrap().twips(),
        Length::from_inches(1.25).twips(),
        "left margin should be 1.25 inches"
    );

    // Letter: 12240 x 15840 twips.
    assert_eq!(section.page_width(&doc).unwrap().twips(), 12240);
    assert_eq!(section.page_height(&doc).unwrap().twips(), 15840);
}

// 2. basic.docx: header and footer text read through the ordinary Paragraph API.
#[test]
fn reads_header_and_footer_text() {
    let mut doc = Document::open(BASIC).unwrap();
    let section = doc.sections()[0];

    let header = section
        .header(&mut doc)
        .expect("section has a default header");
    let header_text: String = header
        .paragraphs(&doc)
        .iter()
        .map(|p| p.text(&doc))
        .collect();
    assert!(
        header_text.contains("Fixture header"),
        "header text was {header_text:?}"
    );

    let footer = section
        .footer(&mut doc)
        .expect("section has a default footer");
    let footer_text: String = footer
        .paragraphs(&doc)
        .iter()
        .map(|p| p.text(&doc))
        .collect();
    assert!(
        footer_text.contains("Fixture footer"),
        "footer text was {footer_text:?}"
    );
}

// 3. Editing a header dirties ONLY the header part: reopened header carries the new text,
//    word/document.xml (and every other part) stays byte-identical to the original.
#[test]
fn editing_header_only_dirties_the_header_part() {
    let original = parts_map(Path::new(BASIC));

    let mut doc = Document::open(BASIC).unwrap();
    let section = doc.sections()[0];
    let header = section.header(&mut doc).unwrap();
    let para = header.paragraphs(&doc)[0];
    let run = para.runs(&doc)[0];
    run.set_text(&mut doc, "Edited fixture header");

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("edited.docx");
    doc.save(&saved_path).unwrap();

    // The reopened header shows the new text.
    let mut reopened = Document::open(&saved_path).unwrap();
    let section = reopened.sections()[0];
    let header = section.header(&mut reopened).unwrap();
    let header_text: String = header
        .paragraphs(&reopened)
        .iter()
        .map(|p| p.text(&reopened))
        .collect();
    assert!(
        header_text.contains("Edited fixture header"),
        "reopened header text was {header_text:?}"
    );

    // Exactly one part changed, and it is the header part.
    let saved = parts_map(&saved_path);
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        saved.keys().collect::<Vec<_>>(),
        "part set changed"
    );
    let changed: Vec<&String> = original
        .iter()
        .filter(|(name, data)| saved.get(*name) != Some(*data))
        .map(|(name, _)| name)
        .collect();
    assert_eq!(
        changed,
        vec!["word/header1.xml"],
        "only the header part should change; changed = {changed:?}"
    );

    // Spell out the load-bearing guarantee: the main document part is untouched.
    assert_eq!(
        original["word/document.xml"], saved["word/document.xml"],
        "word/document.xml must be byte-identical after a header-only edit"
    );
}

// 4. Read-only header access must not dirty anything: save is byte-identical everywhere.
#[test]
fn reading_header_does_not_dirty_any_part() {
    let original = parts_map(Path::new(BASIC));

    let mut doc = Document::open(BASIC).unwrap();
    let section = doc.sections()[0];
    let header = section.header(&mut doc).unwrap(); // lazily parses word/header1.xml
    let _ = header
        .paragraphs(&doc)
        .iter()
        .map(|p| p.text(&doc))
        .collect::<String>();

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("untouched.docx");
    doc.save(&saved_path).unwrap();

    let saved = parts_map(&saved_path);
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        saved.keys().collect::<Vec<_>>(),
        "part set changed"
    );
    for (name, data) in &original {
        assert_eq!(
            saved.get(name),
            Some(data),
            "part {name} changed despite read-only header access"
        );
    }
}

// 5. Page geometry set on a fresh document round-trips to exact twips.
#[test]
fn sets_page_geometry_and_reads_it_back() {
    let mut doc = Document::new();
    {
        let section = doc.sections()[0];
        section.set_page_width(&mut doc, Length::from_twips(15840));
        section.set_page_height(&mut doc, Length::from_twips(12240));
        section.set_left_margin(&mut doc, Length::from_inches(1.0));
        section.set_right_margin(&mut doc, Length::from_twips(1000));
        section.set_top_margin(&mut doc, Length::from_twips(1234));
        section.set_bottom_margin(&mut doc, Length::from_cm(2.54));
    }

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("geometry.docx");
    doc.save(&saved_path).unwrap();

    let reopened = Document::open(&saved_path).unwrap();
    let section = reopened.sections()[0];
    assert_eq!(section.page_width(&reopened).unwrap().twips(), 15840);
    assert_eq!(section.page_height(&reopened).unwrap().twips(), 12240);
    assert_eq!(section.left_margin(&reopened).unwrap().twips(), 1440);
    assert_eq!(section.right_margin(&reopened).unwrap().twips(), 1000);
    assert_eq!(section.top_margin(&reopened).unwrap().twips(), 1234);
    // 2.54 cm == 1 inch == 1440 twips.
    assert_eq!(section.bottom_margin(&reopened).unwrap().twips(), 1440);
}

// 6. A footer containing a PAGE field reads without panicking.
#[test]
fn reads_field_bearing_footer_without_panicking() {
    let mut doc = Document::open(STYLES_TOC).unwrap();
    let section = doc.sections()[0];
    let footer = section
        .footer(&mut doc)
        .expect("styles_toc footer reference");
    let text: String = footer
        .paragraphs(&doc)
        .iter()
        .map(|p| p.text(&doc))
        .collect();
    // The literal "Page " run precedes the PAGE field; the field result run holds "1".
    assert!(text.contains("Page"), "footer text was {text:?}");
}
