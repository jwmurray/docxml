//! Milestone 10: header/footer part creation and first/even-page headers.
//!
//! Creating a header/footer from scratch (part + content-type Override + relationship +
//! `w:sectPr` reference), the idempotence of a second create, first-page headers via
//! `w:titlePg`, even-page footers via `w:evenAndOddHeaders` (with surgical settings-only
//! dirtying), removing a reference while leaving the part orphaned, and the fidelity
//! guarantee that an untouched `Document` open→save is byte-identical.

use std::collections::BTreeMap;
use std::path::Path;

use docxml::opc::Package;
use docxml::{Document, HeaderFooterType};

const BASIC: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

/// Every part of a `.docx`, keyed by name, for byte-level comparison.
fn parts_map(path: &Path) -> BTreeMap<String, Vec<u8>> {
    Package::open(path)
        .unwrap()
        .parts()
        .iter()
        .map(|p| (p.name.clone(), p.data.clone()))
        .collect()
}

/// The set of part names that differ (by bytes) between two package snapshots that share a
/// key set, ignoring keys present in only one side (new parts).
fn changed_common(
    before: &BTreeMap<String, Vec<u8>>,
    after: &BTreeMap<String, Vec<u8>>,
) -> Vec<String> {
    before
        .iter()
        .filter(|(name, data)| after.get(*name).map(|d| d != *data).unwrap_or(false))
        .map(|(name, _)| name.clone())
        .collect()
}

// 1. A fresh document has no header; create_header(Default) builds the whole chain, the
//    header text round-trips through a save/reopen.
#[test]
fn create_default_header_builds_part_content_type_rel_and_reference() {
    let mut doc = Document::new();
    assert!(
        doc.sections()[0].header(&mut doc).is_none(),
        "a blank document must not already have a header"
    );

    let section = doc.sections()[0];
    let header = section.create_header(&mut doc, HeaderFooterType::Default);
    header.add_paragraph(&mut doc, "Created header");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("created.docx");
    doc.save(&saved).unwrap();

    let pkg = Package::open(&saved).unwrap();

    // The header part exists (headerN.xml — the first one, header1.xml).
    assert!(
        pkg.part("word/header1.xml").is_some(),
        "the new header part word/header1.xml should exist"
    );

    // Content-type Override registers the header content type.
    let content_types =
        String::from_utf8(pkg.part("[Content_Types].xml").unwrap().data.clone()).unwrap();
    assert!(
        content_types.contains("/word/header1.xml")
            && content_types.contains("wordprocessingml.header+xml"),
        "content types should register the header part: {content_types}"
    );

    // The document relationships contain a header relationship pointing at the part.
    let rels = String::from_utf8(
        pkg.part("word/_rels/document.xml.rels")
            .unwrap()
            .data
            .clone(),
    )
    .unwrap();
    assert!(
        rels.contains("relationships/header") && rels.contains("header1.xml"),
        "document rels should contain the header relationship: {rels}"
    );

    // The sectPr carries a headerReference of type=default with a valid r:id.
    let document = String::from_utf8(pkg.part("word/document.xml").unwrap().data.clone()).unwrap();
    assert!(
        document.contains("headerReference"),
        "sectPr should contain a headerReference: {document}"
    );

    // Reopen and confirm the header text.
    let mut reopened = Document::open(&saved).unwrap();
    let header = reopened.sections()[0]
        .header(&mut reopened)
        .expect("reopened default header");
    let text: String = header
        .paragraphs(&reopened)
        .iter()
        .map(|p| p.text(&reopened))
        .collect();
    assert_eq!(text, "Created header");
}

// 2. Creating a header twice returns the same part — no duplicate references or parts.
#[test]
fn create_header_twice_is_idempotent() {
    let mut doc = Document::new();
    let section = doc.sections()[0];

    let first = section.create_header(&mut doc, HeaderFooterType::Default);
    let second = section.create_header(&mut doc, HeaderFooterType::Default);
    assert_eq!(first, second, "second create should return the same handle");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("idempotent.docx");
    doc.save(&saved).unwrap();

    let pkg = Package::open(&saved).unwrap();
    // Exactly one header part.
    let header_parts = pkg
        .parts()
        .iter()
        .filter(|p| p.name.starts_with("word/header") && p.name.ends_with(".xml"))
        .count();
    assert_eq!(header_parts, 1, "only one header part should exist");

    // Exactly one headerReference in the document.
    let document = String::from_utf8(pkg.part("word/document.xml").unwrap().data.clone()).unwrap();
    assert_eq!(
        document.matches("headerReference").count(),
        1,
        "only one headerReference should be present"
    );
}

// 3. First-page header: create_header(First) + set_different_first_page(true). The
//    default and first-page headers are distinct parts with distinct content, and titlePg
//    round-trips.
#[test]
fn first_page_header_is_distinct_and_titlepg_roundtrips() {
    let mut doc = Document::open(BASIC).unwrap(); // has a "Fixture header" default header
    let section = doc.sections()[0];

    let first = section.create_header(&mut doc, HeaderFooterType::First);
    first.add_paragraph(&mut doc, "First page header");
    section.set_different_first_page(&mut doc, true);

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("firstpage.docx");
    doc.save(&saved).unwrap();

    let mut reopened = Document::open(&saved).unwrap();
    let section = reopened.sections()[0];

    assert!(
        section.different_first_page(&reopened),
        "titlePg should be set after reopen"
    );

    let default_text: String = section
        .header_of_type(&mut reopened, HeaderFooterType::Default)
        .expect("default header")
        .paragraphs(&reopened)
        .iter()
        .map(|p| p.text(&reopened))
        .collect();
    let first_text: String = section
        .header_of_type(&mut reopened, HeaderFooterType::First)
        .expect("first-page header")
        .paragraphs(&reopened)
        .iter()
        .map(|p| p.text(&reopened))
        .collect();

    assert!(
        default_text.contains("Fixture header"),
        "default: {default_text:?}"
    );
    assert_eq!(first_text, "First page header");
    assert_ne!(default_text, first_text, "distinct content");

    // Two distinct header parts in the package.
    let pkg = Package::open(&saved).unwrap();
    let header_parts = pkg
        .parts()
        .iter()
        .filter(|p| p.name.starts_with("word/header") && p.name.ends_with(".xml"))
        .count();
    assert_eq!(header_parts, 2, "default + first-page header parts");
}

// 4. Even-page footer + evenAndOddHeaders. After building the even footer and saving once,
//    setting the flag alone dirties ONLY word/settings.xml among the (already-created)
//    parts; the flag and even footer both read back.
#[test]
fn even_footer_and_flag_setting_is_settings_only() {
    // Phase A: create the even footer on a fresh document and save.
    let mut doc = Document::new();
    let footer = doc.sections()[0].create_footer(&mut doc, HeaderFooterType::Even);
    footer.add_paragraph(&mut doc, "Even footer");

    let dir = tempfile::tempdir().unwrap();
    let with_footer = dir.path().join("even_footer.docx");
    doc.save(&with_footer).unwrap();

    // Snapshot the package that already contains the even footer part + reference.
    let baseline = parts_map(&with_footer);

    // Phase B: reopen and set ONLY the even/odd flag; this must touch settings.xml alone.
    let mut doc = Document::open(&with_footer).unwrap();
    assert!(
        !doc.even_and_odd_headers(),
        "flag should start off (blank template ships it unset)"
    );
    doc.set_even_and_odd_headers(true);

    let flagged = dir.path().join("even_flag.docx");
    doc.save(&flagged).unwrap();

    let after = parts_map(&flagged);
    assert_eq!(
        changed_common(&baseline, &after),
        vec!["word/settings.xml".to_string()],
        "only word/settings.xml should change when setting the even/odd flag"
    );

    // settings.xml gained the element.
    let settings = String::from_utf8(after["word/settings.xml"].clone()).unwrap();
    assert!(
        settings.contains("evenAndOddHeaders"),
        "settings.xml should contain evenAndOddHeaders: {settings}"
    );

    // Reopen: flag reads true, and the even footer is still there.
    let mut reopened = Document::open(&flagged).unwrap();
    assert!(reopened.even_and_odd_headers(), "flag reads back true");
    let even_footer_text: String = reopened.sections()[0]
        .footer_of_type(&mut reopened, HeaderFooterType::Even)
        .expect("even footer")
        .paragraphs(&reopened)
        .iter()
        .map(|p| p.text(&reopened))
        .collect();
    assert_eq!(even_footer_text, "Even footer");
}

// 5. Removing a header reference drops it from the sectPr but leaves the part orphaned:
//    header() becomes None, the header part bytes remain in the package, and word/document.xml
//    is the only pre-existing part that changed.
#[test]
fn remove_header_reference_orphans_part_and_touches_only_document() {
    let original = parts_map(Path::new(BASIC));
    let original_header = original["word/header1.xml"].clone();

    let mut doc = Document::open(BASIC).unwrap();
    let removed = doc.sections()[0].remove_header_reference(&mut doc, HeaderFooterType::Default);
    assert!(
        removed,
        "basic.docx has a default header reference to remove"
    );

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("removed.docx");
    doc.save(&saved).unwrap();

    // header() is gone.
    let mut reopened = Document::open(&saved).unwrap();
    assert!(
        reopened.sections()[0].header(&mut reopened).is_none(),
        "header reference should be gone after removal"
    );

    // The orphaned header part is still present byte-for-byte (pass-through).
    let after = parts_map(&saved);
    assert_eq!(
        after.get("word/header1.xml"),
        Some(&original_header),
        "the orphaned header part must survive unchanged"
    );

    // Only word/document.xml changed among pre-existing parts.
    assert_eq!(
        changed_common(&original, &after),
        vec!["word/document.xml".to_string()],
        "only word/document.xml should change when removing a reference"
    );
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>(),
        "no parts added or dropped"
    );
}

// 6. Fidelity regression: an untouched Document open→save is byte-identical for every
//    fixture (the milestone's new code must not perturb the pass-through path).
#[test]
fn untouched_document_roundtrip_is_byte_identical() {
    let dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut fixtures: Vec<_> = std::fs::read_dir(&dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().is_some_and(|x| x == "docx"))
        .collect();
    fixtures.sort();
    assert!(fixtures.len() >= 4);

    for fixture in fixtures {
        let original = parts_map(&fixture);

        let doc = Document::open(&fixture).unwrap();
        let tmp = tempfile::tempdir().unwrap();
        let saved = tmp.path().join("saved.docx");
        doc.save(&saved).unwrap();

        let after = parts_map(&saved);
        assert_eq!(
            original.keys().collect::<Vec<_>>(),
            after.keys().collect::<Vec<_>>(),
            "part set changed for {}",
            fixture.display()
        );
        for (name, data) in &original {
            assert_eq!(
                after.get(name),
                Some(data),
                "part {name} changed on untouched round trip for {}",
                fixture.display()
            );
        }
    }
}
