//! The fidelity guarantee: open → save must preserve every part byte-for-byte.

use std::path::{Path, PathBuf};

use docxml::opc::Package;

const FIXTURE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

/// Every `.docx` fixture in `tests/fixtures/`, sorted for stable test output.
fn all_fixtures() -> Vec<PathBuf> {
    let dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut fixtures: Vec<PathBuf> = std::fs::read_dir(&dir)
        .unwrap_or_else(|e| panic!("reading {}: {e}", dir.display()))
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .filter(|p| p.extension().is_some_and(|ext| ext == "docx"))
        .collect();
    fixtures.sort();
    assert!(
        fixtures.len() >= 4,
        "expected at least 4 .docx fixtures in {}, found {}",
        dir.display(),
        fixtures.len()
    );
    fixtures
}

#[test]
fn roundtrip_preserves_every_part_exactly() {
    for fixture in all_fixtures() {
        let original = Package::open(&fixture)
            .unwrap_or_else(|e| panic!("opening {}: {e}", fixture.display()));

        let dir = tempfile::tempdir().unwrap();
        let saved_path = dir.path().join("saved.docx");
        original.save(&saved_path).unwrap();

        let saved = Package::open(&saved_path).unwrap();

        let original_names: Vec<_> = original.parts().iter().map(|p| p.name.as_str()).collect();
        let saved_names: Vec<_> = saved.parts().iter().map(|p| p.name.as_str()).collect();
        assert_eq!(
            original_names,
            saved_names,
            "part names/order changed for {}",
            fixture.display()
        );

        for (a, b) in original.parts().iter().zip(saved.parts()) {
            assert_eq!(
                a.data,
                b.data,
                "part {} changed on round trip for {}",
                a.name,
                fixture.display()
            );
        }
    }
}

#[test]
fn finds_main_document_via_relationships() {
    let pkg = Package::open(FIXTURE).unwrap();
    let doc = pkg.main_document_part().unwrap();
    assert_eq!(doc.name, "word/document.xml");
    let xml = std::str::from_utf8(&doc.data).unwrap();
    assert!(xml.contains("<w:document"), "not a WordprocessingML body");
    assert!(xml.contains("docxml round-trip fixture"));
}

#[test]
fn every_fixture_has_a_locatable_main_document() {
    for fixture in all_fixtures() {
        let pkg = Package::open(&fixture)
            .unwrap_or_else(|e| panic!("opening {}: {e}", fixture.display()));
        let doc = pkg
            .main_document_part()
            .unwrap_or_else(|e| panic!("main document for {}: {e}", fixture.display()));
        let xml = std::str::from_utf8(&doc.data).unwrap();
        assert!(
            xml.contains("<w:document"),
            "{} main document is not a WordprocessingML body",
            fixture.display()
        );
    }
}

#[test]
fn package_relationships_parse() {
    let pkg = Package::open(FIXTURE).unwrap();
    let rels = pkg.relationships().unwrap();
    assert!(!rels.is_empty());
    assert!(rels.iter().all(|r| r.id.starts_with("rId")));
}

#[test]
fn part_lookup_ignores_leading_slash() {
    let pkg = Package::open(FIXTURE).unwrap();
    assert!(pkg.part("/word/document.xml").is_some());
    assert!(pkg.part("word/document.xml").is_some());
    assert!(pkg.part("word/nonexistent.xml").is_none());
}

#[test]
fn open_rejects_non_zip() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("not-a-docx.docx");
    std::fs::write(&path, b"this is not a zip archive").unwrap();
    assert!(Package::open(&path).is_err());
}
