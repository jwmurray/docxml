//! The fidelity guarantee: open → save must preserve every part byte-for-byte.

use docxml::opc::Package;

const FIXTURE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

#[test]
fn roundtrip_preserves_every_part_exactly() {
    let original = Package::open(FIXTURE).unwrap();

    let dir = tempfile::tempdir().unwrap();
    let saved_path = dir.path().join("saved.docx");
    original.save(&saved_path).unwrap();

    let saved = Package::open(&saved_path).unwrap();

    let original_names: Vec<_> = original.parts().iter().map(|p| p.name.as_str()).collect();
    let saved_names: Vec<_> = saved.parts().iter().map(|p| p.name.as_str()).collect();
    assert_eq!(original_names, saved_names, "part names/order changed");

    for (a, b) in original.parts().iter().zip(saved.parts()) {
        assert_eq!(a.data, b.data, "part {} changed on round trip", a.name);
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
