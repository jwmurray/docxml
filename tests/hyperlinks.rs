//! Milestone 12: hyperlinks (read + write with relationship creation) and bookmarks.
//!
//! Reading a real external hyperlink out of a python-docx-built fixture; creating hyperlinks
//! in the body and in a header (with the relationship landing in the right part's rels);
//! bookmarks and anchor (internal) hyperlinks; distinct rIds and untouched existing rels;
//! and the fidelity guarantee that an untouched open->save is byte-identical.

use std::collections::BTreeMap;
use std::path::Path;

use docxml::opc::Package;
use docxml::{Document, HeaderFooterType};

const HYPERLINKS: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/hyperlinks_images.docx"
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

// 1. Fixture read: the hyperlinks_images.docx fixture has one external hyperlink; the
//    paragraph reports it with an https URL and non-empty text, and paragraph.text()
//    includes the link text.
#[test]
fn fixture_external_hyperlink_reads_and_text_includes_it() {
    let mut doc = Document::open(HYPERLINKS).unwrap();

    // Find a paragraph that carries a hyperlink.
    let paras = doc.paragraphs();
    let mut found = None;
    for p in &paras {
        let links = p.hyperlinks(&mut doc);
        if let Some(link) = links.into_iter().find(|l| l.url.is_some()) {
            found = Some((*p, link));
            break;
        }
    }
    let (para, link) = found.expect("fixture has a paragraph with an external hyperlink");

    let url = link.url.as_deref().expect("external hyperlink has a url");
    assert!(
        url.starts_with("https://"),
        "external hyperlink url should be an https URL, got {url:?}"
    );
    assert!(!link.text.is_empty(), "hyperlink text should be non-empty");
    assert!(link.anchor.is_none(), "an external link has no anchor");

    // paragraph.text() must include the hyperlink's text (the run inside w:hyperlink).
    let text = para.text(&doc);
    assert!(
        text.contains(&link.text),
        "paragraph.text() {text:?} should include hyperlink text {:?}",
        link.text
    );
}

// 2. Create in the body: add a hyperlink, save/reopen, read it back. The document rels gain
//    exactly one hyperlink relationship with TargetMode=External, and the run carries the
//    Hyperlink character style.
#[test]
fn create_body_hyperlink_roundtrips_with_external_rel_and_rstyle() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Visit ");
    let run = p.add_hyperlink(&mut doc, "https://example.com/", "Example");
    assert_eq!(run.text(&doc), "Example");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("body_link.docx");
    doc.save(&saved).unwrap();

    // Reopen and read the hyperlink back.
    let mut reopened = Document::open(&saved).unwrap();
    let para = reopened.paragraphs()[0];
    let links = para.hyperlinks(&mut reopened);
    assert_eq!(links.len(), 1, "one hyperlink expected");
    assert_eq!(links[0].url.as_deref(), Some("https://example.com/"));
    assert_eq!(links[0].text, "Example");
    assert_eq!(para.text(&reopened), "Visit Example");

    // The document rels contain exactly one hyperlink relationship, TargetMode=External.
    let pkg = Package::open(&saved).unwrap();
    let rels = String::from_utf8(
        pkg.part("word/_rels/document.xml.rels")
            .unwrap()
            .data
            .clone(),
    )
    .unwrap();
    assert_eq!(
        rels.matches("relationships/hyperlink").count(),
        1,
        "exactly one hyperlink relationship: {rels}"
    );
    assert!(
        rels.contains("TargetMode=\"External\"") || rels.contains("TargetMode='External'"),
        "hyperlink relationship must be External: {rels}"
    );
    assert!(
        rels.contains("https://example.com/"),
        "relationship target should be the url: {rels}"
    );

    // The run inside the hyperlink carries w:rStyle w:val="Hyperlink".
    let document = String::from_utf8(pkg.part("word/document.xml").unwrap().data.clone()).unwrap();
    assert!(
        document.contains("rStyle") && document.contains("Hyperlink"),
        "the hyperlink run should carry an rStyle Hyperlink reference: {document}"
    );
}

// 3. Create in a header: the relationship resolves through word/_rels/headerN.xml.rels
//    (created if absent), and the document rels do NOT gain the hyperlink relationship.
#[test]
fn create_header_hyperlink_uses_header_rels_not_document_rels() {
    let mut doc = Document::new();
    let section = doc.sections()[0];
    let header = section.create_header(&mut doc, HeaderFooterType::Default);
    let hp = header.add_paragraph(&mut doc, "Site: ");
    hp.add_hyperlink(&mut doc, "https://headerlink.example/", "home");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("header_link.docx");
    doc.save(&saved).unwrap();

    let pkg = Package::open(&saved).unwrap();

    // The header's own rels part exists and carries the hyperlink relationship.
    let header_rels = pkg
        .part("word/_rels/header1.xml.rels")
        .expect("header rels part should have been created");
    let header_rels = String::from_utf8(header_rels.data.clone()).unwrap();
    assert!(
        header_rels.contains("relationships/hyperlink")
            && header_rels.contains("https://headerlink.example/"),
        "header rels should contain the hyperlink relationship: {header_rels}"
    );

    // document.xml.rels must NOT have gained the hyperlink relationship.
    let doc_rels = String::from_utf8(
        pkg.part("word/_rels/document.xml.rels")
            .unwrap()
            .data
            .clone(),
    )
    .unwrap();
    assert!(
        !doc_rels.contains("relationships/hyperlink"),
        "document rels must not contain the header's hyperlink: {doc_rels}"
    );

    // The link reads back through the header part.
    let mut reopened = Document::open(&saved).unwrap();
    let header = reopened.sections()[0]
        .header(&mut reopened)
        .expect("reopened header");
    // create_header seeds the part with an empty paragraph, so the hyperlink is in the
    // paragraph we added, not necessarily the first one.
    let hpara = header
        .paragraphs(&reopened)
        .into_iter()
        .find(|p| p.text(&reopened).starts_with("Site:"))
        .expect("header paragraph carrying the hyperlink");
    let links = hpara.hyperlinks(&mut reopened);
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].url.as_deref(), Some("https://headerlink.example/"));
    assert_eq!(links[0].text, "home");
}

// 4. Bookmark + anchor hyperlink: the anchor link reports anchor Some / url None, and a
//    second bookmark on the same document gets a distinct id.
#[test]
fn bookmark_and_anchor_hyperlink_roundtrip_with_unique_ids() {
    let mut doc = Document::new();
    let target = doc.add_paragraph("Signature block");
    target.add_bookmark(&mut doc, "sig_block");

    let linker = doc.add_paragraph("Jump to the ");
    linker.add_anchor_hyperlink(&mut doc, "sig_block", "see signature");

    // A second bookmark elsewhere must get a distinct id.
    let other = doc.add_paragraph("Second anchor");
    other.add_bookmark(&mut doc, "second");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("bookmarks.docx");
    doc.save(&saved).unwrap();

    // The anchor hyperlink reads back with anchor Some and url None.
    let mut reopened = Document::open(&saved).unwrap();
    let linker = reopened
        .paragraphs()
        .into_iter()
        .find(|p| p.text(&reopened).starts_with("Jump to the"))
        .expect("the paragraph carrying the anchor hyperlink");
    let links = linker.hyperlinks(&mut reopened);
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].anchor.as_deref(), Some("sig_block"));
    assert!(links[0].url.is_none(), "anchor hyperlink has no url");
    assert_eq!(links[0].text, "see signature");

    // No anchor hyperlink created a relationship in the document rels.
    let pkg = Package::open(&saved).unwrap();
    let doc_rels = String::from_utf8(
        pkg.part("word/_rels/document.xml.rels")
            .unwrap()
            .data
            .clone(),
    )
    .unwrap();
    assert!(
        !doc_rels.contains("relationships/hyperlink"),
        "an anchor hyperlink must not create a relationship: {doc_rels}"
    );

    // The two bookmarks have distinct ids.
    let document = String::from_utf8(pkg.part("word/document.xml").unwrap().data.clone()).unwrap();
    let ids = bookmark_ids(&document);
    assert_eq!(ids.len(), 2, "two bookmarkStart elements: {document}");
    assert_ne!(ids[0], ids[1], "bookmark ids must be unique: {ids:?}");
}

/// The `w:id` values of every `w:bookmarkStart` in a document.xml string, in order.
fn bookmark_ids(xml: &str) -> Vec<String> {
    let mut ids = Vec::new();
    let mut rest = xml;
    while let Some(i) = rest.find("<w:bookmarkStart") {
        rest = &rest[i..];
        let end = rest.find('>').unwrap_or(rest.len());
        let tag = &rest[..end];
        if let Some(j) = tag.find("w:id=\"") {
            let after = &tag[j + 6..];
            if let Some(k) = after.find('"') {
                ids.push(after[..k].to_string());
            }
        }
        rest = &rest[end..];
    }
    ids
}

// 5. Two hyperlinks in one document get distinct rIds, and an existing document's rels are
//    left untouched apart from the additions (the fixture's rId9 hyperlink survives, new
//    ids do not collide with it).
#[test]
fn two_hyperlinks_get_distinct_rids_and_existing_rels_untouched() {
    let original_rels = {
        let pkg = Package::open(HYPERLINKS).unwrap();
        String::from_utf8(
            pkg.part("word/_rels/document.xml.rels")
                .unwrap()
                .data
                .clone(),
        )
        .unwrap()
    };

    let mut doc = Document::open(HYPERLINKS).unwrap();
    let p = doc.add_paragraph("Links: ");
    p.add_hyperlink(&mut doc, "https://first.example/", "first");
    p.add_hyperlink(&mut doc, "https://second.example/", "second");

    let dir = tempfile::tempdir().unwrap();
    let saved = dir.path().join("two_links.docx");
    doc.save(&saved).unwrap();

    let mut reopened = Document::open(&saved).unwrap();
    let para = reopened
        .paragraphs()
        .into_iter()
        .find(|p| p.text(&reopened).starts_with("Links:"))
        .expect("the new paragraph");
    let links = para.hyperlinks(&mut reopened);
    assert_eq!(links.len(), 2);
    assert_eq!(links[0].url.as_deref(), Some("https://first.example/"));
    assert_eq!(links[1].url.as_deref(), Some("https://second.example/"));

    // Distinct rIds in the raw document.xml: the fixture's original hyperlink (rId9) plus
    // the two we added — three w:hyperlink elements, all with distinct r:ids (so the new
    // ones neither collide with each other nor with the pre-existing one).
    let pkg = Package::open(&saved).unwrap();
    let document = String::from_utf8(pkg.part("word/document.xml").unwrap().data.clone()).unwrap();
    let rids = hyperlink_rids(&document);
    assert_eq!(
        rids.len(),
        3,
        "original + two new hyperlink r:ids: {document}"
    );
    let unique: std::collections::HashSet<_> = rids.iter().collect();
    assert_eq!(
        unique.len(),
        3,
        "every hyperlink r:id must be distinct: {rids:?}"
    );

    // The fixture's original relationships are all still present verbatim.
    let new_rels = String::from_utf8(
        pkg.part("word/_rels/document.xml.rels")
            .unwrap()
            .data
            .clone(),
    )
    .unwrap();
    for target in [
        "https://github.com/jwmurray/docxml",
        "media/image1.png",
        "styles.xml",
    ] {
        assert!(
            original_rels.contains(target) && new_rels.contains(target),
            "existing relationship target {target} must survive: {new_rels}"
        );
    }
    // Three hyperlink relationships total now: the original + the two added.
    assert_eq!(
        new_rels.matches("relationships/hyperlink").count(),
        3,
        "original hyperlink + two new ones: {new_rels}"
    );
}

/// The `r:id` values of every `<w:hyperlink>` in a document.xml string, in order.
fn hyperlink_rids(xml: &str) -> Vec<String> {
    let mut ids = Vec::new();
    let mut rest = xml;
    while let Some(i) = rest.find("<w:hyperlink") {
        rest = &rest[i..];
        let end = rest.find('>').unwrap_or(rest.len());
        let tag = &rest[..end];
        if let Some(j) = tag.find("r:id=\"") {
            let after = &tag[j + 6..];
            if let Some(k) = after.find('"') {
                ids.push(after[..k].to_string());
            }
        }
        rest = &rest[end..];
    }
    ids
}

// 6. Fidelity regression: an untouched Document open->save is byte-identical for every
//    fixture.
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
