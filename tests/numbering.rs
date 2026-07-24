//! Milestone 9: numbering / lists.
//!
//! Covers reading a paragraph's numbering (direct `w:numPr` and style-resolved), the
//! direct `set_numbering` / `clear_numbering` writers, the `add_bullet_paragraph` /
//! `add_numbered_paragraph` convenience methods (python-docx parity through the template's
//! List styles), and `create_numbering` — including the from-scratch creation of the
//! numbering part, its content-type `Override`, and its relationship. The final test
//! re-establishes the fidelity contract: an untouched open→save is byte-identical across
//! every fixture.

use std::path::{Path, PathBuf};

use docxml::opc::Package;
use docxml::xml::{NodeId, XmlTree};
use docxml::{Document, NumberFormat};

const BASIC: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/basic.docx");

/// Parse a named part of a saved `.docx` into a tree.
fn part_tree(path: &Path, part: &str) -> XmlTree {
    let pkg = Package::open(path).unwrap();
    let bytes = &pkg.part(part).unwrap().data;
    XmlTree::parse(bytes).unwrap()
}

/// Raw UTF-8 bytes of a named part.
fn part_text(path: &Path, part: &str) -> String {
    let pkg = Package::open(path).unwrap();
    String::from_utf8(pkg.part(part).unwrap().data.clone()).unwrap()
}

/// Local name of an element (`w:num` → `num`).
fn local(tree: &XmlTree, id: NodeId) -> Option<String> {
    tree.name(id)
        .map(|n| n.rsplit(':').next().unwrap().to_string())
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

// 1. Reading numbering out of the basic.docx fixture. Its "List Bullet" and "List Number"
//    paragraphs carry only a pStyle (no direct w:numPr) — the numbering lives on the
//    style — so numbering() must resolve through styles.xml. Both groups report Some, and
//    the bullet group's numId differs from the number group's.
#[test]
fn reads_style_resolved_numbering_from_fixture() {
    let doc = Document::open(BASIC).unwrap();

    let mut bullet_ids = Vec::new();
    let mut number_ids = Vec::new();
    for p in doc.paragraphs() {
        match p.style_id(&doc).as_deref() {
            Some("ListBullet") => {
                let (num_id, ilvl) = p
                    .numbering(&doc)
                    .expect("List Bullet paragraph reports numbering via its style");
                assert_eq!(ilvl, 0, "level defaults to 0");
                bullet_ids.push(num_id);
            }
            Some("ListNumber") => {
                let (num_id, _) = p
                    .numbering(&doc)
                    .expect("List Number paragraph reports numbering via its style");
                number_ids.push(num_id);
            }
            _ => {}
        }
    }

    assert!(!bullet_ids.is_empty(), "fixture has List Bullet paragraphs");
    assert!(!number_ids.is_empty(), "fixture has List Number paragraphs");
    // Every bullet paragraph resolves to the same numId; likewise every number paragraph.
    assert!(bullet_ids.iter().all(|&n| n == bullet_ids[0]));
    assert!(number_ids.iter().all(|&n| n == number_ids[0]));
    // The two groups reference distinct numberings.
    assert_ne!(
        bullet_ids[0], number_ids[0],
        "bullet and number groups use different numIds"
    );
}

// 2. The convenience constructors give python-docx-parity style ids and, on a
//    Document::new() document, functioning numbering out of the box.
//
//    Reading the numbering *through the style* is exercised by numbering() here, but the
//    load-bearing template invariant is asserted directly against styles.xml: the raw part
//    text links styleId="ListBullet"/"ListNumber" to a w:numId inside a w:numPr, which is
//    what makes the styled paragraphs render as lists.
#[test]
fn convenience_paragraphs_carry_list_styles_and_numbering() {
    let mut doc = Document::new();
    doc.add_bullet_paragraph("first bullet");
    doc.add_bullet_paragraph("second bullet");
    doc.add_numbered_paragraph("first number");
    doc.add_numbered_paragraph("second number");

    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("lists.docx");
    doc.save(&path).unwrap();

    let reopened = Document::open(&path).unwrap();
    let styles: Vec<Option<String>> = reopened
        .paragraphs()
        .iter()
        .map(|p| p.style_id(&reopened))
        .collect();
    assert_eq!(
        styles,
        vec![
            Some("ListBullet".to_string()),
            Some("ListBullet".to_string()),
            Some("ListNumber".to_string()),
            Some("ListNumber".to_string()),
        ]
    );

    // Numbering is present (resolved through the style) for every list paragraph.
    for p in reopened.paragraphs() {
        assert!(
            p.numbering(&reopened).is_some(),
            "styled list paragraph reports numbering"
        );
    }

    // The template's List styles reference numbering: styles.xml carries a w:numPr with a
    // w:numId inside each of the ListBullet / ListNumber style definitions. This is the
    // link the convenience methods rely on (the style, not the paragraph, holds the numPr).
    let styles_xml = part_text(&path, "word/styles.xml");
    for style_id in ["ListBullet", "ListNumber"] {
        let idx = styles_xml
            .find(&format!("w:styleId=\"{style_id}\""))
            .unwrap_or_else(|| panic!("{style_id} style present"));
        let after = &styles_xml[idx..];
        let end = after.find("</w:style>").unwrap();
        let style_def = &after[..end];
        assert!(
            style_def.contains("<w:numPr>") && style_def.contains("<w:numId"),
            "{style_id} style references numbering via w:numPr/w:numId"
        );
    }
}

// 3. create_numbering builds independent definitions: two decimal lists get distinct
//    numIds and distinct abstractNumIds, numbering.xml keeps all w:abstractNum before all
//    w:num, both survive a round trip, and the paragraphs report the right (numId, ilvl).
#[test]
fn create_numbering_makes_independent_definitions() {
    let mut doc = Document::new();
    let num_a = doc.create_numbering(NumberFormat::Decimal);
    let num_b = doc.create_numbering(NumberFormat::Decimal);
    assert_ne!(num_a, num_b, "the two definitions get distinct numIds");

    let p1 = doc.add_paragraph("list A item");
    p1.set_numbering(&mut doc, num_a, 0);
    let p2 = doc.add_paragraph("list B item");
    p2.set_numbering(&mut doc, num_b, 0);

    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("created.docx");
    doc.save(&path).unwrap();

    // Schema order: every abstractNum precedes every num in word/numbering.xml.
    let tree = part_tree(&path, "word/numbering.xml");
    let root = tree.root();
    let kinds: Vec<String> = tree
        .children(root)
        .iter()
        .filter_map(|&c| local(&tree, c))
        .filter(|n| n == "abstractNum" || n == "num")
        .collect();
    let last_abstract = kinds.iter().rposition(|n| n == "abstractNum").unwrap();
    let first_num = kinds.iter().position(|n| n == "num").unwrap();
    assert!(
        last_abstract < first_num,
        "all w:abstractNum must come before all w:num, got {kinds:?}"
    );

    // The two new definitions map to distinct abstractNumIds.
    let abstract_of = |numbering: &XmlTree, want: u32| -> u32 {
        for c in numbering.children(numbering.root()).iter().copied() {
            if local(numbering, c).as_deref() == Some("num")
                && numbering.attr(c, "w:numId") == Some(&want.to_string())
            {
                let child = numbering.children(c)[0];
                return numbering.attr(child, "w:val").unwrap().parse().unwrap();
            }
        }
        panic!("numId {want} not found");
    };
    let abs_a = abstract_of(&tree, num_a);
    let abs_b = abstract_of(&tree, num_b);
    assert_ne!(abs_a, abs_b, "distinct abstractNumIds");

    // Round trip preserves both, and the paragraphs report their direct numbering.
    let reopened = Document::open(&path).unwrap();
    let paras = reopened.paragraphs();
    assert_eq!(paras[0].numbering(&reopened), Some((num_a, 0)));
    assert_eq!(paras[1].numbering(&reopened), Some((num_b, 0)));
}

// 4. Part-creation path: open a document that has NO numbering part, call create_numbering,
//    and verify it creates the part, registers a content-type Override, adds a numbering
//    relationship, and produces a file that re-opens and parses.
#[test]
fn create_numbering_creates_the_part_when_absent() {
    let dir = tempfile::tempdir().unwrap();
    let stripped = dir.path().join("no-numbering.docx");
    strip_numbering_part(Path::new(BASIC), &stripped);

    // Sanity: the stripped document genuinely lacks the numbering part and its wiring.
    {
        let pkg = Package::open(&stripped).unwrap();
        assert!(pkg.part("word/numbering.xml").is_none());
        let ct = part_text(&stripped, "[Content_Types].xml");
        assert!(!ct.contains("numbering+xml"));
        let rels = part_text(&stripped, "word/_rels/document.xml.rels");
        assert!(!rels.contains("/numbering"));
    }

    let mut doc = Document::open(&stripped).unwrap();
    let num = doc.create_numbering(NumberFormat::Decimal);
    let p = doc.add_paragraph("numbered");
    p.set_numbering(&mut doc, num, 0);

    let out = dir.path().join("with-numbering.docx");
    doc.save(&out).unwrap();

    // The numbering part now exists and parses.
    let pkg = Package::open(&out).unwrap();
    assert!(
        pkg.part("word/numbering.xml").is_some(),
        "numbering part created"
    );
    let numbering = part_tree(&out, "word/numbering.xml");
    assert_eq!(
        local(&numbering, numbering.root()).as_deref(),
        Some("numbering")
    );

    // Content-type Override registered for the new part.
    let ct = part_text(&out, "[Content_Types].xml");
    assert!(
        ct.contains("PartName=\"/word/numbering.xml\"")
            && ct.contains(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
            ),
        "content-type Override for the numbering part is present"
    );

    // Relationship from the document part points at the numbering part.
    let rels = part_text(&out, "word/_rels/document.xml.rels");
    assert!(
        rels.contains("relationships/numbering") && rels.contains("Target=\"numbering.xml\""),
        "numbering relationship present"
    );

    // Re-opens cleanly and the paragraph reports its numbering.
    let reopened = Document::open(&out).unwrap();
    let paras = reopened.paragraphs();
    assert_eq!(paras.last().unwrap().numbering(&reopened), Some((num, 0)));
}

// 5. clear_numbering removes the direct w:numPr, so numbering() no longer reports it.
#[test]
fn clear_numbering_removes_direct_numpr() {
    let mut doc = Document::new();
    let num = doc.create_numbering(NumberFormat::Decimal);
    let p = doc.add_paragraph("item");
    p.set_numbering(&mut doc, num, 0);
    assert_eq!(p.numbering(&doc), Some((num, 0)));

    p.clear_numbering(&mut doc);
    assert_eq!(p.numbering(&doc), None);

    // The w:numPr is gone from the serialized paragraph.
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("cleared.docx");
    doc.save(&path).unwrap();
    let body = part_text(&path, "word/document.xml");
    assert!(!body.contains("<w:numPr>"), "w:numPr removed on clear");
}

// Fidelity regression: opening and re-saving any fixture without touching it keeps every
// part byte-identical (milestone 9 adds read paths that must not perturb this).
#[test]
fn untouched_roundtrip_stays_byte_identical() {
    for fixture in all_fixtures() {
        let doc = Document::open(&fixture).unwrap();
        // Reading numbering must not dirty any part.
        for p in doc.paragraphs() {
            let _ = p.numbering(&doc);
        }

        let dir = tempfile::tempdir().unwrap();
        let saved = dir.path().join("saved.docx");
        doc.save(&saved).unwrap();

        let original = Package::open(&fixture).unwrap();
        let after = Package::open(&saved).unwrap();
        for (a, b) in original.parts().iter().zip(after.parts()) {
            assert_eq!(
                a.data,
                b.data,
                "part {} changed on untouched round trip for {}",
                a.name,
                fixture.display()
            );
        }
    }
}

/// Write a copy of `src` to `dst` with the numbering part, its content-type `Override`, and
/// its document relationship removed — a document that defines no lists at all. Built with
/// the OPC + XML layers directly (no numbering-specific API), so it is a faithful stand-in
/// for a real document opened from a file that never used numbering.
fn strip_numbering_part(src: &Path, dst: &Path) {
    let mut pkg = Package::open(src).unwrap();

    // Drop the Override for /word/numbering.xml from [Content_Types].xml.
    {
        let ct = pkg.part("[Content_Types].xml").unwrap();
        let mut tree = XmlTree::parse(&ct.data).unwrap();
        let root = tree.root();
        if let Some(over) = tree.children(root).iter().copied().find(|&c| {
            local(&tree, c).as_deref() == Some("Override")
                && tree.attr(c, "PartName") == Some("/word/numbering.xml")
        }) {
            tree.remove_from_parent(over);
        }
        pkg.add_part("[Content_Types].xml".to_string(), tree.serialize());
    }

    // Drop the numbering relationship from word/_rels/document.xml.rels.
    {
        let rels = pkg.part("word/_rels/document.xml.rels").unwrap();
        let mut tree = XmlTree::parse(&rels.data).unwrap();
        let root = tree.root();
        let victims: Vec<NodeId> = tree
            .children(root)
            .iter()
            .copied()
            .filter(|&c| {
                local(&tree, c).as_deref() == Some("Relationship")
                    && tree
                        .attr(c, "Type")
                        .is_some_and(|t| t.ends_with("/numbering"))
            })
            .collect();
        for v in victims {
            tree.remove_from_parent(v);
        }
        pkg.add_part("word/_rels/document.xml.rels".to_string(), tree.serialize());
    }

    // Remove the numbering part itself.
    assert!(pkg.remove_part("word/numbering.xml"));

    pkg.save(dst).unwrap();
}
