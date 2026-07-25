//! Comments and tracked changes: part creation, body anchors, final-view text
//! semantics, read-back, and the author → save → reopen round trip.

use std::io::Cursor;

use docxml::{Document, RevisionKind};

fn reopen(doc: &Document) -> Document {
    let mut buf = Cursor::new(Vec::new());
    doc.write(&mut buf).expect("document writes");
    Document::read(Cursor::new(buf.into_inner())).expect("written document reopens")
}

#[test]
fn comments_round_trip_with_anchors() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("The motion is timely.");
    let c = p.add_comment(&mut doc, "Reviewer", "R", "Cite the scheduling order.");
    assert_eq!(c.id(&doc), 0);

    let mut reopened = reopen(&doc);
    let comments = reopened.comments();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].id(&reopened), 0);
    assert_eq!(comments[0].author(&reopened), "Reviewer");
    assert_eq!(comments[0].initials(&reopened), "R");
    assert_eq!(comments[0].text(&reopened), "Cite the scheduling order.");

    // The body carries the full anchor set, sharing the comment's id.
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    assert!(xml.contains(r#"<w:commentRangeStart w:id="0"/>"#));
    assert!(xml.contains(r#"<w:commentRangeEnd w:id="0"/>"#));
    assert!(xml.contains(r#"<w:commentReference w:id="0"/>"#));

    // The comments part is registered and related.
    let ct = reopened
        .package()
        .part("[Content_Types].xml")
        .expect("content types");
    assert!(
        String::from_utf8_lossy(&ct.data).contains(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
        )
    );

    // The anchors do not disturb the paragraph's visible text.
    assert_eq!(
        reopened.paragraphs()[0].text(&reopened),
        "The motion is timely."
    );
}

#[test]
fn comment_ids_increment() {
    let mut doc = Document::new();
    let p1 = doc.add_paragraph("First.");
    let p2 = doc.add_paragraph("Second.");
    p1.add_comment(&mut doc, "A", "A", "one");
    p2.add_comment(&mut doc, "B", "B", "two");

    let mut reopened = reopen(&doc);
    let ids: Vec<i64> = reopened
        .comments()
        .iter()
        .map(|c| c.id(&reopened))
        .collect();
    assert_eq!(ids, [0, 1]);
}

#[test]
fn tracked_insertion_reads_as_final_text() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("The court ");
    let run = p.add_inserted_run(&mut doc, "respectfully ", "Editor");
    run.italic(&mut doc, true); // the returned run formats like any other

    let reopened = reopen(&doc);
    let p = reopened.paragraphs()[0];
    assert_eq!(p.text(&reopened), "The court respectfully ");

    let revs = p.revisions(&reopened);
    assert_eq!(revs.len(), 1);
    assert_eq!(revs[0].kind, RevisionKind::Inserted);
    assert_eq!(revs[0].author, "Editor");
    assert_eq!(revs[0].text, "respectfully ");
}

#[test]
fn tracked_deletion_preserves_text_outside_the_final_view() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Now ");
    p.add_deleted_run(&mut doc, "therefore ", "Editor");
    p.add_run(&mut doc, "comes Plaintiff.");

    let reopened = reopen(&doc);
    let p = reopened.paragraphs()[0];
    // Final view: the deleted text is gone...
    assert_eq!(p.text(&reopened), "Now comes Plaintiff.");
    // ...but preserved, attributed, in the revision read.
    let revs = p.revisions(&reopened);
    assert_eq!(revs.len(), 1);
    assert_eq!(revs[0].kind, RevisionKind::Deleted);
    assert_eq!(revs[0].text, "therefore ");

    // w:delText (with space preserved), not w:t.
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    assert!(xml.contains(r#"<w:delText xml:space="preserve">therefore </w:delText>"#));
}

#[test]
fn revision_ids_are_unique_across_kinds() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Base ");
    p.add_inserted_run(&mut doc, "ins one ", "A");
    p.add_deleted_run(&mut doc, "del one ", "A");
    p.add_inserted_run(&mut doc, "ins two ", "B");

    let reopened = reopen(&doc);
    let part = reopened
        .package()
        .part("word/document.xml")
        .expect("main part");
    let xml = String::from_utf8_lossy(&part.data);
    for id in 1..=3 {
        assert!(
            xml.contains(&format!(r#"w:id="{id}" w:author="#)),
            "revision id {id} missing"
        );
    }

    let revs = reopened.paragraphs()[0].revisions(&reopened);
    assert_eq!(revs.len(), 3);
    assert_eq!(
        revs.iter().map(|r| r.kind).collect::<Vec<_>>(),
        [
            RevisionKind::Inserted,
            RevisionKind::Deleted,
            RevisionKind::Inserted
        ]
    );
}

#[test]
fn documents_without_comments_read_back_empty() {
    let mut doc = Document::new();
    assert!(doc.comments().is_empty());
    assert!(doc.paragraphs().is_empty());
}
