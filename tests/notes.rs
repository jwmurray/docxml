//! Footnote and endnote authoring: part creation with stub notes, reference marks,
//! style/point-size control, read-back accessors, and the author → save → reopen round
//! trip.

use std::io::Cursor;

use docxml::{Document, Pt};

/// Save a document to memory and reopen it — the round-trip half of every test here.
fn reopen(doc: &Document) -> Document {
    let mut buf = Cursor::new(Vec::new());
    doc.write(&mut buf).expect("document writes");
    Document::read(Cursor::new(buf.into_inner())).expect("written document reopens")
}

#[test]
fn add_footnote_creates_part_with_stub_notes() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Body text.");
    p.add_footnote(&mut doc, " First note.");

    let reopened = reopen(&doc);

    // The part exists, with the two Word-required stub notes and the real note.
    let part = reopened
        .package()
        .part("word/footnotes.xml")
        .expect("footnotes part exists");
    let xml = String::from_utf8_lossy(&part.data);
    assert!(xml.contains(r#"w:type="separator" w:id="-1""#));
    assert!(xml.contains(r#"w:type="continuationSeparator" w:id="0""#));
    assert!(xml.contains(r#"<w:separator/>"#));
    assert!(xml.contains(r#"<w:continuationSeparator/>"#));
    assert!(xml.contains(r#"<w:footnote w:id="1">"#));

    // The content type and relationship were registered.
    let ct = reopened
        .package()
        .part("[Content_Types].xml")
        .expect("content types part exists");
    assert!(
        String::from_utf8_lossy(&ct.data).contains(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
        )
    );
    let rels = reopened
        .package()
        .part("word/_rels/document.xml.rels")
        .expect("document rels exist");
    assert!(String::from_utf8_lossy(&rels.data).contains("relationships/footnotes"));
}

#[test]
fn footnotes_round_trip_with_text_ids_and_marks() {
    let mut doc = Document::new();
    let p1 = doc.add_paragraph("First paragraph.");
    let p2 = doc.add_paragraph("Second paragraph.");
    p1.add_footnote(&mut doc, " See the declaration.");
    p2.add_footnote(&mut doc, " Same rule, later case.");

    let mut reopened = reopen(&doc);

    // Read-back: count and text, stubs excluded.
    let notes = reopened.footnotes();
    assert_eq!(notes.len(), 2);
    assert_eq!(notes[0].id(&reopened), 1);
    assert_eq!(notes[1].id(&reopened), 2);
    assert_eq!(notes[0].text(&reopened), " See the declaration.");
    assert_eq!(notes[1].text(&reopened), " Same rule, later case.");

    // Body-side reference marks pair with the note ids, in paragraph order.
    let paras = reopened.paragraphs();
    let ref_ids: Vec<i64> = paras
        .iter()
        .flat_map(|p| p.runs(&reopened))
        .filter_map(|r| r.footnote_reference_id(&reopened))
        .collect();
    assert_eq!(ref_ids, [1, 2]);

    // The reference mark does not disturb the paragraph's visible text.
    assert_eq!(paras[0].text(&reopened), "First paragraph.");
}

#[test]
fn footnote_style_and_point_size_survive_round_trip() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Regulated footnote sizes.");
    let note = p.add_footnote(&mut doc, " Sized per court rule.");

    // Per-note paragraph style + point size — the controls CRC 2.108-style rules regulate.
    let note_para = note.paragraphs(&doc)[0];
    note_para.set_style_id(&mut doc, "FootnoteText");
    for run in note_para.runs(&doc) {
        run.set_size(&mut doc, Pt(12.0));
    }

    let mut reopened = reopen(&doc);
    let notes = reopened.footnotes();
    assert_eq!(notes.len(), 1);
    assert_eq!(
        notes[0].paragraphs(&reopened)[0]
            .style_id(&reopened)
            .as_deref(),
        Some("FootnoteText")
    );
    assert_eq!(notes[0].point_sizes(&reopened), [Pt(12.0)]);
}

#[test]
fn multi_paragraph_footnote_reads_back_joined() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Body.");
    let note = p.add_footnote(&mut doc, " First note paragraph.");
    note.add_paragraph(&mut doc, "Second note paragraph.");

    let mut reopened = reopen(&doc);
    let notes = reopened.footnotes();
    assert_eq!(notes[0].paragraphs(&reopened).len(), 2);
    assert_eq!(
        notes[0].text(&reopened),
        " First note paragraph.\nSecond note paragraph."
    );
}

#[test]
fn endnotes_round_trip_independently_of_footnotes() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("Endnote-bearing paragraph.");
    p.add_endnote(&mut doc, " The endnote.");
    p.add_footnote(&mut doc, " The footnote.");

    let mut reopened = reopen(&doc);

    let endnotes = reopened.endnotes();
    assert_eq!(endnotes.len(), 1);
    assert_eq!(endnotes[0].id(&reopened), 1);
    assert_eq!(endnotes[0].text(&reopened), " The endnote.");
    assert_eq!(reopened.footnotes().len(), 1);

    // Endnotes live in their own part with their own stubs.
    let part = reopened
        .package()
        .part("word/endnotes.xml")
        .expect("endnotes part exists");
    let xml = String::from_utf8_lossy(&part.data);
    assert!(xml.contains(r#"w:type="separator" w:id="-1""#));
    assert!(xml.contains(r#"<w:endnote w:id="1">"#));

    let run_ids: Vec<i64> = reopened
        .paragraphs()
        .iter()
        .flat_map(|p| p.runs(&reopened))
        .filter_map(|r| r.endnote_reference_id(&reopened))
        .collect();
    assert_eq!(run_ids, [1]);
}

#[test]
fn footnote_ids_continue_after_reopen() {
    let mut doc = Document::new();
    let p = doc.add_paragraph("One.");
    p.add_footnote(&mut doc, " Note one.");

    let mut reopened = reopen(&doc);
    let p2 = reopened.add_paragraph("Two.");
    p2.add_footnote(&mut reopened, " Note two.");

    let mut round_two = reopen(&reopened);
    let notes = round_two.footnotes();
    assert_eq!(notes.len(), 2);
    assert_eq!(notes[1].id(&round_two), 2);
}

#[test]
fn footnotes_read_back_empty_without_a_part() {
    let mut doc = Document::new();
    assert!(doc.footnotes().is_empty());
    assert!(doc.endnotes().is_empty());
}
