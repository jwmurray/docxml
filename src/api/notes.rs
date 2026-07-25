//! Footnote and endnote authoring: the [`Footnote`] / [`Endnote`] handles, the
//! [`Paragraph::add_footnote`] / [`Paragraph::add_endnote`] authoring entry points, and the
//! [`Document::footnotes`] / [`Document::endnotes`] read-back accessors.
//!
//! Footnotes live in their own package part (`word/footnotes.xml`, root `w:footnotes`),
//! endnotes in `word/endnotes.xml` (root `w:endnotes`), each referenced from the main
//! document part by relationship. A note is a `w:footnote` / `w:endnote` child of that root
//! carrying a document-unique `w:id`; the body marks the note's location with a
//! `w:footnoteReference` / `w:endnoteReference` run child pointing at that id.
//!
//! Word requires two stub notes in every notes part: the *separator* (the short rule drawn
//! between body text and the notes) at `w:id="-1"` and the *continuation separator* (the
//! full-width rule drawn when a note continues onto the next page) at `w:id="0"`. Creating
//! a notes part authors both stubs; real notes number from `1`. The read-back accessors
//! exclude the stubs, so [`Document::footnotes`] counts only real notes.
//!
//! A created note contains one paragraph whose first run holds the `w:footnoteRef` /
//! `w:endnoteRef` self-number mark (superscripted explicitly, since an arbitrary document
//! may not define Word's `FootnoteReference` character style) followed by a run carrying
//! the note text verbatim — no whitespace is inserted, so a caller that wants Word's
//! conventional space between the number and the text includes it in the text (the examples
//! do). The returned handle's [`paragraphs`](Footnote::paragraphs) carry the notes part's
//! id, so the ordinary [`Paragraph`] / [`Run`](crate::Run) API controls per-note style
//! (`set_style_id`) and point size (`set_size`) — the controls court rules regulate.

use crate::xml::NodeId;

use super::{Document, Paragraph, PartId, Run, is_wml_element};

/// Content type registered for a freshly created footnotes part.
const FOOTNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
/// Content type registered for a freshly created endnotes part.
const ENDNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
/// Relationship types identifying the footnotes part (transitional and strict).
const FOOTNOTES_REL_TYPES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes",
];
/// Relationship types identifying the endnotes part (transitional and strict).
const ENDNOTES_REL_TYPES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes",
];

/// The two stub-note `w:type` values every notes part carries (at `w:id="-1"` / `"0"`),
/// plus the continuation notice — none of which are real, referenceable notes. The
/// read-back accessors filter these out.
const STUB_TYPES: [&str; 3] = ["separator", "continuationSeparator", "continuationNotice"];

/// Whether a footnote or an endnote is being authored/read — selects the part basename,
/// root/note/reference element names, content type, and relationship type that differ
/// between the two otherwise-identical flows (the notes twin of `header.rs`'s `Kind`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NoteKind {
    Footnote,
    Endnote,
}

impl NoteKind {
    /// The part's file-name stem: `footnote` → `word/footnotes.xml`.
    fn basename(self) -> &'static str {
        match self {
            NoteKind::Footnote => "footnotes",
            NoteKind::Endnote => "endnotes",
        }
    }

    /// The part's root element local name (`footnotes` / `endnotes`).
    fn root_local(self) -> &'static str {
        self.basename()
    }

    /// The per-note element local name (`footnote` / `endnote`).
    fn note_local(self) -> &'static str {
        match self {
            NoteKind::Footnote => "footnote",
            NoteKind::Endnote => "endnote",
        }
    }

    /// The body-side reference element local name (`footnoteReference` / `endnoteReference`).
    fn ref_local(self) -> &'static str {
        match self {
            NoteKind::Footnote => "footnoteReference",
            NoteKind::Endnote => "endnoteReference",
        }
    }

    /// The in-note self-number mark local name (`footnoteRef` / `endnoteRef`).
    fn mark_local(self) -> &'static str {
        match self {
            NoteKind::Footnote => "footnoteRef",
            NoteKind::Endnote => "endnoteRef",
        }
    }

    /// The content type registered for a created part of this kind.
    fn content_type(self) -> &'static str {
        match self {
            NoteKind::Footnote => FOOTNOTES_CONTENT_TYPE,
            NoteKind::Endnote => ENDNOTES_CONTENT_TYPE,
        }
    }

    /// The relationship types resolving the part (transitional first, strict second).
    fn rel_types(self) -> &'static [&'static str; 2] {
        match self {
            NoteKind::Footnote => &FOOTNOTES_REL_TYPES,
            NoteKind::Endnote => &ENDNOTES_REL_TYPES,
        }
    }
}

/// A lightweight handle to one `w:footnote` in the footnotes part.
///
/// Obtained from [`Paragraph::add_footnote`] (authoring) or [`Document::footnotes`]
/// (read-back). Like the other handles it is `Copy` and borrows nothing — it carries the
/// [`PartId`] of the footnotes part plus the `w:footnote` node id, so its
/// [`paragraphs`](Self::paragraphs) read and edit that part's tree through the ordinary
/// [`Paragraph`] / [`Run`](crate::Run) API.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Footnote {
    part: PartId,
    node: NodeId,
}

/// A lightweight handle to one `w:endnote` in the endnotes part — the endnote twin of
/// [`Footnote`]; see it for the handle semantics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Endnote {
    part: PartId,
    node: NodeId,
}

macro_rules! note_handle_impl {
    ($ty:ident, $kind:expr, $what:literal) => {
        impl $ty {
            /// The note's `w:id` — the number its body reference mark points at. Real
            /// notes are positive; the separator stubs (excluded from the read-back
            /// accessors) hold `-1` and `0`.
            pub fn id(&self, doc: &Document) -> i64 {
                note_id(doc, self.part, self.node).unwrap_or(0)
            }

            #[doc = concat!("The ", $what, "'s paragraphs, in order.")]
            ///
            /// The returned [`Paragraph`]s carry the notes part's id, so the ordinary
            /// paragraph/run API reads and edits them: `set_style_id` for the note's
            /// paragraph style, [`Run::set_size`](crate::Run::set_size) for the point
            /// size court rules regulate, and so on.
            pub fn paragraphs(&self, doc: &Document) -> Vec<Paragraph> {
                let tree = doc.tree(self.part);
                tree.children(self.node)
                    .iter()
                    .copied()
                    .filter(|&c| is_wml_element(tree, c, "p"))
                    .map(|c| Paragraph::from_node(self.part, c))
                    .collect()
            }

            #[doc = concat!("Append a paragraph carrying `text` to the ", $what, ", returning it.")]
            ///
            /// The text is written verbatim (a run is added only when it is non-empty).
            /// Unlike the first paragraph a created note starts with, an appended
            /// paragraph carries no self-number mark — notes number once.
            pub fn add_paragraph(&self, doc: &mut Document, text: &str) -> Paragraph {
                let name = doc.qn(self.part, "p");
                let tree = doc.tree_mut(self.part);
                let p = tree.create_element(name);
                tree.append_child(self.node, p);
                let para = Paragraph::from_node(self.part, p);
                if !text.is_empty() {
                    para.add_run(doc, text);
                }
                para
            }

            #[doc = concat!("The ", $what, "'s text: every paragraph's text, joined with `\\n`.")]
            ///
            /// The self-number mark contributes nothing (it is not text), so this reads
            /// back exactly what was authored.
            pub fn text(&self, doc: &Document) -> String {
                self.paragraphs(doc)
                    .iter()
                    .map(|p| p.text(doc))
                    .collect::<Vec<_>>()
                    .join("\n")
            }

            #[doc = concat!("The explicit point sizes of the ", $what, "'s text-bearing runs, in order.")]
            ///
            /// Runs without an explicit `w:sz` (inheriting their size from a style or the
            /// document defaults) are skipped, as are the mark-only runs, which carry no
            /// text. This is the read-back a format probe asserts against a court rule
            /// that regulates note point size.
            pub fn point_sizes(&self, doc: &Document) -> Vec<super::Pt> {
                self.paragraphs(doc)
                    .iter()
                    .flat_map(|p| p.runs(doc))
                    .filter(|r| !r.text(doc).is_empty())
                    .filter_map(|r| r.size(doc))
                    .collect()
            }
        }
    };
}

note_handle_impl!(Footnote, NoteKind::Footnote, "footnote");
note_handle_impl!(Endnote, NoteKind::Endnote, "endnote");

impl Paragraph {
    /// Append a footnote: a superscripted reference mark at the end of this paragraph and
    /// a new note in the footnotes part carrying `text`, returning the note's handle.
    ///
    /// The footnotes part is created on first use — `word/footnotes.xml` with the two
    /// stub notes Word requires (the separator at `w:id="-1"` and the continuation
    /// separator at `w:id="0"`), a `[Content_Types].xml` override, and a `footnotes`
    /// relationship from the document part. Real notes number consecutively from `1`.
    ///
    /// The created note holds one paragraph: the `w:footnoteRef` self-number mark
    /// (explicitly superscripted) followed by a run carrying `text` verbatim — include a
    /// leading space for Word's conventional gap between the number and the text. Style
    /// and size the note through the returned handle's [`paragraphs`](Footnote::paragraphs).
    ///
    /// # Panics
    ///
    /// Panics only if the notes part cannot be created or parsed — which for a document
    /// built from [`Document::new`] or opened from a valid package does not happen.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, Pt};
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("As the court held, the motion fails.");
    /// let note = p.add_footnote(&mut doc, " See the accompanying declaration.");
    /// // Point-size control per note (some courts regulate footnote size):
    /// for para in note.paragraphs(&doc) {
    ///     for run in para.runs(&doc) {
    ///         run.set_size(&mut doc, Pt(12.0));
    ///     }
    /// }
    /// assert_eq!(doc.footnotes().len(), 1);
    /// ```
    pub fn add_footnote(&self, doc: &mut Document, text: &str) -> Footnote {
        let (part, node) = add_note(self, doc, NoteKind::Footnote, text);
        Footnote { part, node }
    }

    /// Append an endnote: the endnote twin of [`add_footnote`](Self::add_footnote) —
    /// same part-creation (`word/endnotes.xml`), stub notes, numbering, and note shape,
    /// with the reference mark written as `w:endnoteReference`.
    ///
    /// # Panics
    ///
    /// As [`add_footnote`](Self::add_footnote).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("Discussed further below.");
    /// p.add_endnote(&mut doc, " Full citation in the appendix.");
    /// assert_eq!(doc.endnotes().len(), 1);
    /// ```
    pub fn add_endnote(&self, doc: &mut Document, text: &str) -> Endnote {
        let (part, node) = add_note(self, doc, NoteKind::Endnote, text);
        Endnote { part, node }
    }
}

impl Run {
    /// The id of the footnote this run references (`w:footnoteReference w:id`), or `None`
    /// when the run holds no footnote reference — the body-side read-back that pairs a
    /// reference mark with [`Document::footnotes`].
    pub fn footnote_reference_id(&self, doc: &Document) -> Option<i64> {
        reference_id(doc, self.part(), self.node(), NoteKind::Footnote)
    }

    /// The id of the endnote this run references (`w:endnoteReference w:id`), or `None`
    /// when the run holds no endnote reference.
    pub fn endnote_reference_id(&self, doc: &Document) -> Option<i64> {
        reference_id(doc, self.part(), self.node(), NoteKind::Endnote)
    }
}

impl Document {
    /// The document's real footnotes, in part order — the `w:footnote` children of the
    /// footnotes part, excluding the separator/continuation stub notes.
    ///
    /// Takes `&mut self` because the footnotes part is parsed lazily on first access;
    /// parsing alone does not mark it modified, so a read-only pass leaves every part
    /// byte-identical on save. Returns an empty vector when the document has no footnotes
    /// part. `footnotes().len()` is the footnote count a format probe asserts.
    pub fn footnotes(&mut self) -> Vec<Footnote> {
        notes(self, NoteKind::Footnote)
            .into_iter()
            .map(|(part, node)| Footnote { part, node })
            .collect()
    }

    /// The document's real endnotes, in part order — the endnote twin of
    /// [`footnotes`](Self::footnotes).
    pub fn endnotes(&mut self) -> Vec<Endnote> {
        notes(self, NoteKind::Endnote)
            .into_iter()
            .map(|(part, node)| Endnote { part, node })
            .collect()
    }
}

/// A note element's `w:id`, parsed as the signed decimal ST_DecimalNumber it is.
fn note_id(doc: &Document, part: PartId, node: NodeId) -> Option<i64> {
    let tree = doc.tree(part);
    tree.attr(node, &doc.qn(part, "id"))
        .and_then(|v| v.trim().parse::<i64>().ok())
}

/// Whether a note element is one of the stub notes (separator, continuation separator,
/// continuation notice) rather than a real, referenceable note.
fn is_stub_note(doc: &Document, part: PartId, node: NodeId) -> bool {
    let tree = doc.tree(part);
    matches!(tree.attr(node, &doc.qn(part, "type")), Some(t) if STUB_TYPES.contains(&t))
}

/// The id of the note of `kind` referenced by a run, if any: the `w:id` of a direct
/// `w:footnoteReference` / `w:endnoteReference` child.
fn reference_id(doc: &Document, part: PartId, run: NodeId, kind: NoteKind) -> Option<i64> {
    let tree = doc.tree(part);
    let el = tree
        .children(run)
        .iter()
        .copied()
        .find(|&c| is_wml_element(tree, c, kind.ref_local()))?;
    tree.attr(el, &doc.qn(part, "id"))
        .and_then(|v| v.trim().parse::<i64>().ok())
}

/// The real (non-stub) notes of `kind`, as `(part, node)` pairs, or empty when the
/// document has no notes part of that kind.
fn notes(doc: &mut Document, kind: NoteKind) -> Vec<(PartId, NodeId)> {
    let Some(part_name) = doc.part_by_rel_type(kind.rel_types()) else {
        return Vec::new();
    };
    let Some(part) = doc.ensure_part(&part_name) else {
        return Vec::new();
    };
    let tree = doc.tree(part);
    let root = tree.root();
    tree.children(root)
        .iter()
        .copied()
        .filter(|&c| is_wml_element(tree, c, kind.note_local()))
        .filter(|&c| !is_stub_note(doc, part, c))
        .map(|c| (part, c))
        .collect()
}

/// Shared authoring path for [`Paragraph::add_footnote`] / [`Paragraph::add_endnote`]:
/// ensure the notes part (creating it with its stub notes on first use), allocate the next
/// id, author the note element with its first paragraph, and append the reference-mark run
/// to the body paragraph. Returns the notes part id and the new note's node.
fn add_note(para: &Paragraph, doc: &mut Document, kind: NoteKind, text: &str) -> (PartId, NodeId) {
    let part = ensure_notes_part(doc, kind);

    // Next free note id: max over existing w:footnote/w:endnote ids (stubs included, so a
    // part holding only the -1/0 stubs starts real notes at 1).
    let tree = doc.tree(part);
    let root = tree.root();
    let next_id = tree
        .children(root)
        .iter()
        .copied()
        .filter(|&c| is_wml_element(tree, c, kind.note_local()))
        .filter_map(|c| note_id(doc, part, c))
        .max()
        .unwrap_or(0)
        .max(0)
        + 1;

    // The note element: <w:footnote w:id="N"> with one paragraph holding the
    // superscripted self-number mark, then the text verbatim.
    let note_name = doc.qn(part, kind.note_local());
    let id_attr = doc.qn(part, "id");
    let tree = doc.tree_mut(part);
    let note = tree.create_element(note_name);
    tree.set_attr(note, id_attr, next_id.to_string());
    tree.append_child(root, note);

    let note_para = Paragraph::from_node(part, {
        let p_name = doc.qn(part, "p");
        let tree = doc.tree_mut(part);
        let p = tree.create_element(p_name);
        tree.append_child(note, p);
        p
    });
    append_mark_run(doc, part, note_para.node(), kind.mark_local(), None);
    if !text.is_empty() {
        note_para.add_run(doc, text);
    }

    // The body-side reference mark: a superscripted run holding
    // <w:footnoteReference w:id="N"> at the end of the source paragraph.
    append_mark_run(
        doc,
        para.part(),
        para.node(),
        kind.ref_local(),
        Some(next_id),
    );

    (part, note)
}

/// Append a run holding a single note-mark element (`w:footnoteRef`, `w:footnoteReference`,
/// …) to `parent`, superscripted via an explicit `w:rPr/w:vertAlign` — Word's built-in
/// `FootnoteReference` character style may not exist in an arbitrary document, so the
/// formatting is written directly. `id` writes a `w:id` on the mark (the body-side
/// reference form); `None` writes a bare mark (the in-note self-number form).
fn append_mark_run(doc: &mut Document, part: PartId, parent: NodeId, local: &str, id: Option<i64>) {
    let r_name = doc.qn(part, "r");
    let rpr_name = doc.qn(part, "rPr");
    let vert_name = doc.qn(part, "vertAlign");
    let val_attr = doc.qn(part, "val");
    let mark_name = doc.qn(part, local);
    let id_attr = doc.qn(part, "id");

    let tree = doc.tree_mut(part);
    let r = tree.create_element(r_name);
    let rpr = tree.create_element(rpr_name);
    let vert = tree.create_element(vert_name);
    tree.set_attr(vert, val_attr, "superscript");
    tree.append_child(rpr, vert);
    tree.append_child(r, rpr);
    let mark = tree.create_element(mark_name);
    if let Some(id) = id {
        tree.set_attr(mark, id_attr, id.to_string());
    }
    tree.append_child(r, mark);
    tree.append_child(parent, r);
}

/// Ensure the notes part of `kind` exists and is parsed, creating it on first use: a
/// `word/footnotes.xml` / `word/endnotes.xml` part whose root carries the main document
/// root's namespace declarations and the two stub notes, a content-type override, and a
/// relationship from the document part (the notes twin of `header.rs`'s create path).
fn ensure_notes_part(doc: &mut Document, kind: NoteKind) -> PartId {
    if let Some(existing) = doc.part_by_rel_type(kind.rel_types()) {
        return doc
            .ensure_part(&existing)
            .expect("existing notes part parses");
    }

    // The document part's directory (`word/`), where the new part and its relative
    // relationship target live.
    let source = doc.main_part_name().to_string();
    let dir = match source.rfind('/') {
        Some(i) => source[..=i].to_string(),
        None => String::new(),
    };

    let part_name = format!("{dir}{}.xml", kind.basename());
    let xml = build_notes_xml(doc, kind);
    doc.add_part(part_name.clone(), xml);
    doc.ensure_content_type_override(&format!("/{part_name}"), kind.content_type())
        .expect("[Content_Types].xml is editable");
    let target = part_name
        .strip_prefix(&dir)
        .unwrap_or(&part_name)
        .to_string();
    doc.add_relationship(&source, kind.rel_types()[0], &target, false)
        .expect("document relationships part is editable");

    doc.ensure_part(&part_name)
        .expect("created notes part parses")
}

/// Build the raw bytes of a minimal, self-contained notes part: an XML declaration, a
/// `w:footnotes` / `w:endnotes` root carrying every namespace declaration from the main
/// document root (matching `header.rs`'s part construction), and the two stub notes Word
/// requires — the separator at `w:id="-1"` and the continuation separator at `w:id="0"`,
/// each a single-line paragraph (zero spacing, so the rule stays a rule) holding its
/// `w:separator` / `w:continuationSeparator` mark.
fn build_notes_xml(doc: &Document, kind: NoteKind) -> Vec<u8> {
    let root_name = doc.qn(PartId::MAIN, kind.root_local());
    let note = doc.qn(PartId::MAIN, kind.note_local());
    let p = doc.qn(PartId::MAIN, "p");
    let ppr = doc.qn(PartId::MAIN, "pPr");
    let spacing = doc.qn(PartId::MAIN, "spacing");
    let after = doc.qn(PartId::MAIN, "after");
    let line = doc.qn(PartId::MAIN, "line");
    let line_rule = doc.qn(PartId::MAIN, "lineRule");
    let r = doc.qn(PartId::MAIN, "r");
    let type_attr = doc.qn(PartId::MAIN, "type");
    let id_attr = doc.qn(PartId::MAIN, "id");
    let separator = doc.qn(PartId::MAIN, "separator");
    let continuation = doc.qn(PartId::MAIN, "continuationSeparator");

    let main = doc.tree(PartId::MAIN);
    let root = main.root();
    let mut decls = String::new();
    for (key, value) in main.attrs(root) {
        if key == "xmlns" || key.starts_with("xmlns:") {
            decls.push(' ');
            decls.push_str(key);
            decls.push_str("=\"");
            decls.push_str(value);
            decls.push('"');
        }
    }

    let stub = |type_val: &str, id: i64, mark: &str| {
        format!(
            "<{note} {type_attr}=\"{type_val}\" {id_attr}=\"{id}\">\
             <{p}><{ppr}><{spacing} {after}=\"0\" {line}=\"240\" {line_rule}=\"auto\"/></{ppr}>\
             <{r}><{mark}/></{r}></{p}></{note}>"
        )
    };
    let separator_stub = stub("separator", -1, &separator);
    let continuation_stub = stub("continuationSeparator", 0, &continuation);

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <{root_name}{decls}>{separator_stub}{continuation_stub}</{root_name}>"
    )
    .into_bytes()
}
