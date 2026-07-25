//! Redlining: comment authoring ([`Paragraph::add_comment`], [`Document::comments`]) and
//! tracked-change authoring ([`Paragraph::add_inserted_run`], [`Paragraph::add_deleted_run`],
//! [`Paragraph::revisions`]).
//!
//! Comments live in their own package part (`word/comments.xml`, root `w:comments`),
//! referenced from the main document part by relationship — the same part pattern as
//! headers and notes. A comment is anchored in the body by a `w:commentRangeStart` /
//! `w:commentRangeEnd` pair plus a run holding `w:commentReference`, all sharing the
//! comment's `w:id`.
//!
//! Tracked changes are body-inline: an insertion wraps ordinary runs in `w:ins`, and a
//! deletion wraps runs whose text lives in `w:delText` (not `w:t` — that is what lets a
//! "final" text read skip deleted content for free) in `w:del`. Both carry `w:id` and
//! `w:author`; `w:date` is optional in the schema and is written only when the caller
//! provides one, so authoring stays deterministic.
//!
//! [`Paragraph::text`](super::Paragraph::text) reads as Word's *final* view: inserted runs
//! contribute their text, deleted runs contribute nothing.

use crate::xml::NodeId;

use super::{Document, Paragraph, PartId, Run, is_wml_element};

/// Content type registered for a freshly created comments part.
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
/// Relationship types identifying the comments part (transitional and strict).
const COMMENTS_REL_TYPES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/comments",
];

/// A lightweight handle to one `w:comment` in the comments part.
///
/// Obtained from [`Paragraph::add_comment`] (authoring) or [`Document::comments`]
/// (read-back). Like the other handles it is `Copy` and borrows nothing; its
/// [`paragraphs`](Self::paragraphs) carry the comments part's id, so the ordinary
/// [`Paragraph`] / [`Run`](crate::Run) API reads and edits the comment body.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Comment {
    part: PartId,
    node: NodeId,
}

impl Comment {
    /// The comment's `w:id` — the number its body range and reference mark share.
    pub fn id(&self, doc: &Document) -> i64 {
        let tree = doc.tree(self.part);
        tree.attr(self.node, &doc.qn(self.part, "id"))
            .and_then(|v| v.trim().parse::<i64>().ok())
            .unwrap_or(0)
    }

    /// The comment's author (`w:author`), or an empty string when unset.
    pub fn author(&self, doc: &Document) -> String {
        self.attr(doc, "author")
    }

    /// The comment's author initials (`w:initials`), or an empty string when unset.
    pub fn initials(&self, doc: &Document) -> String {
        self.attr(doc, "initials")
    }

    /// The comment's paragraphs, in order, as ordinary [`Paragraph`] handles into the
    /// comments part.
    pub fn paragraphs(&self, doc: &Document) -> Vec<Paragraph> {
        let tree = doc.tree(self.part);
        tree.children(self.node)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "p"))
            .map(|c| Paragraph::from_node(self.part, c))
            .collect()
    }

    /// The comment's text: every paragraph's text, joined with `\n`.
    pub fn text(&self, doc: &Document) -> String {
        self.paragraphs(doc)
            .iter()
            .map(|p| p.text(doc))
            .collect::<Vec<_>>()
            .join("\n")
    }

    fn attr(&self, doc: &Document, local: &str) -> String {
        let tree = doc.tree(self.part);
        tree.attr(self.node, &doc.qn(self.part, local))
            .unwrap_or_default()
            .to_string()
    }
}

/// One tracked change in a paragraph, as read back by [`Paragraph::revisions`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Revision {
    /// Whether the change inserts or deletes text.
    pub kind: RevisionKind,
    /// The change's `w:author`, empty when unset.
    pub author: String,
    /// The changed text: inserted text for an insertion, the deleted text (from
    /// `w:delText`) for a deletion.
    pub text: String,
}

/// The two tracked-change kinds this API authors.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionKind {
    /// A `w:ins` — text added under change tracking.
    Inserted,
    /// A `w:del` — text removed under change tracking, preserved as `w:delText`.
    Deleted,
}

impl Paragraph {
    /// Comment on this paragraph: wrap it in a comment range and append a reference
    /// mark, creating the comment (with one paragraph carrying `text`) in the comments
    /// part. Returns the comment's handle.
    ///
    /// The comments part (`word/comments.xml`) is created on first use, with its
    /// content-type override and document relationship. The range spans the whole
    /// paragraph: `w:commentRangeStart` first, then `w:commentRangeEnd` and the
    /// `w:commentReference` run appended after the existing content. `author` and
    /// `initials` are written on the `w:comment`; no `w:date` is written (the schema
    /// makes it optional), so authoring is deterministic.
    ///
    /// # Panics
    ///
    /// Panics only if the comments part cannot be created or parsed — which for a
    /// document built from [`Document::new`] or opened from a valid package does not
    /// happen.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("The motion is timely.");
    /// p.add_comment(&mut doc, "Reviewer", "R", "Cite the scheduling order here.");
    /// assert_eq!(doc.comments().len(), 1);
    /// ```
    pub fn add_comment(
        &self,
        doc: &mut Document,
        author: &str,
        initials: &str,
        text: &str,
    ) -> Comment {
        let part = ensure_comments_part(doc);

        // Next free comment id: max over existing w:comment ids, floored at 0.
        let tree = doc.tree(part);
        let root = tree.root();
        let next_id = tree
            .children(root)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "comment"))
            .filter_map(|c| {
                doc.tree(part)
                    .attr(c, &doc.qn(part, "id"))
                    .and_then(|v| v.trim().parse::<i64>().ok())
            })
            .max()
            .unwrap_or(-1)
            + 1;

        // The w:comment element: id, author, initials, one paragraph of text.
        let comment_name = doc.qn(part, "comment");
        let id_attr = doc.qn(part, "id");
        let author_attr = doc.qn(part, "author");
        let initials_attr = doc.qn(part, "initials");
        let p_name = doc.qn(part, "p");
        let tree = doc.tree_mut(part);
        let comment = tree.create_element(comment_name);
        // CT_Comment attribute order follows CT_TrackChange: id, author, then initials.
        tree.set_attr(comment, id_attr, next_id.to_string());
        tree.set_attr(comment, author_attr, author);
        tree.set_attr(comment, initials_attr, initials);
        tree.append_child(root, comment);
        let p = tree.create_element(p_name);
        tree.append_child(comment, p);
        let para = Paragraph::from_node(part, p);
        if !text.is_empty() {
            para.add_run(doc, text);
        }

        // Body anchors: range start first, range end and the reference run last.
        let start_name = doc.qn(self.part(), "commentRangeStart");
        let end_name = doc.qn(self.part(), "commentRangeEnd");
        let ref_name = doc.qn(self.part(), "commentReference");
        let r_name = doc.qn(self.part(), "r");
        let body_id_attr = doc.qn(self.part(), "id");
        let node = self.node();
        let tree = doc.tree_mut(self.part());
        let start = tree.create_element(start_name);
        tree.set_attr(start, body_id_attr.clone(), next_id.to_string());
        tree.insert_child(node, 0, start);
        let end = tree.create_element(end_name);
        tree.set_attr(end, body_id_attr.clone(), next_id.to_string());
        tree.append_child(node, end);
        let r = tree.create_element(r_name);
        let reference = tree.create_element(ref_name);
        tree.set_attr(reference, body_id_attr, next_id.to_string());
        tree.append_child(r, reference);
        tree.append_child(node, r);

        Comment {
            part,
            node: comment,
        }
    }

    /// Append a tracked insertion: a run carrying `text`, wrapped in `w:ins` attributed
    /// to `author`. Returns the inner run for further formatting.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("The court ");
    /// p.add_inserted_run(&mut doc, "respectfully ", "Editor");
    /// assert_eq!(p.text(&doc), "The court respectfully ");
    /// ```
    pub fn add_inserted_run(&self, doc: &mut Document, text: &str, author: &str) -> Run {
        let (_, run) = self.tracked_wrapper(doc, "ins", author);
        let run = Run::from_node(self.part(), run);
        run.set_text(doc, text);
        run
    }

    /// Append a tracked deletion: a run whose text is preserved as `w:delText`, wrapped
    /// in `w:del` attributed to `author`. Returns the inner run.
    ///
    /// Deleted text does not appear in [`text`](Self::text) — that read is the *final*
    /// view — but reads back through [`revisions`](Self::revisions).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, RevisionKind};
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("Now ");
    /// p.add_deleted_run(&mut doc, "therefore ", "Editor");
    /// assert_eq!(p.text(&doc), "Now ");
    /// assert_eq!(p.revisions(&doc)[0].kind, RevisionKind::Deleted);
    /// ```
    pub fn add_deleted_run(&self, doc: &mut Document, text: &str, author: &str) -> Run {
        let (_, run) = self.tracked_wrapper(doc, "del", author);
        let deltext_name = doc.qn(self.part(), "delText");
        let tree = doc.tree_mut(self.part());
        let dt = tree.create_element(deltext_name);
        if super::needs_space_preserve(text) {
            tree.set_attr(dt, "xml:space", "preserve");
        }
        let content = tree.create_text(text);
        tree.append_child(dt, content);
        tree.append_child(run, dt);
        Run::from_node(self.part(), run)
    }

    /// The paragraph's tracked changes, in document order: one [`Revision`] per direct
    /// `w:ins` / `w:del` child, with its author and its inserted or deleted text.
    pub fn revisions(&self, doc: &Document) -> Vec<Revision> {
        let tree = doc.tree(self.part());
        let author_attr = doc.qn(self.part(), "author");
        tree.children(self.node())
            .iter()
            .copied()
            .filter_map(|c| {
                let kind = if is_wml_element(tree, c, "ins") {
                    RevisionKind::Inserted
                } else if is_wml_element(tree, c, "del") {
                    RevisionKind::Deleted
                } else {
                    return None;
                };
                let author = tree.attr(c, &author_attr).unwrap_or_default().to_string();
                let mut text = String::new();
                for id in tree.descendants(c) {
                    if let Some(name) = tree.name(id) {
                        let local = super::split_qname(name).1;
                        if local == "t" || local == "delText" {
                            text.push_str(&tree.text_content(id));
                        }
                    }
                }
                Some(Revision { kind, author, text })
            })
            .collect()
    }

    /// Create a `w:ins`/`w:del` wrapper (with the next free revision id and `author`)
    /// holding one empty run, both appended to this paragraph. Returns `(wrapper, run)`.
    fn tracked_wrapper(&self, doc: &mut Document, local: &str, author: &str) -> (NodeId, NodeId) {
        let next_id = next_revision_id(doc, self.part());
        let wrap_name = doc.qn(self.part(), local);
        let id_attr = doc.qn(self.part(), "id");
        let author_attr = doc.qn(self.part(), "author");
        let r_name = doc.qn(self.part(), "r");
        let node = self.node();
        let tree = doc.tree_mut(self.part());
        let wrapper = tree.create_element(wrap_name);
        // CT_TrackChange attribute order: id, author (w:date optional, not written).
        tree.set_attr(wrapper, id_attr, next_id.to_string());
        tree.set_attr(wrapper, author_attr, author);
        tree.append_child(node, wrapper);
        let r = tree.create_element(r_name);
        tree.append_child(wrapper, r);
        (wrapper, r)
    }
}

/// The next free tracked-change id in `part`: max over every `w:ins`/`w:del` id, plus one.
fn next_revision_id(doc: &Document, part: PartId) -> i64 {
    let tree = doc.tree(part);
    let id_attr = doc.qn(part, "id");
    let mut max = 0i64;
    for id in tree.descendants(tree.root()) {
        let Some(name) = tree.name(id) else { continue };
        let local = super::split_qname(name).1;
        if (local == "ins" || local == "del")
            && is_wml_element(tree, id, local)
            && let Some(n) = tree.attr(id, &id_attr).and_then(|v| v.trim().parse().ok())
        {
            max = max.max(n);
        }
    }
    max + 1
}

impl Document {
    /// The document's comments, in part order — the `w:comment` children of the comments
    /// part. Takes `&mut self` because the part is parsed lazily on first access; a
    /// read-only pass leaves every part byte-identical on save. Empty when the document
    /// has no comments part.
    pub fn comments(&mut self) -> Vec<Comment> {
        let Some(part_name) = self.part_by_rel_type(&COMMENTS_REL_TYPES) else {
            return Vec::new();
        };
        let Some(part) = self.ensure_part(&part_name) else {
            return Vec::new();
        };
        let tree = self.tree(part);
        let root = tree.root();
        tree.children(root)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "comment"))
            .map(|c| Comment { part, node: c })
            .collect()
    }
}

/// Ensure the comments part exists and is parsed, creating it on first use (root with the
/// main document root's namespace declarations; no stub content — unlike a notes part,
/// `w:comments` needs none).
fn ensure_comments_part(doc: &mut Document) -> PartId {
    if let Some(existing) = doc.part_by_rel_type(&COMMENTS_REL_TYPES) {
        return doc
            .ensure_part(&existing)
            .expect("existing comments part parses");
    }

    let source = doc.main_part_name().to_string();
    let dir = match source.rfind('/') {
        Some(i) => source[..=i].to_string(),
        None => String::new(),
    };
    let part_name = format!("{dir}comments.xml");

    let root_name = doc.qn(PartId::MAIN, "comments");
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
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <{root_name}{decls}></{root_name}>"
    );

    doc.add_part(part_name.clone(), xml.into_bytes());
    doc.ensure_content_type_override(&format!("/{part_name}"), COMMENTS_CONTENT_TYPE)
        .expect("[Content_Types].xml is editable");
    let target = part_name
        .strip_prefix(&dir)
        .unwrap_or(&part_name)
        .to_string();
    doc.add_relationship(&source, COMMENTS_REL_TYPES[0], &target, false)
        .expect("document relationships part is editable");
    doc.ensure_part(&part_name)
        .expect("created comments part parses")
}
