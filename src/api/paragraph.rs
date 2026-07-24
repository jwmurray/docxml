//! The [`Paragraph`] handle.

use crate::xml::{NodeId, XmlTree};

use super::header::{rel_id_attr, rel_id_attr_name};
use super::{
    Alignment, BorderEdge, BorderStyle, Document, FrameAnchor, FrameOptions, FrameWrap, Length,
    LineSpacing, PartId, Pt, RgbColor, Run, TabAlignment, TabLeader, is_wml_element,
    needs_space_preserve, ordered_insert_index, rank_in,
};

/// The relationship type of a hyperlink relationship (the transitional URI Word writes).
/// A hyperlink to an external target is always `TargetMode="External"` with this type.
const HYPERLINK_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/// A hyperlink read out of a paragraph — one `w:hyperlink` element and its resolved link.
///
/// A `w:hyperlink` is either *external* (an `r:id` pointing at a `TargetMode="External"`
/// relationship, e.g. a web address) or *internal* (a `w:anchor` naming a bookmark in the
/// same document); the two are mutually exclusive in practice. [`url`](Self::url) is the
/// relationship target returned verbatim (see [`Paragraph::hyperlinks`]) and is `Some` only
/// for an external link; [`anchor`](Self::anchor) is `Some` only for an internal one.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkInfo {
    /// The external target URL, resolved from the hyperlink's `r:id` through the owning
    /// part's relationships and returned verbatim (not path-resolved). `None` for an
    /// anchor-only (internal) hyperlink or when the `r:id` does not resolve.
    pub url: Option<String>,
    /// The internal bookmark name from `w:anchor`, or `None` for an external hyperlink.
    pub anchor: Option<String>,
    /// The hyperlink's visible text: the concatenated text of the runs inside it.
    pub text: String,
}

/// Canonical `w:pPr` child order (ECMA-376 §17.3.1.26, `CT_PPr` sequence), local names
/// only. New properties are inserted to keep `w:pPr`'s children in this order so the
/// output is schema-valid: `pStyle` comes first, `jc` sits late (after the spacing /
/// indentation group), and `rPr` / `sectPr` come near the end. Unlisted children rank
/// last and stay after authored properties.
const PPR_ORDER: &[&str] = &[
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
];

/// Canonical `w:numPr` child order (ECMA-376 §17.9.24, `CT_NumPr` sequence), local names
/// only: `w:ilvl` (the level) precedes `w:numId` (the numbering reference), which precede
/// the rarely written revision children. New children are inserted to keep this order.
const NUMPR_ORDER: &[&str] = &["ilvl", "numId", "numberingChange", "ins"];

/// Canonical `w:pBdr` child order (ECMA-376 §17.3.1.24, `CT_PBdr` sequence), local names
/// only: the four edges `w:top`, `w:left`, `w:bottom`, `w:right`, then `w:between` and
/// `w:bar`. New edges are inserted to keep this order so any pass-through `w:between`/`w:bar`
/// stays after the authored edges.
const PBDR_ORDER: &[&str] = &["top", "left", "bottom", "right", "between", "bar"];

/// A lightweight handle to a `w:p` paragraph.
///
/// `Paragraph` is `Copy` and borrows nothing — it is just an arena node id (plus the id of
/// the part it lives in) with phantom typing. Pass a [`Document`] back to it (`&Document`
/// to read, `&mut Document` to edit) to do anything useful. A paragraph read from a header
/// or footer carries that part's id, so its text and runs resolve against the header/footer
/// tree, not the body.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Paragraph {
    part: PartId,
    node: NodeId,
}

impl Paragraph {
    /// Wrap a known-`w:p` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Paragraph { part, node }
    }

    /// The paragraph's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The part this paragraph lives in (used by field-code construction).
    pub(crate) fn part(&self) -> PartId {
        self.part
    }

    /// The paragraph's visible text: its runs' text concatenated.
    ///
    /// Within each run, `w:t` contributes its text, `w:tab` a tab, and `w:br` / `w:cr` a
    /// newline — matching python-docx's `Paragraph.text`. Text inside a `w:hyperlink` child
    /// is included, in document order, so a paragraph with a hyperlink reads as the full
    /// sentence (again matching python-docx, which walks hyperlink runs for `.text`). Runs
    /// nested more deeply (inside a `w:smartTag`, `w:ins`, …) are not walked.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("");
    /// p.add_run(&mut doc, "Hello, ");
    /// p.add_run(&mut doc, "world");
    /// assert_eq!(p.text(&doc), "Hello, world");
    /// ```
    pub fn text(&self, doc: &Document) -> String {
        let tree = doc.tree(self.part);
        let mut out = String::new();
        for &child in tree.children(self.node) {
            // Direct runs and the runs inside a direct-child hyperlink both contribute
            // text; append_run_text walks descendant w:t/w:tab/w:br of either.
            if is_wml_element(tree, child, "r") || is_wml_element(tree, child, "hyperlink") {
                append_run_text(tree, child, &mut out);
            }
        }
        out
    }

    /// The paragraph's runs, in order.
    ///
    /// These are the paragraph's *direct* `w:r` children only — matching python-docx's
    /// `Paragraph.runs`. Runs nested inside a `w:hyperlink` are intentionally excluded here
    /// (read them via [`hyperlinks`](Self::hyperlinks)); [`text`](Self::text), by contrast,
    /// does include hyperlink text.
    pub fn runs(&self, doc: &Document) -> Vec<Run> {
        self.run_nodes(doc.tree(self.part))
            .map(|r| Run::from_node(self.part, r))
            .collect()
    }

    /// The paragraph's hyperlinks, in document order — every direct-child `w:hyperlink`.
    ///
    /// For each, [`text`](HyperlinkInfo::text) is the concatenated text of the runs inside
    /// the hyperlink. An external hyperlink's `r:id` is resolved through *this paragraph's
    /// own part's* relationships (so a hyperlink in a header resolves against
    /// `word/_rels/headerN.xml.rels`, a body one against `word/_rels/document.xml.rels`) and
    /// its `TargetMode="External"` target is returned verbatim in
    /// [`url`](HyperlinkInfo::url), not path-resolved. An internal hyperlink's `w:anchor`
    /// (a bookmark name) is reported in [`anchor`](HyperlinkInfo::anchor).
    ///
    /// Takes `&mut Document` because resolution may need to reach the part's relationships;
    /// like the header resolution path it never marks the document modified, so a read-only
    /// call leaves every part byte-identical on save.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("See ");
    /// p.add_hyperlink(&mut doc, "https://example.com/", "the site");
    /// let links = p.hyperlinks(&mut doc);
    /// assert_eq!(links.len(), 1);
    /// assert_eq!(links[0].url.as_deref(), Some("https://example.com/"));
    /// assert_eq!(links[0].text, "the site");
    /// ```
    pub fn hyperlinks(&self, doc: &mut Document) -> Vec<HyperlinkInfo> {
        let part_name = doc.part_name(self.part).to_string();
        let anchor_attr = doc.qn(self.part, "anchor");

        // Gather each hyperlink's r:id, anchor, and text under an immutable borrow first,
        // then resolve r:id -> url (which borrows the document again).
        let collected: Vec<(Option<String>, Option<String>, String)> = {
            let tree = doc.tree(self.part);
            tree.children(self.node)
                .iter()
                .copied()
                .filter(|&c| is_wml_element(tree, c, "hyperlink"))
                .map(|h| {
                    let mut text = String::new();
                    append_run_text(tree, h, &mut text);
                    let r_id = rel_id_attr(tree, h).map(str::to_owned);
                    let anchor = tree.attr(h, &anchor_attr).map(str::to_owned);
                    (r_id, anchor, text)
                })
                .collect()
        };

        collected
            .into_iter()
            .map(|(r_id, anchor, text)| {
                let url = r_id.and_then(|id| doc.rel_target_raw(&part_name, &id));
                HyperlinkInfo { url, anchor, text }
            })
            .collect()
    }

    /// Append an external hyperlink to the paragraph, returning the run carrying its text.
    ///
    /// Adds an `External`-mode relationship (type
    /// `…/officeDocument/2006/relationships/hyperlink`) with `url` as its target to *this
    /// paragraph's part's* relationships — creating that part's rels part if it has none —
    /// and appends `<w:hyperlink r:id="…">` containing one run whose `w:rPr` opens with a
    /// `w:rStyle w:val="Hyperlink"` reference. A hyperlink in a body paragraph therefore
    /// registers its relationship in `word/_rels/document.xml.rels`, one in a header in
    /// `word/_rels/headerN.xml.rels`.
    ///
    /// The `Hyperlink` character style gives the run Word's conventional blue-underlined
    /// look. A document that does not define that style (the blank template built by
    /// [`Document::new`] does not) still produces schema-valid output — the `w:rStyle`
    /// reference to an undefined style is simply ignored on render — so authoring a
    /// hyperlink never requires first creating the style.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("Visit ");
    /// let run = p.add_hyperlink(&mut doc, "https://example.com/", "Example");
    /// assert_eq!(run.text(&doc), "Example");
    /// assert_eq!(p.text(&doc), "Visit Example");
    /// ```
    pub fn add_hyperlink(&self, doc: &mut Document, url: &str, text: &str) -> Run {
        let part_name = doc.part_name(self.part).to_string();
        let r_id = doc
            .add_relationship(&part_name, HYPERLINK_REL_TYPE, url, true)
            .expect("relationships part is editable");
        self.build_hyperlink(doc, Some(&r_id), None, text)
    }

    /// Append an internal (anchor) hyperlink to the paragraph, returning the run carrying
    /// its text.
    ///
    /// The same as [`add_hyperlink`](Self::add_hyperlink) but the `w:hyperlink` carries a
    /// `w:anchor` naming a bookmark in the same document (see [`add_bookmark`](Self::add_bookmark))
    /// instead of an `r:id` — so **no relationship is created**. The run is styled with the
    /// `Hyperlink` character style just as for an external link.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let target = doc.add_paragraph("Signature block");
    /// target.add_bookmark(&mut doc, "sig_block");
    /// let p = doc.add_paragraph("Jump to the ");
    /// p.add_anchor_hyperlink(&mut doc, "sig_block", "signature");
    /// assert_eq!(p.hyperlinks(&mut doc)[0].anchor.as_deref(), Some("sig_block"));
    /// ```
    pub fn add_anchor_hyperlink(&self, doc: &mut Document, anchor: &str, text: &str) -> Run {
        self.build_hyperlink(doc, None, Some(anchor), text)
    }

    /// Append a bookmark anchor point (`w:bookmarkStart` + `w:bookmarkEnd`) to the paragraph.
    ///
    /// The pair is appended around nothing — a zero-length anchor point in python-docx's
    /// style — so it marks a location an [`add_anchor_hyperlink`](Self::add_anchor_hyperlink)
    /// can jump to. Both elements share a fresh `w:id` unique across the part (the max
    /// existing bookmark id plus one), and the start carries the `w:name`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("Here");
    /// p.add_bookmark(&mut doc, "here");
    /// ```
    pub fn add_bookmark(&self, doc: &mut Document, name: &str) {
        let id_attr = doc.qn(self.part, "id");
        let id = next_bookmark_id(doc.tree(self.part), &id_attr);
        let start_name = doc.qn(self.part, "bookmarkStart");
        let end_name = doc.qn(self.part, "bookmarkEnd");
        let name_attr = doc.qn(self.part, "name");
        let id_str = id.to_string();

        let tree = doc.tree_mut(self.part);
        let start = tree.create_element(start_name);
        // CT_Bookmark attribute order: w:id then w:name.
        tree.set_attr(start, id_attr.clone(), id_str.clone());
        tree.set_attr(start, name_attr, name);
        tree.append_child(self.node, start);

        let end = tree.create_element(end_name);
        tree.set_attr(end, id_attr, id_str);
        tree.append_child(self.node, end);
    }

    /// Append a run carrying `text`, returning a handle to it.
    ///
    /// Builds `<w:r><w:t>text</w:t></w:r>`, setting `xml:space="preserve"` on the `w:t`
    /// when `text` has leading or trailing whitespace.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("");
    /// let r = p.add_run(&mut doc, "bold text");
    /// r.bold(&mut doc, true);
    /// assert!(r.is_bold(&doc));
    /// ```
    pub fn add_run(&self, doc: &mut Document, text: &str) -> Run {
        let r_name = doc.qn(self.part, "r");
        let t_name = doc.qn(self.part, "t");

        let tree = doc.tree_mut(self.part);
        let r = tree.create_element(r_name);
        let t = tree.create_element(t_name);
        let content = tree.create_text(text);
        tree.append_child(t, content);
        tree.append_child(r, t);
        tree.append_child(self.node, r);
        if needs_space_preserve(text) {
            tree.set_attr(t, "xml:space", "preserve");
        }
        Run::from_node(self.part, r)
    }

    /// The paragraph's alignment, if `w:pPr/w:jc` is set to a recognized value
    /// (`w:val="distribute"` and other unrecognized values read as `None`).
    pub fn alignment(&self, doc: &Document) -> Option<Alignment> {
        let tree = doc.tree(self.part);
        let jc = self.ppr_child(tree, "jc")?;
        Alignment::from_val(tree.attr(jc, &doc.qn(self.part, "val"))?)
    }

    /// Set the paragraph alignment (`w:pPr/w:jc`).
    pub fn set_alignment(&self, doc: &mut Document, alignment: Alignment) -> Paragraph {
        let ppr = self.ensure_ppr(doc);
        set_alignment_in(doc, self.part, ppr, alignment);
        *self
    }

    /// The paragraph's style id, if `w:pPr/w:pStyle` is set.
    ///
    /// This is the *styleId* — the internal key such as `"Heading1"` — not the human
    /// display name (`"heading 1"`). Resolving a styleId to its display name requires
    /// reading `styles.xml`, which is a later milestone.
    pub fn style_id(&self, doc: &Document) -> Option<String> {
        let tree = doc.tree(self.part);
        let pstyle = self.ppr_child(tree, "pStyle")?;
        tree.attr(pstyle, &doc.qn(self.part, "val"))
            .map(str::to_owned)
    }

    /// The paragraph's style *display name*, resolved through `styles.xml`.
    ///
    /// Where [`style_id`](Self::style_id) returns the internal key (`"Heading1"`), this
    /// resolves that id to the human name Word shows (`"heading 1"`) by reading the style's
    /// `w:name`. Takes `&mut Document` because it parses `styles.xml` through the lazily
    /// cached part; the parse does not mark anything modified, so a read leaves every part
    /// byte-identical on save. `None` when the paragraph has no `w:pStyle`, the style id is
    /// not defined in `styles.xml`, or the style has no `w:name`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let h = doc.add_heading("Chapter 1", 1);
    /// assert_eq!(h.style_name(&mut doc).as_deref(), Some("heading 1"));
    /// ```
    pub fn style_name(&self, doc: &mut Document) -> Option<String> {
        let style_id = self.style_id(doc)?;
        let style = doc.style_by_id(&style_id)?;
        style.display_name(doc)
    }

    /// Set the paragraph's style id (`w:pPr/w:pStyle w:val`).
    ///
    /// `style_id` is the internal styleId (e.g. `"Heading1"`), not the display name; see
    /// [`style_id`](Self::style_id).
    pub fn set_style_id(&self, doc: &mut Document, style_id: &str) -> Paragraph {
        let val = doc.qn(self.part, "val");
        let pstyle = self.ensure_ppr_child(doc, "pStyle");
        doc.tree_mut(self.part).set_attr(pstyle, val, style_id);
        *self
    }

    /// The paragraph's list numbering as `(numId, ilvl)`, or `None` when it is not part of
    /// a list.
    ///
    /// A direct `w:pPr/w:numPr` is read first: its `w:numId/@w:val` is the numbering id and
    /// `w:ilvl/@w:val` the indent level, defaulting to `0` when `w:ilvl` is absent. When the
    /// paragraph carries no direct `w:numPr`, the numbering contributed by its paragraph
    /// *style* is resolved through `styles.xml` (following the `w:basedOn` chain) — this is
    /// how a "List Bullet" / "List Number" paragraph, whose `w:numPr` lives on the style
    /// rather than the paragraph, still reports its numbering, matching how Word renders it.
    /// Returns `None` when neither the paragraph nor its style chain defines a `w:numId`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let num = doc.create_numbering(docxml::NumberFormat::Decimal);
    /// let p = doc.add_paragraph("item");
    /// p.set_numbering(&mut doc, num, 0);
    /// assert_eq!(p.numbering(&doc), Some((num, 0)));
    /// ```
    pub fn numbering(&self, doc: &Document) -> Option<(u32, u32)> {
        let tree = doc.tree(self.part);
        let val = doc.qn(self.part, "val");
        if let Some(numpr) = self.ppr_child(tree, "numPr") {
            if let Some(pair) = read_numpr(tree, numpr, &val) {
                return Some(pair);
            }
        }
        // Fallback: numbering defined on the paragraph style (and its basedOn chain).
        let style_id = self.style_id(doc)?;
        doc.style_numbering(&style_id)
    }

    /// Apply direct list numbering to the paragraph (`w:pPr/w:numPr`).
    ///
    /// Writes `w:ilvl` (the level) then `w:numId` (the numbering reference) — the
    /// `CT_NumPr` child order — creating each if absent, with the `w:numPr` itself placed in
    /// `w:pPr` per the canonical property order. `num_id` is a numbering id defined in
    /// `word/numbering.xml` (see [`Document::create_numbering`]); `level` is the zero-based
    /// list level. Any existing `w:numId` / `w:ilvl` values are overwritten.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let num = doc.create_numbering(docxml::NumberFormat::Bullet);
    /// let p = doc.add_paragraph("bullet");
    /// p.set_numbering(&mut doc, num, 0);
    /// assert_eq!(p.numbering(&doc), Some((num, 0)));
    /// ```
    pub fn set_numbering(&self, doc: &mut Document, num_id: u32, level: u32) -> Paragraph {
        let val = doc.qn(self.part, "val");
        let numpr = self.ensure_ppr_child(doc, "numPr");
        // w:ilvl precedes w:numId in CT_NumPr; create ilvl first so numId slots after it.
        let ilvl = self.ensure_numpr_child(doc, numpr, "ilvl");
        doc.tree_mut(self.part)
            .set_attr(ilvl, val.clone(), level.to_string());
        let numid = self.ensure_numpr_child(doc, numpr, "numId");
        doc.tree_mut(self.part)
            .set_attr(numid, val, num_id.to_string());
        *self
    }

    /// Remove the paragraph's direct list numbering, deleting `w:pPr/w:numPr`.
    ///
    /// Only a direct `w:numPr` is removed; numbering contributed by the paragraph's style is
    /// untouched (clear that by changing the style). After this,
    /// [`numbering`](Self::numbering) falls back to the style, or returns `None`.
    pub fn clear_numbering(&self, doc: &mut Document) {
        if let Some(numpr) = self.ppr_child(doc.tree(self.part), "numPr") {
            doc.tree_mut(self.part).remove_from_parent(numpr);
        }
    }

    /// The paragraph's line spacing (`w:pPr/w:spacing`), if set.
    ///
    /// Reads `w:line` together with `w:lineRule` and maps them to a [`LineSpacing`]
    /// (see that type for the auto-multiple vs. exact/at-least rules). A missing
    /// `w:lineRule` alongside a `w:line` value is read as an auto multiple.
    pub fn line_spacing(&self, doc: &Document) -> Option<LineSpacing> {
        let tree = doc.tree(self.part);
        let sp = self.ppr_child(tree, "spacing")?;
        let line = tree.attr(sp, &doc.qn(self.part, "line"))?;
        let rule = tree.attr(sp, &doc.qn(self.part, "lineRule"));
        LineSpacing::from_line_and_rule(line, rule)
    }

    /// Set the paragraph line spacing (`w:pPr/w:spacing` `w:line` + `w:lineRule`).
    ///
    /// `w:spacing` is a single element shared with [`space_before`](Self::space_before) and
    /// [`space_after`](Self::space_after); this touches only `w:line` and `w:lineRule`,
    /// leaving any `w:before` / `w:after` intact.
    pub fn set_line_spacing(&self, doc: &mut Document, spacing: LineSpacing) -> Paragraph {
        let ppr = self.ensure_ppr(doc);
        set_line_spacing_in(doc, self.part, ppr, spacing);
        *self
    }

    /// The space above the paragraph (`w:pPr/w:spacing/@w:before`), if set.
    pub fn space_before(&self, doc: &Document) -> Option<Pt> {
        self.read_spacing_pt(doc, "before")
    }

    /// The space below the paragraph (`w:pPr/w:spacing/@w:after`), if set.
    pub fn space_after(&self, doc: &Document) -> Option<Pt> {
        self.read_spacing_pt(doc, "after")
    }

    /// Set the space above the paragraph (`w:pPr/w:spacing/@w:before`, in twentieths of a
    /// point). Preserves the other `w:spacing` attributes (`line`, `lineRule`, `after`).
    pub fn set_space_before(&self, doc: &mut Document, space: Pt) -> Paragraph {
        self.set_spacing_pt(doc, "before", space);
        *self
    }

    /// Set the space below the paragraph (`w:pPr/w:spacing/@w:after`, in twentieths of a
    /// point). Preserves the other `w:spacing` attributes (`line`, `lineRule`, `before`).
    pub fn set_space_after(&self, doc: &mut Document, space: Pt) -> Paragraph {
        self.set_spacing_pt(doc, "after", space);
        *self
    }

    /// The paragraph's left indent (`w:pPr/w:ind`), if set. Reads the transitional
    /// `w:left` and the strict `w:start` spelling.
    pub fn left_indent(&self, doc: &Document) -> Option<Length> {
        self.read_ind(doc, &["left", "start"])
    }

    /// The paragraph's right indent (`w:pPr/w:ind`), if set. Reads the transitional
    /// `w:right` and the strict `w:end` spelling.
    pub fn right_indent(&self, doc: &Document) -> Option<Length> {
        self.read_ind(doc, &["right", "end"])
    }

    /// The paragraph's first-line indent (`w:pPr/w:ind`), if set.
    ///
    /// A `w:firstLine` reads as its positive value; a `w:hanging` reads as the negative of
    /// its (positive) value — python-docx's convention, where a hanging indent is a
    /// negative first-line indent. `w:hanging` takes precedence when both are present.
    pub fn first_line_indent(&self, doc: &Document) -> Option<Length> {
        let tree = doc.tree(self.part);
        let ind = self.ppr_child(tree, "ind")?;
        if let Some(hanging) = tree
            .attr(ind, &doc.qn(self.part, "hanging"))
            .and_then(Length::from_twips_str)
        {
            return Some(Length::from_twips(-hanging.twips()));
        }
        tree.attr(ind, &doc.qn(self.part, "firstLine"))
            .and_then(Length::from_twips_str)
    }

    /// Set the left indent (`w:pPr/w:ind/@w:left`), creating `w:ind` if needed.
    pub fn set_left_indent(&self, doc: &mut Document, indent: Length) -> Paragraph {
        self.set_ind(doc, "left", indent);
        *self
    }

    /// Set the right indent (`w:pPr/w:ind/@w:right`), creating `w:ind` if needed.
    pub fn set_right_indent(&self, doc: &mut Document, indent: Length) -> Paragraph {
        self.set_ind(doc, "right", indent);
        *self
    }

    /// Set the first-line indent (`w:pPr/w:ind`), creating `w:ind` if needed.
    ///
    /// A non-negative `indent` writes `w:firstLine`; a negative one writes `w:hanging` with
    /// the magnitude (python-docx's convention). Whichever of the two is not used is
    /// removed, so a paragraph never carries both spellings at once.
    pub fn set_first_line_indent(&self, doc: &mut Document, indent: Length) -> Paragraph {
        let first_attr = doc.qn(self.part, "firstLine");
        let hanging_attr = doc.qn(self.part, "hanging");
        let ind = self.ensure_ppr_child(doc, "ind");
        let twips = indent.twips();
        let tree = doc.tree_mut(self.part);
        if twips < 0 {
            tree.remove_attr(ind, &first_attr);
            tree.set_attr(ind, hanging_attr, (-twips).to_string());
        } else {
            tree.remove_attr(ind, &hanging_attr);
            tree.set_attr(ind, first_attr, twips.to_string());
        }
        *self
    }

    /// Whether the paragraph's lines are kept together on one page (`w:pPr/w:keepLines`).
    pub fn keep_together(&self, doc: &Document) -> bool {
        self.ppr_toggle(doc, "keepLines")
    }

    /// Set whether the paragraph's lines are kept together (`w:pPr/w:keepLines`).
    pub fn set_keep_together(&self, doc: &mut Document, on: bool) -> Paragraph {
        self.set_ppr_toggle(doc, "keepLines", on);
        *self
    }

    /// Whether the paragraph is kept on the same page as the next (`w:pPr/w:keepNext`).
    pub fn keep_with_next(&self, doc: &Document) -> bool {
        self.ppr_toggle(doc, "keepNext")
    }

    /// Set whether the paragraph is kept with the next (`w:pPr/w:keepNext`).
    pub fn set_keep_with_next(&self, doc: &mut Document, on: bool) -> Paragraph {
        self.set_ppr_toggle(doc, "keepNext", on);
        *self
    }

    /// Whether a page break is forced before the paragraph (`w:pPr/w:pageBreakBefore`).
    pub fn page_break_before(&self, doc: &Document) -> bool {
        self.ppr_toggle(doc, "pageBreakBefore")
    }

    /// Set whether a page break is forced before the paragraph (`w:pPr/w:pageBreakBefore`).
    pub fn set_page_break_before(&self, doc: &mut Document, on: bool) -> Paragraph {
        self.set_ppr_toggle(doc, "pageBreakBefore", on);
        *self
    }

    /// Whether line numbering is suppressed for this paragraph (`w:pPr/w:suppressLineNumbers`).
    ///
    /// In a section with [line numbering](crate::Section::set_line_numbering), a paragraph
    /// with this flag is skipped by the numbered-line count — used for captions or block
    /// quotes that should not consume a line number on pleading paper. Same toggle rule as
    /// [`keep_together`](Self::keep_together): present and not `w:val="0"/"false"` is on.
    pub fn suppress_line_numbers(&self, doc: &Document) -> bool {
        self.ppr_toggle(doc, "suppressLineNumbers")
    }

    /// Set whether line numbering is suppressed for this paragraph
    /// (`w:pPr/w:suppressLineNumbers`).
    ///
    /// `w:suppressLineNumbers` sits in `CT_PPr` order after `w:numPr` and before `w:pBdr`
    /// (ECMA-376 §17.3.1.26; present in [`PPR_ORDER`]). On ensures a bare element, off removes
    /// it.
    pub fn set_suppress_line_numbers(&self, doc: &mut Document, on: bool) -> Paragraph {
        self.set_ppr_toggle(doc, "suppressLineNumbers", on);
        *self
    }

    /// This paragraph's frame (`w:pPr/w:framePr`), or `None` when the paragraph is not framed.
    ///
    /// Reads `w:w` / `w:h` (as twips [`Length`]s) into
    /// [`width`](FrameOptions::width) / [`height`](FrameOptions::height), `w:x` / `w:y` into
    /// the offsets, `w:hAnchor` / `w:vAnchor` into the anchors (defaulting to
    /// [`Text`](FrameAnchor::Text) when absent or unrecognized), and `w:wrap` into
    /// [`wrap`](FrameOptions::wrap).
    pub fn frame(&self, doc: &Document) -> Option<FrameOptions> {
        let tree = doc.tree(self.part);
        let fp = self.ppr_child(tree, "framePr")?;
        let read_len = |local: &str| {
            tree.attr(fp, &doc.qn(self.part, local))
                .and_then(Length::from_twips_str)
        };
        let read_anchor = |local: &str| {
            tree.attr(fp, &doc.qn(self.part, local))
                .and_then(FrameAnchor::from_val)
                .unwrap_or(FrameAnchor::Text)
        };
        Some(FrameOptions {
            width: read_len("w"),
            height: read_len("h"),
            x: read_len("x"),
            y: read_len("y"),
            h_anchor: read_anchor("hAnchor"),
            v_anchor: read_anchor("vAnchor"),
            wrap: tree
                .attr(fp, &doc.qn(self.part, "wrap"))
                .and_then(FrameWrap::from_val),
        })
    }

    /// Set this paragraph's frame (`w:pPr/w:framePr`), creating the element in canonical
    /// `w:pPr` order if absent.
    ///
    /// `w:framePr` sits early in `CT_PPr` — after `w:pageBreakBefore` and before
    /// `w:widowControl`/`w:numPr` (ECMA-376 §17.3.1.26; present in [`PPR_ORDER`]). Writes
    /// `w:w` / `w:h` (in twips) for a set width/height — a set height also writes
    /// `w:hRule="atLeast"` — `w:x` / `w:y` for the offsets, always `w:hAnchor` / `w:vAnchor`,
    /// and `w:wrap` when set. The managed attributes are cleared first so a re-set never
    /// leaves a stale one behind.
    pub fn set_frame(&self, doc: &mut Document, frame: FrameOptions) -> Paragraph {
        let w_attr = doc.qn(self.part, "w");
        let h_attr = doc.qn(self.part, "h");
        let hrule_attr = doc.qn(self.part, "hRule");
        let x_attr = doc.qn(self.part, "x");
        let y_attr = doc.qn(self.part, "y");
        let hanchor_attr = doc.qn(self.part, "hAnchor");
        let vanchor_attr = doc.qn(self.part, "vAnchor");
        let wrap_attr = doc.qn(self.part, "wrap");
        let fp = self.ensure_ppr_child(doc, "framePr");
        let tree = doc.tree_mut(self.part);
        for attr in [
            &w_attr,
            &h_attr,
            &hrule_attr,
            &x_attr,
            &y_attr,
            &hanchor_attr,
            &vanchor_attr,
            &wrap_attr,
        ] {
            tree.remove_attr(fp, attr);
        }
        if let Some(width) = frame.width {
            tree.set_attr(fp, w_attr, width.to_twips_string());
        }
        if let Some(height) = frame.height {
            tree.set_attr(fp, h_attr, height.to_twips_string());
            tree.set_attr(fp, hrule_attr, "atLeast");
        }
        if let Some(x) = frame.x {
            tree.set_attr(fp, x_attr, x.to_twips_string());
        }
        if let Some(y) = frame.y {
            tree.set_attr(fp, y_attr, y.to_twips_string());
        }
        tree.set_attr(fp, hanchor_attr, frame.h_anchor.to_val());
        tree.set_attr(fp, vanchor_attr, frame.v_anchor.to_val());
        if let Some(wrap) = frame.wrap {
            tree.set_attr(fp, wrap_attr, wrap.to_val());
        }
        *self
    }

    /// Remove this paragraph's frame, deleting `w:pPr/w:framePr`.
    pub fn clear_frame(&self, doc: &mut Document) {
        if let Some(fp) = self.ppr_child(doc.tree(self.part), "framePr") {
            doc.tree_mut(self.part).remove_from_parent(fp);
        }
    }

    /// This paragraph's borders as `(top, bottom, left, right)` (`w:pPr/w:pBdr`).
    ///
    /// Each edge is `Some` only when the matching `w:pBdr` child is present *and* carries a
    /// modeled [`BorderStyle`]; an edge whose `w:val` is a style this API does not model
    /// reads back as `None` (the enum is closed — see [`BorderStyle`]). `w:color="auto"` (or
    /// an unparsable color) reads as [`color`](BorderEdge::color) `None`.
    pub fn borders(
        &self,
        doc: &Document,
    ) -> (
        Option<BorderEdge>,
        Option<BorderEdge>,
        Option<BorderEdge>,
        Option<BorderEdge>,
    ) {
        let tree = doc.tree(self.part);
        let Some(pbdr) = self.ppr_child(tree, "pBdr") else {
            return (None, None, None, None);
        };
        let read_edge = |local: &str| -> Option<BorderEdge> {
            let el = tree
                .children(pbdr)
                .iter()
                .copied()
                .find(|&c| is_wml_element(tree, c, local))?;
            let style = BorderStyle::from_val(tree.attr(el, &doc.qn(self.part, "val"))?)?;
            let size = tree
                .attr(el, &doc.qn(self.part, "sz"))
                .and_then(|v| v.trim().parse::<u8>().ok())
                .unwrap_or(0);
            let space = tree
                .attr(el, &doc.qn(self.part, "space"))
                .and_then(|v| v.trim().parse::<u8>().ok())
                .unwrap_or(0);
            let color = tree
                .attr(el, &doc.qn(self.part, "color"))
                .and_then(RgbColor::from_hex);
            Some(BorderEdge {
                style,
                size,
                space,
                color,
            })
        };
        (
            read_edge("top"),
            read_edge("bottom"),
            read_edge("left"),
            read_edge("right"),
        )
    }

    /// Set this paragraph's borders (`w:pPr/w:pBdr`).
    ///
    /// Each of `top`, `bottom`, `left`, `right` writes (or replaces) the matching `w:pBdr`
    /// edge; a `None` edge omits it. The edges are (re)built in `CT_PBdr` order — `w:top`,
    /// `w:left`, `w:bottom`, `w:right` (ECMA-376 §17.3.1.24; [`PBDR_ORDER`]) — so any
    /// pass-through `w:between`/`w:bar` stays after them. When all four are `None` the whole
    /// `w:pBdr` is removed (an empty `w:pBdr` would be pointless). `w:pBdr` itself sits in
    /// `CT_PPr` after `w:suppressLineNumbers` and before `w:shd`.
    pub fn set_borders(
        &self,
        doc: &mut Document,
        top: Option<BorderEdge>,
        bottom: Option<BorderEdge>,
        left: Option<BorderEdge>,
        right: Option<BorderEdge>,
    ) -> Paragraph {
        if top.is_none() && bottom.is_none() && left.is_none() && right.is_none() {
            if let Some(pbdr) = self.ppr_child(doc.tree(self.part), "pBdr") {
                doc.tree_mut(self.part).remove_from_parent(pbdr);
            }
            return *self;
        }
        let pbdr = self.ensure_ppr_child(doc, "pBdr");
        // Remove the four managed edges, then rebuild the present ones in schema order.
        for local in ["top", "left", "bottom", "right"] {
            let existing = {
                let tree = doc.tree(self.part);
                tree.children(pbdr)
                    .iter()
                    .copied()
                    .find(|&c| is_wml_element(tree, c, local))
            };
            if let Some(el) = existing {
                doc.tree_mut(self.part).remove_from_parent(el);
            }
        }
        for (local, edge) in [
            ("top", top),
            ("left", left),
            ("bottom", bottom),
            ("right", right),
        ] {
            if let Some(edge) = edge {
                self.insert_border_edge(doc, pbdr, local, edge);
            }
        }
        *self
    }

    /// The paragraph's tab stops (`w:pPr/w:tabs`), in document order.
    ///
    /// Each entry is the stop's position, alignment, and leader. A `w:tab` with an
    /// unrecognized `w:val` defaults to [`TabAlignment::Left`]; a missing or `"none"`
    /// leader is [`TabLeader::None`]; a `w:tab` without a parsable `w:pos` (e.g. a
    /// `w:val="clear"` stop) is skipped.
    pub fn tab_stops(&self, doc: &Document) -> Vec<(Length, TabAlignment, TabLeader)> {
        let tree = doc.tree(self.part);
        let Some(tabs) = self.ppr_child(tree, "tabs") else {
            return Vec::new();
        };
        let pos_attr = doc.qn(self.part, "pos");
        let val_attr = doc.qn(self.part, "val");
        let leader_attr = doc.qn(self.part, "leader");
        tree.children(tabs)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "tab"))
            .filter_map(|c| {
                let pos = tree.attr(c, &pos_attr).and_then(Length::from_twips_str)?;
                let align = tree
                    .attr(c, &val_attr)
                    .and_then(TabAlignment::from_val)
                    .unwrap_or(TabAlignment::Left);
                let leader = tree
                    .attr(c, &leader_attr)
                    .map(TabLeader::from_val)
                    .unwrap_or(TabLeader::None);
                Some((pos, align, leader))
            })
            .collect()
    }

    /// Add a tab stop (`w:pPr/w:tabs/w:tab`), creating `w:tabs` if needed.
    ///
    /// The new `w:tab` carries `w:pos` (in twips), `w:val` (alignment), and `w:leader`
    /// (omitted for [`TabLeader::None`]). Stops are kept sorted ascending by position — the
    /// new one is inserted before the first existing stop that sits farther right — matching
    /// python-docx's position-ordered `TabStops`.
    pub fn add_tab_stop(
        &self,
        doc: &mut Document,
        pos: Length,
        alignment: TabAlignment,
        leader: TabLeader,
    ) {
        let val_attr = doc.qn(self.part, "val");
        let pos_attr = doc.qn(self.part, "pos");
        let leader_attr = doc.qn(self.part, "leader");
        let tab_name = doc.qn(self.part, "tab");
        let tabs = self.ensure_ppr_child(doc, "tabs");

        // Insert before the first existing stop positioned farther right (ascending order).
        let pos_twips = pos.twips();
        let index = {
            let tree = doc.tree(self.part);
            tree.children(tabs)
                .iter()
                .position(|&c| {
                    is_wml_element(tree, c, "tab")
                        && tree
                            .attr(c, &pos_attr)
                            .and_then(Length::from_twips_str)
                            .is_some_and(|l| l.twips() > pos_twips)
                })
                .unwrap_or_else(|| tree.children(tabs).len())
        };

        let tab = doc.tree_mut(self.part).create_element(tab_name);
        // Schema order of CT_TabStop attributes: val, leader, pos.
        doc.tree_mut(self.part)
            .set_attr(tab, val_attr, alignment.to_val());
        if let Some(l) = leader.to_val() {
            doc.tree_mut(self.part).set_attr(tab, leader_attr, l);
        }
        doc.tree_mut(self.part)
            .set_attr(tab, pos_attr, pos.to_twips_string());
        doc.tree_mut(self.part).insert_child(tabs, index, tab);
    }

    /// Remove all tab stops, deleting the `w:pPr/w:tabs` element.
    ///
    /// The whole `w:tabs` is removed rather than left empty: `CT_Tabs` requires at least one
    /// `w:tab`, so an empty `w:tabs` would be schema-invalid. After this,
    /// [`tab_stops`](Self::tab_stops) returns an empty vector.
    pub fn clear_tab_stops(&self, doc: &mut Document) {
        if let Some(tabs) = self.ppr_child(doc.tree(self.part), "tabs") {
            doc.tree_mut(self.part).remove_from_parent(tabs);
        }
    }

    /// Read a twentieths-of-a-point `w:spacing` attribute as [`Pt`].
    fn read_spacing_pt(&self, doc: &Document, attr_local: &str) -> Option<Pt> {
        let tree = doc.tree(self.part);
        let sp = self.ppr_child(tree, "spacing")?;
        Pt::from_twentieths_str(tree.attr(sp, &doc.qn(self.part, attr_local))?)
    }

    /// Set a twentieths-of-a-point `w:spacing` attribute, creating `w:spacing` if needed.
    fn set_spacing_pt(&self, doc: &mut Document, attr_local: &str, space: Pt) {
        let ppr = self.ensure_ppr(doc);
        set_spacing_pt_in(doc, self.part, ppr, attr_local, space);
    }

    /// Read the first present of `attr_locals` on `w:ind` as a [`Length`].
    fn read_ind(&self, doc: &Document, attr_locals: &[&str]) -> Option<Length> {
        let tree = doc.tree(self.part);
        let ind = self.ppr_child(tree, "ind")?;
        for local in attr_locals {
            if let Some(v) = tree.attr(ind, &doc.qn(self.part, local)) {
                return Length::from_twips_str(v);
            }
        }
        None
    }

    /// Set a twips-valued `w:ind` attribute, creating `w:ind` if needed.
    fn set_ind(&self, doc: &mut Document, attr_local: &str, indent: Length) {
        let attr = doc.qn(self.part, attr_local);
        let ind = self.ensure_ppr_child(doc, "ind");
        doc.tree_mut(self.part)
            .set_attr(ind, attr, indent.to_twips_string());
    }

    /// Read an on/off `w:pPr` toggle child (`w:keepLines`, `w:keepNext`, …): present and not
    /// `w:val="0"/"false"` is on.
    fn ppr_toggle(&self, doc: &Document, local: &str) -> bool {
        let tree = doc.tree(self.part);
        let Some(el) = self.ppr_child(tree, local) else {
            return false;
        };
        match tree.attr(el, &doc.qn(self.part, "val")) {
            Some(v) => !matches!(v, "0" | "false"),
            None => true,
        }
    }

    /// Set or clear an on/off `w:pPr` toggle child. On ensures a bare element (clearing an
    /// explicit `w:val="0"/"false"`); off removes it.
    fn set_ppr_toggle(&self, doc: &mut Document, local: &str, on: bool) {
        if on {
            let el = self.ensure_ppr_child(doc, local);
            let val = doc.qn(self.part, "val");
            let tree = doc.tree_mut(self.part);
            if let Some(v) = tree.attr(el, &val) {
                if matches!(v, "0" | "false") {
                    tree.remove_attr(el, &val);
                }
            }
        } else if let Some(el) = self.ppr_child(doc.tree(self.part), local) {
            doc.tree_mut(self.part).remove_from_parent(el);
        }
    }

    /// Build and append a `w:hyperlink` (external `r:id` or internal `w:anchor`) carrying a
    /// single run: `<w:hyperlink …><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>text</w:t></w:r></w:hyperlink>`.
    /// Returns the run handle. Exactly one of `r_id` / `anchor` is `Some` at the call sites.
    fn build_hyperlink(
        &self,
        doc: &mut Document,
        r_id: Option<&str>,
        anchor: Option<&str>,
        text: &str,
    ) -> Run {
        let hlink_name = doc.qn(self.part, "hyperlink");
        let anchor_attr = doc.qn(self.part, "anchor");
        let r_name = doc.qn(self.part, "r");
        let rpr_name = doc.qn(self.part, "rPr");
        let rstyle_name = doc.qn(self.part, "rStyle");
        let t_name = doc.qn(self.part, "t");
        let val_attr = doc.qn(self.part, "val");
        let id_attr = rel_id_attr_name(doc.tree(self.part));
        let preserve = needs_space_preserve(text);

        let tree = doc.tree_mut(self.part);
        let hlink = tree.create_element(hlink_name);
        if let Some(id) = r_id {
            tree.set_attr(hlink, id_attr, id);
        }
        if let Some(a) = anchor {
            tree.set_attr(hlink, anchor_attr, a);
        }

        // The run: rStyle="Hyperlink" is the first (and here only) w:rPr child per RPR_ORDER.
        let r = tree.create_element(r_name);
        let rpr = tree.create_element(rpr_name);
        let rstyle = tree.create_element(rstyle_name);
        tree.set_attr(rstyle, val_attr, "Hyperlink");
        tree.append_child(rpr, rstyle);
        tree.append_child(r, rpr);

        let t = tree.create_element(t_name);
        let content = tree.create_text(text);
        tree.append_child(t, content);
        tree.append_child(r, t);
        if preserve {
            tree.set_attr(t, "xml:space", "preserve");
        }

        tree.append_child(hlink, r);
        tree.append_child(self.node, hlink);
        Run::from_node(self.part, r)
    }

    /// The paragraph's direct `w:r` children as node ids.
    fn run_nodes<'a>(&self, tree: &'a XmlTree) -> impl Iterator<Item = NodeId> + 'a {
        tree.children(self.node)
            .iter()
            .copied()
            .filter(move |&c| is_wml_element(tree, c, "r"))
    }

    /// The paragraph's `w:pPr`, if present.
    fn ppr(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "pPr"))
    }

    /// The paragraph's `w:pPr`, creating it as the first child if absent (schema
    /// requires `w:pPr` before the paragraph content).
    fn ensure_ppr(&self, doc: &mut Document) -> NodeId {
        if let Some(ppr) = self.ppr(doc.tree(self.part)) {
            return ppr;
        }
        let name = doc.qn(self.part, "pPr");
        let tree = doc.tree_mut(self.part);
        let ppr = tree.create_element(name);
        tree.insert_child(self.node, 0, ppr);
        ppr
    }

    /// A direct `w:numPr` child with the given local name, creating it in `CT_NumPr` order
    /// (`w:ilvl` before `w:numId`) if absent.
    fn ensure_numpr_child(&self, doc: &mut Document, numpr: NodeId, local: &str) -> NodeId {
        if let Some(existing) = {
            let tree = doc.tree(self.part);
            tree.children(numpr)
                .iter()
                .copied()
                .find(|&c| is_wml_element(tree, c, local))
        } {
            return existing;
        }
        let name = doc.qn(self.part, local);
        let index = ordered_insert_index(
            doc.tree(self.part),
            numpr,
            rank_in(NUMPR_ORDER, local),
            NUMPR_ORDER,
        );
        let el = doc.tree_mut(self.part).create_element(name);
        doc.tree_mut(self.part).insert_child(numpr, index, el);
        el
    }

    /// Create a `w:pBdr` edge element (`w:top`/`w:left`/`w:bottom`/`w:right`) carrying the
    /// edge's `w:val`, `w:sz`, `w:space`, and `w:color`, and insert it into `pbdr` in
    /// `CT_PBdr` order. A `None` color writes `w:color="auto"`.
    fn insert_border_edge(&self, doc: &mut Document, pbdr: NodeId, local: &str, edge: BorderEdge) {
        let val_attr = doc.qn(self.part, "val");
        let sz_attr = doc.qn(self.part, "sz");
        let space_attr = doc.qn(self.part, "space");
        let color_attr = doc.qn(self.part, "color");
        let name = doc.qn(self.part, local);
        let index = ordered_insert_index(
            doc.tree(self.part),
            pbdr,
            rank_in(PBDR_ORDER, local),
            PBDR_ORDER,
        );
        let tree = doc.tree_mut(self.part);
        let el = tree.create_element(name);
        // CT_Border attribute order: val, then sz, space, and color.
        tree.set_attr(el, val_attr, edge.style.to_val());
        tree.set_attr(el, sz_attr, edge.size.to_string());
        tree.set_attr(el, space_attr, edge.space.to_string());
        match edge.color {
            Some(color) => tree.set_attr(el, color_attr, color.to_hex()),
            None => tree.set_attr(el, color_attr, "auto"),
        }
        tree.insert_child(pbdr, index, el);
    }

    /// A direct `w:pPr` child with the given WML local name, if present.
    fn ppr_child(&self, tree: &XmlTree, local: &str) -> Option<NodeId> {
        let ppr = self.ppr(tree)?;
        tree.children(ppr)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, local))
    }

    /// A direct `w:pPr` child with the given local name, creating it (in canonical
    /// schema order) if absent. Creates `w:pPr` first if needed. Thin wrapper over the
    /// part-agnostic [`ensure_ppr_child_in`] that ensures the paragraph's own `w:pPr`.
    fn ensure_ppr_child(&self, doc: &mut Document, local: &str) -> NodeId {
        let ppr = self.ensure_ppr(doc);
        ensure_ppr_child_in(doc, self.part, ppr, local)
    }
}

// --- Reusable `w:pPr` property writers ---------------------------------------------
//
// Like the `w:rPr` writers in [`run`](super::run), these operate on any `(part, w:pPr)`
// pair so both [`Paragraph`] and a paragraph-type [`Style`](super::Style) share one
// implementation. Only the placement of the `w:pPr` within its parent differs (a paragraph
// puts it first; a style slots it in `CT_Style` order), so each caller ensures its own
// `w:pPr` and passes the node id here.

/// A direct `w:pPr` child element with WML local name `local`, if present.
pub(super) fn ppr_child_in(tree: &XmlTree, ppr: NodeId, local: &str) -> Option<NodeId> {
    tree.children(ppr)
        .iter()
        .copied()
        .find(|&c| is_wml_element(tree, c, local))
}

/// Ensure a direct `w:pPr` child `local` exists, inserted in canonical `CT_PPr` order
/// ([`PPR_ORDER`]); return it. `ppr` is a `w:pPr` element living in `part`.
pub(super) fn ensure_ppr_child_in(
    doc: &mut Document,
    part: PartId,
    ppr: NodeId,
    local: &str,
) -> NodeId {
    if let Some(existing) = ppr_child_in(doc.tree(part), ppr, local) {
        return existing;
    }
    let name = doc.qn(part, local);
    let index = ordered_insert_index(doc.tree(part), ppr, rank_in(PPR_ORDER, local), PPR_ORDER);
    let el = doc.tree_mut(part).create_element(name);
    doc.tree_mut(part).insert_child(ppr, index, el);
    el
}

/// Set the alignment on `ppr` (`w:jc`).
pub(super) fn set_alignment_in(
    doc: &mut Document,
    part: PartId,
    ppr: NodeId,
    alignment: Alignment,
) {
    let val = doc.qn(part, "val");
    let jc = ensure_ppr_child_in(doc, part, ppr, "jc");
    doc.tree_mut(part).set_attr(jc, val, alignment.to_val());
}

/// Set a twentieths-of-a-point `w:spacing` attribute (`before`/`after`) on `ppr`, leaving
/// the other spacing attributes intact.
pub(super) fn set_spacing_pt_in(
    doc: &mut Document,
    part: PartId,
    ppr: NodeId,
    attr_local: &str,
    space: Pt,
) {
    let attr = doc.qn(part, attr_local);
    let sp = ensure_ppr_child_in(doc, part, ppr, "spacing");
    doc.tree_mut(part)
        .set_attr(sp, attr, space.to_twentieths_string());
}

/// Set the line spacing on `ppr` (`w:spacing` `w:line` + `w:lineRule`), leaving any
/// `w:before`/`w:after` intact.
pub(super) fn set_line_spacing_in(
    doc: &mut Document,
    part: PartId,
    ppr: NodeId,
    spacing: LineSpacing,
) {
    let (line, rule) = spacing.to_line_and_rule();
    let line_attr = doc.qn(part, "line");
    let rule_attr = doc.qn(part, "lineRule");
    let sp = ensure_ppr_child_in(doc, part, ppr, "spacing");
    let tree = doc.tree_mut(part);
    tree.set_attr(sp, line_attr, line);
    tree.set_attr(sp, rule_attr, rule);
}

/// Read `(numId, ilvl)` out of a `w:numPr` element (`val_name` is the part's `w:val`
/// qualified attribute name). `None` when no parsable `w:numId` child is present; `w:ilvl`
/// defaults to `0` when absent or unparsable.
pub(super) fn read_numpr(tree: &XmlTree, numpr: NodeId, val_name: &str) -> Option<(u32, u32)> {
    let numid = tree
        .children(numpr)
        .iter()
        .copied()
        .find(|&c| is_wml_element(tree, c, "numId"))
        .and_then(|c| tree.attr(c, val_name))
        .and_then(|v| v.trim().parse::<u32>().ok())?;
    let ilvl = tree
        .children(numpr)
        .iter()
        .copied()
        .find(|&c| is_wml_element(tree, c, "ilvl"))
        .and_then(|c| tree.attr(c, val_name))
        .and_then(|v| v.trim().parse::<u32>().ok())
        .unwrap_or(0);
    Some((numid, ilvl))
}

/// The next free bookmark `w:id` for a part: the maximum `w:id` on any `w:bookmarkStart`
/// anywhere in the part, plus one (`0` when the part has no bookmarks). `id_attr` is the
/// part's qualified `w:id` attribute name. Scanning the whole tree — not just this
/// paragraph — keeps ids unique across the part, as the schema requires.
fn next_bookmark_id(tree: &XmlTree, id_attr: &str) -> u32 {
    let mut max: i64 = -1;
    for node in tree.descendants(tree.root()) {
        if is_wml_element(tree, node, "bookmarkStart") {
            if let Some(v) = tree
                .attr(node, id_attr)
                .and_then(|s| s.trim().parse::<i64>().ok())
            {
                max = max.max(v);
            }
        }
    }
    (max + 1) as u32
}

/// Append a run's text (python-docx semantics) to `out`: `w:t` text verbatim, `w:tab`
/// as a tab, `w:br` / `w:cr` as a newline.
pub(super) fn append_run_text(tree: &XmlTree, run: NodeId, out: &mut String) {
    for descendant in tree.descendants(run) {
        if is_wml_element(tree, descendant, "t") {
            out.push_str(&tree.text_content(descendant));
        } else if is_wml_element(tree, descendant, "tab") {
            out.push('\t');
        } else if is_wml_element(tree, descendant, "br") || is_wml_element(tree, descendant, "cr") {
            out.push('\n');
        }
    }
}
