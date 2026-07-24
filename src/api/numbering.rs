//! Numbering definitions: authoring new independent list definitions in
//! `word/numbering.xml`.
//!
//! [`Document::create_numbering`] appends a fresh `w:abstractNum` (a single level-0
//! definition) plus a `w:num` mapping to it and returns the new `numId`, which
//! [`Paragraph::set_numbering`](crate::Paragraph::set_numbering) then applies to a
//! paragraph. Each call produces an *independent* list: two decimal lists created this way
//! each restart their numbering at 1, because each `w:num` references its own
//! `w:abstractNum` — the point of creating a definition rather than reusing the template's
//! shared "List Number" numbering.
//!
//! When the document has no numbering part at all (e.g. one opened from a file that never
//! defined lists), the part is created from scratch: a minimal `w:numbering` root, an
//! `[Content_Types].xml` `Override` registering its content type, and a `numbering`
//! relationship from the main document part.

use crate::error::{Error, Result};

use super::{Document, PartId, is_wml_element};

/// Relationship type for the numbering part (transitional and strict).
const NUMBERING_REL_TYPES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/numbering",
];
/// Relationship type written when creating a numbering part (the transitional URI Word uses).
const NUMBERING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
/// Content type registered for a freshly created numbering part.
const NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
/// The transitional WordprocessingML main URI, declared on a from-scratch numbering root.
const WML_MAIN_URI: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// A list numbering format — the `w:numFmt` of a level-0 definition created by
/// [`Document::create_numbering`].
///
/// The numeric formats render their level text as `%1.` (the value followed by a period);
/// [`Bullet`](Self::Bullet) renders a `•` glyph in the Symbol font.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberFormat {
    /// `1.`, `2.`, `3.`, …
    Decimal,
    /// `a.`, `b.`, `c.`, …
    LowerLetter,
    /// `A.`, `B.`, `C.`, …
    UpperLetter,
    /// `i.`, `ii.`, `iii.`, …
    LowerRoman,
    /// `I.`, `II.`, `III.`, …
    UpperRoman,
    /// A `•` bullet (rendered in the Symbol font), not a counter.
    Bullet,
}

impl NumberFormat {
    /// The `w:numFmt/@w:val` string for this format.
    fn num_fmt(self) -> &'static str {
        match self {
            NumberFormat::Decimal => "decimal",
            NumberFormat::LowerLetter => "lowerLetter",
            NumberFormat::UpperLetter => "upperLetter",
            NumberFormat::LowerRoman => "lowerRoman",
            NumberFormat::UpperRoman => "upperRoman",
            NumberFormat::Bullet => "bullet",
        }
    }

    /// Whether this format is the bullet (glyph) form rather than a numeric counter.
    fn is_bullet(self) -> bool {
        matches!(self, NumberFormat::Bullet)
    }
}

impl Document {
    /// Create a new, independent list numbering definition and return its `numId`.
    ///
    /// Appends a `w:abstractNum` carrying a single level-0 definition (`w:start` 1, the
    /// requested `w:numFmt`, a `%1.` level text for numeric formats or a `•` bullet for
    /// [`NumberFormat::Bullet`], and a basic hanging indent) plus a `w:num` that maps a new
    /// `numId` to that `w:abstractNum`, into `word/numbering.xml`. Ids are `max existing + 1`
    /// (abstractNumIds from 0, numIds from 1). All `w:abstractNum` elements are kept before
    /// all `w:num` elements, per the `CT_Numbering` content model.
    ///
    /// Because each definition has its own `w:abstractNum`, definitions are independent:
    /// applying two separately created decimal numberings to two groups of paragraphs
    /// numbers each group from 1 — the behavior a restartable list (e.g. a pleading whose
    /// paragraphs count from 1) needs.
    ///
    /// If the document has no numbering part, one is created (part, content-type `Override`,
    /// and `numbering` relationship). Apply the returned id with
    /// [`Paragraph::set_numbering`](crate::Paragraph::set_numbering).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, NumberFormat};
    ///
    /// let mut doc = Document::new();
    /// let a = doc.create_numbering(NumberFormat::Decimal);
    /// let b = doc.create_numbering(NumberFormat::Decimal);
    /// assert_ne!(a, b); // two independent lists
    /// ```
    ///
    /// # Panics
    ///
    /// Panics only if the numbering part cannot be created — which for a document built from
    /// [`Document::new`] or opened from a valid package does not happen.
    pub fn create_numbering(&mut self, format: NumberFormat) -> u32 {
        self.try_create_numbering(format)
            .expect("numbering part is present or can be created")
    }

    /// Fallible core of [`create_numbering`](Self::create_numbering): errors only when the
    /// numbering part must be created and `[Content_Types].xml` or the relationships part
    /// cannot be edited.
    fn try_create_numbering(&mut self, format: NumberFormat) -> Result<u32> {
        let num_part = self.ensure_numbering_part()?;
        let (abstract_id, num_id) = self.next_numbering_ids(num_part);

        let abstract_num = self.build_abstract_num(num_part, abstract_id, format);
        let num = self.build_num(num_part, num_id, abstract_id);

        // Keep all w:abstractNum before all w:num: insert the abstractNum just before the
        // first existing w:num (or at the end when there is none), and append the w:num.
        let tree = self.tree(num_part);
        let root = tree.root();
        let abstract_index = tree
            .children(root)
            .iter()
            .position(|&c| is_wml_element(tree, c, "num"))
            .unwrap_or_else(|| tree.children(root).len());
        let tree = self.tree_mut(num_part);
        tree.insert_child(root, abstract_index, abstract_num);
        tree.append_child(root, num);

        Ok(num_id)
    }

    /// Ensure the numbering part is present and parsed, returning its [`PartId`].
    ///
    /// Uses the existing part if the main document references one via a `numbering`
    /// relationship; otherwise creates it: a minimal `w:numbering` root part, an
    /// `[Content_Types].xml` `Override`, and a `numbering` relationship from the document
    /// part.
    fn ensure_numbering_part(&mut self) -> Result<PartId> {
        if let Some(name) = self.part_by_rel_type(&NUMBERING_REL_TYPES) {
            if let Some(id) = self.ensure_part(&name) {
                return Ok(id);
            }
        }

        // Create the part in the main document part's directory (`word/numbering.xml`).
        let source = self.main_part_name().to_string();
        let dir = match source.rfind('/') {
            Some(i) => source[..=i].to_string(),
            None => String::new(),
        };
        let name = format!("{dir}numbering.xml");
        let minimal = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <w:numbering xmlns:w=\"{WML_MAIN_URI}\"/>"
        );
        self.add_part(name.clone(), minimal.into_bytes());
        self.ensure_content_type_override(&format!("/{name}"), NUMBERING_CONTENT_TYPE)?;
        let target = name.strip_prefix(&dir).unwrap_or(&name).to_string();
        self.add_relationship(&source, NUMBERING_REL_TYPE, &target, false)?;

        self.ensure_part(&name)
            .ok_or_else(|| Error::InvalidPackage("created numbering part does not parse".into()))
    }

    /// The next free `(abstractNumId, numId)` for the numbering part: `max existing + 1`
    /// over the `w:abstractNum/@w:abstractNumId` and `w:num/@w:numId` present. AbstractNumIds
    /// start at 0 and numIds at 1 when the part is empty.
    fn next_numbering_ids(&self, num_part: PartId) -> (u32, u32) {
        let tree = self.tree(num_part);
        let root = tree.root();
        let abstract_attr = self.qn(num_part, "abstractNumId");
        let num_attr = self.qn(num_part, "numId");
        let mut max_abstract: Option<u32> = None;
        let mut max_num: Option<u32> = None;
        for c in tree.children(root).iter().copied() {
            if is_wml_element(tree, c, "abstractNum") {
                if let Some(v) = tree.attr(c, &abstract_attr).and_then(parse_u32) {
                    max_abstract = Some(max_abstract.map_or(v, |m| m.max(v)));
                }
            } else if is_wml_element(tree, c, "num") {
                if let Some(v) = tree.attr(c, &num_attr).and_then(parse_u32) {
                    max_num = Some(max_num.map_or(v, |m| m.max(v)));
                }
            }
        }
        (
            max_abstract.map_or(0, |m| m + 1),
            max_num.map_or(1, |m| m + 1),
        )
    }

    /// Build a detached `w:abstractNum` with a single level-0 definition of `format`.
    fn build_abstract_num(
        &mut self,
        part: PartId,
        abstract_id: u32,
        format: NumberFormat,
    ) -> super::NodeId {
        let abstract_name = self.qn(part, "abstractNum");
        let abstract_id_attr = self.qn(part, "abstractNumId");
        let mlt_name = self.qn(part, "multiLevelType");
        let lvl_name = self.qn(part, "lvl");
        let ilvl_attr = self.qn(part, "ilvl");
        let start_name = self.qn(part, "start");
        let numfmt_name = self.qn(part, "numFmt");
        let lvltext_name = self.qn(part, "lvlText");
        let lvljc_name = self.qn(part, "lvlJc");
        let ppr_name = self.qn(part, "pPr");
        let ind_name = self.qn(part, "ind");
        let rpr_name = self.qn(part, "rPr");
        let rfonts_name = self.qn(part, "rFonts");
        let val_attr = self.qn(part, "val");
        let left_attr = self.qn(part, "left");
        let hanging_attr = self.qn(part, "hanging");
        let ascii_attr = self.qn(part, "ascii");
        let hansi_attr = self.qn(part, "hAnsi");
        let hint_attr = self.qn(part, "hint");

        let lvl_text = if format.is_bullet() {
            "\u{2022}"
        } else {
            "%1."
        };

        let tree = self.tree_mut(part);
        let abstract_num = tree.create_element(abstract_name);
        tree.set_attr(abstract_num, abstract_id_attr, abstract_id.to_string());

        // CT_AbstractNum: multiLevelType precedes the w:lvl entries.
        let mlt = tree.create_element(mlt_name);
        tree.set_attr(mlt, val_attr.clone(), "singleLevel");
        tree.append_child(abstract_num, mlt);

        let lvl = tree.create_element(lvl_name);
        tree.set_attr(lvl, ilvl_attr, "0");

        // CT_Lvl order: start, numFmt, lvlText, lvlJc, pPr, rPr.
        let start = tree.create_element(start_name);
        tree.set_attr(start, val_attr.clone(), "1");
        tree.append_child(lvl, start);

        let numfmt = tree.create_element(numfmt_name);
        tree.set_attr(numfmt, val_attr.clone(), format.num_fmt());
        tree.append_child(lvl, numfmt);

        let lvltext = tree.create_element(lvltext_name);
        tree.set_attr(lvltext, val_attr.clone(), lvl_text);
        tree.append_child(lvl, lvltext);

        let lvljc = tree.create_element(lvljc_name);
        tree.set_attr(lvljc, val_attr.clone(), "left");
        tree.append_child(lvl, lvljc);

        let ppr = tree.create_element(ppr_name);
        let ind = tree.create_element(ind_name);
        tree.set_attr(ind, left_attr, "720");
        tree.set_attr(ind, hanging_attr, "360");
        tree.append_child(ppr, ind);
        tree.append_child(lvl, ppr);

        // Bullets render their glyph in the Symbol font.
        if format.is_bullet() {
            let rpr = tree.create_element(rpr_name);
            let rfonts = tree.create_element(rfonts_name);
            tree.set_attr(rfonts, ascii_attr, "Symbol");
            tree.set_attr(rfonts, hansi_attr, "Symbol");
            tree.set_attr(rfonts, hint_attr, "default");
            tree.append_child(rpr, rfonts);
            tree.append_child(lvl, rpr);
        }

        tree.append_child(abstract_num, lvl);
        abstract_num
    }

    /// Build a detached `w:num` with `num_id` mapping to `abstract_id`.
    fn build_num(&mut self, part: PartId, num_id: u32, abstract_id: u32) -> super::NodeId {
        let num_name = self.qn(part, "num");
        let num_id_attr = self.qn(part, "numId");
        let abstract_ref_name = self.qn(part, "abstractNumId");
        let val_attr = self.qn(part, "val");

        let tree = self.tree_mut(part);
        let num = tree.create_element(num_name);
        tree.set_attr(num, num_id_attr, num_id.to_string());
        let abstract_ref = tree.create_element(abstract_ref_name);
        tree.set_attr(abstract_ref, val_attr, abstract_id.to_string());
        tree.append_child(num, abstract_ref);
        num
    }
}

/// Parse a trimmed decimal `u32`, or `None`.
fn parse_u32(s: &str) -> Option<u32> {
    s.trim().parse::<u32>().ok()
}
