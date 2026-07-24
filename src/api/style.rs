//! The style catalog and style authoring: reading and writing `word/styles.xml`.
//!
//! [`Style`] is a `Copy` handle (the [`PartId`] of the styles part plus the [`NodeId`] of a
//! `w:style` element), following the crate's handle pattern. [`Document::styles`],
//! [`style_by_id`](Document::style_by_id), and [`style_by_name`](Document::style_by_name)
//! enumerate and look styles up; [`Document::create_style`] appends new ones; the [`Style`]
//! setters write a style's `w:basedOn`/`w:next` and its character (`w:rPr`) and paragraph
//! (`w:pPr`) formatting, reusing the same ordered-insert machinery [`Run`](crate::Run) and
//! [`Paragraph`](crate::Paragraph) use for direct formatting.
//!
//! # Two readers for `styles.xml`
//!
//! This module parses `styles.xml` through the lazily *cached* per-part tree
//! ([`Document::ensure_part`]), because its handle API and effective-formatting reads take
//! `&mut Document`. The older `&self` read helpers
//! ([`Document::style_numbering`](Document::style_numbering), reached from
//! [`Paragraph::numbering`](crate::Paragraph::numbering)) still parse `styles.xml` from its
//! raw bytes on demand, since those accessors are `&self` and cannot ensure a cached parse.
//! The two paths agree on content; the raw-bytes path is only kept where the read signature
//! predates this milestone (migrating it would force `numbering()` to `&mut self` and break
//! the existing read API).

use crate::error::{Error, Result};
use crate::xml::{NodeId, XmlTree};

use super::paragraph::{set_alignment_in, set_line_spacing_in, set_spacing_pt_in};
use super::run::{
    remove_child_in, rpr_child_in, set_color_in, set_font_in, set_size_in, toggle_on_in,
};
use super::{
    Alignment, Document, LineSpacing, PartId, Pt, RgbColor, is_wml_element, ordered_insert_index,
    rank_in,
};

/// Relationship types identifying the styles part (transitional and strict).
const STYLES_REL_TYPES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/styles",
];
/// Relationship type written when creating a styles part (the transitional URI Word uses).
const STYLES_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
/// Content type registered for a freshly created styles part.
const STYLES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/// Canonical `w:style` child order (ECMA-376 §17.7.4.17, `CT_Style` sequence), local names
/// only. New children (`w:basedOn`, `w:next`, `w:pPr`, `w:rPr`, …) are inserted to keep this
/// order so the output is schema-valid: `w:name` leads, `w:basedOn`/`w:next` precede
/// `w:qFormat`, and `w:pPr` precedes `w:rPr` near the end. Unlisted children rank last.
const STYLE_ORDER: &[&str] = &[
    "name",
    "aliases",
    "basedOn",
    "next",
    "link",
    "autoRedefine",
    "hidden",
    "uiPriority",
    "semiHidden",
    "unhideWhenUsed",
    "qFormat",
    "locked",
    "personal",
    "personalCompose",
    "personalReply",
    "rsid",
    "pPr",
    "rPr",
    "tblPr",
    "trPr",
    "tcPr",
    "tblStylePr",
];

/// The kind of a style, mirroring the `w:type` attribute of a `w:style` element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleType {
    /// A paragraph style (`w:type="paragraph"`).
    Paragraph,
    /// A character style (`w:type="character"`).
    Character,
    /// A table style (`w:type="table"`).
    Table,
    /// A numbering style (`w:type="numbering"`).
    Numbering,
}

impl StyleType {
    /// The `w:type` attribute value for this style kind.
    fn to_val(self) -> &'static str {
        match self {
            StyleType::Paragraph => "paragraph",
            StyleType::Character => "character",
            StyleType::Table => "table",
            StyleType::Numbering => "numbering",
        }
    }

    /// Parse a `w:type` attribute value, or `None` for an unrecognized one.
    fn from_val(val: &str) -> Option<StyleType> {
        match val {
            "paragraph" => Some(StyleType::Paragraph),
            "character" => Some(StyleType::Character),
            "table" => Some(StyleType::Table),
            "numbering" => Some(StyleType::Numbering),
            _ => None,
        }
    }
}

/// A lightweight handle to a `w:style` element in `word/styles.xml`.
///
/// Like [`Paragraph`](crate::Paragraph) and [`Run`](crate::Run), `Style` is `Copy` and
/// borrows nothing — it is the [`PartId`] of the styles part plus the [`NodeId`] of the
/// `w:style` element. Obtain one from [`Document::styles`],
/// [`Document::style_by_id`](Document::style_by_id),
/// [`Document::style_by_name`](Document::style_by_name), or
/// [`Document::create_style`](Document::create_style); pass a [`Document`] back to read or
/// edit it. The setters return the style so calls chain.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Style {
    part: PartId,
    node: NodeId,
}

impl Style {
    /// Wrap a known-`w:style` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Style { part, node }
    }

    /// The style's underlying tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The style's internal id (`w:style/@w:styleId`), e.g. `"Heading1"`. Empty when the
    /// element somehow carries no `w:styleId` (not valid WordprocessingML, but tolerated).
    pub fn style_id(&self, doc: &Document) -> String {
        let tree = doc.tree(self.part);
        tree.attr(self.node, &doc.qn(self.part, "styleId"))
            .unwrap_or_default()
            .to_string()
    }

    /// The style's display name (`w:style/w:name/@w:val`), e.g. `"heading 1"` — the human
    /// name Word shows in its styles pane, as opposed to the internal
    /// [`style_id`](Self::style_id). `None` when the style has no `w:name`.
    pub fn display_name(&self, doc: &Document) -> Option<String> {
        let tree = doc.tree(self.part);
        let name = style_child(tree, self.node, "name")?;
        tree.attr(name, &doc.qn(self.part, "val"))
            .map(str::to_owned)
    }

    /// The style's kind (`w:style/@w:type`), defaulting to [`StyleType::Paragraph`] when the
    /// attribute is absent or unrecognized (the schema default for `w:type`).
    pub fn style_type(&self, doc: &Document) -> StyleType {
        let tree = doc.tree(self.part);
        tree.attr(self.node, &doc.qn(self.part, "type"))
            .and_then(StyleType::from_val)
            .unwrap_or(StyleType::Paragraph)
    }

    /// The id of the style this one is based on (`w:style/w:basedOn/@w:val`), or `None` when
    /// the style has no `w:basedOn`. This is the parent in the inheritance chain that
    /// effective-formatting reads walk.
    pub fn based_on(&self, doc: &Document) -> Option<String> {
        let tree = doc.tree(self.part);
        let b = style_child(tree, self.node, "basedOn")?;
        tree.attr(b, &doc.qn(self.part, "val")).map(str::to_owned)
    }

    /// Set the style this one is based on (`w:style/w:basedOn/@w:val`), creating `w:basedOn`
    /// in canonical `CT_Style` order if absent.
    pub fn set_based_on(&self, doc: &mut Document, id: &str) -> Style {
        self.set_child_val(doc, "basedOn", id);
        *self
    }

    /// Set the style applied to the *next* paragraph after one in this style
    /// (`w:style/w:next/@w:val`) — e.g. a heading whose next paragraph reverts to `Normal`.
    pub fn set_next(&self, doc: &mut Document, id: &str) -> Style {
        self.set_child_val(doc, "next", id);
        *self
    }

    /// Turn bold on or off in the style's `w:rPr` (`w:b`). Same representation as
    /// [`Run::bold`](crate::Run::bold): on adds a bare `w:b`, off removes it.
    pub fn set_bold(&self, doc: &mut Document, on: bool) -> Style {
        self.set_toggle(doc, "b", on);
        *self
    }

    /// Turn italic on or off in the style's `w:rPr` (`w:i`).
    pub fn set_italic(&self, doc: &mut Document, on: bool) -> Style {
        self.set_toggle(doc, "i", on);
        *self
    }

    /// Set the font size in the style's `w:rPr` (`w:sz` + `w:szCs`, half-points).
    pub fn set_size(&self, doc: &mut Document, size: Pt) -> Style {
        let rpr = self.ensure_rpr(doc);
        set_size_in(doc, self.part, rpr, size);
        *self
    }

    /// Set the color in the style's `w:rPr` (`w:color w:val`).
    pub fn set_color(&self, doc: &mut Document, color: RgbColor) -> Style {
        let rpr = self.ensure_rpr(doc);
        set_color_in(doc, self.part, rpr, color);
        *self
    }

    /// Set the font (typeface) in the style's `w:rPr` (`w:rFonts` `w:ascii`/`w:hAnsi`).
    pub fn set_font(&self, doc: &mut Document, name: &str) -> Style {
        let rpr = self.ensure_rpr(doc);
        set_font_in(doc, self.part, rpr, name);
        *self
    }

    /// Set the paragraph alignment in the style's `w:pPr` (`w:jc`). Meaningful for
    /// paragraph styles ([`StyleType::Paragraph`]).
    pub fn set_alignment(&self, doc: &mut Document, alignment: Alignment) -> Style {
        let ppr = self.ensure_ppr(doc);
        set_alignment_in(doc, self.part, ppr, alignment);
        *self
    }

    /// Set the space above paragraphs in this style (`w:pPr/w:spacing/@w:before`).
    pub fn set_space_before(&self, doc: &mut Document, space: Pt) -> Style {
        let ppr = self.ensure_ppr(doc);
        set_spacing_pt_in(doc, self.part, ppr, "before", space);
        *self
    }

    /// Set the space below paragraphs in this style (`w:pPr/w:spacing/@w:after`).
    pub fn set_space_after(&self, doc: &mut Document, space: Pt) -> Style {
        let ppr = self.ensure_ppr(doc);
        set_spacing_pt_in(doc, self.part, ppr, "after", space);
        *self
    }

    /// Set the line spacing in the style's `w:pPr` (`w:spacing` `w:line` + `w:lineRule`).
    pub fn set_line_spacing(&self, doc: &mut Document, spacing: LineSpacing) -> Style {
        let ppr = self.ensure_ppr(doc);
        set_line_spacing_in(doc, self.part, ppr, spacing);
        *self
    }

    /// Set a single-valued child element (`w:basedOn`, `w:next`) to `id`, creating it in
    /// `CT_Style` order if absent.
    fn set_child_val(&self, doc: &mut Document, local: &str, id: &str) {
        let val = doc.qn(self.part, "val");
        let el = self.ensure_style_child(doc, local);
        doc.tree_mut(self.part).set_attr(el, val, id);
    }

    /// Set or clear a boolean toggle in the style's `w:rPr` (mirrors
    /// [`Run`](crate::Run)'s toggle behavior): on ensures a bare element, off removes it —
    /// without leaving an empty `w:rPr` behind when the property was never set.
    fn set_toggle(&self, doc: &mut Document, local: &str, on: bool) {
        if on {
            let rpr = self.ensure_rpr(doc);
            toggle_on_in(doc, self.part, rpr, local);
        } else if let Some(rpr) = style_child(doc.tree(self.part), self.node, "rPr") {
            remove_child_in(doc, self.part, rpr, local);
        }
    }

    /// The style's `w:rPr`, creating it in canonical `CT_Style` order if absent.
    fn ensure_rpr(&self, doc: &mut Document) -> NodeId {
        self.ensure_style_child(doc, "rPr")
    }

    /// The style's `w:pPr`, creating it in canonical `CT_Style` order if absent.
    fn ensure_ppr(&self, doc: &mut Document) -> NodeId {
        self.ensure_style_child(doc, "pPr")
    }

    /// A direct child of the `w:style` with local name `local`, creating it in canonical
    /// `CT_Style` order ([`STYLE_ORDER`]) if absent.
    fn ensure_style_child(&self, doc: &mut Document, local: &str) -> NodeId {
        if let Some(existing) = style_child(doc.tree(self.part), self.node, local) {
            return existing;
        }
        let name = doc.qn(self.part, local);
        let index = ordered_insert_index(
            doc.tree(self.part),
            self.node,
            rank_in(STYLE_ORDER, local),
            STYLE_ORDER,
        );
        let el = doc.tree_mut(self.part).create_element(name);
        doc.tree_mut(self.part).insert_child(self.node, index, el);
        el
    }
}

impl Document {
    /// Every `w:style` defined in `word/styles.xml`, in document order.
    ///
    /// Parses `styles.xml` through the lazily cached part machinery
    /// ([`ensure_part`](Self::ensure_part)); the parse does not mark anything modified, so a
    /// read leaves every part byte-identical on save. Returns an empty vector when the
    /// document has no styles part. Latent style definitions (`w:latentStyles`) are not
    /// `w:style` elements and are excluded, matching python-docx's `styles`.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// assert!(!doc.styles().is_empty()); // the blank template ships built-in styles
    /// ```
    pub fn styles(&mut self) -> Vec<Style> {
        let Some(part) = self.styles_part() else {
            return Vec::new();
        };
        let tree = self.tree(part);
        let root = tree.root();
        tree.children(root)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "style"))
            .map(|c| Style::from_node(part, c))
            .collect()
    }

    /// The style with internal id `id` (`w:styleId`), or `None` when no style has it.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let normal = doc.style_by_id("Normal").expect("template defines Normal");
    /// assert_eq!(normal.style_id(&doc), "Normal");
    /// ```
    pub fn style_by_id(&mut self, id: &str) -> Option<Style> {
        let styles = self.styles();
        styles.into_iter().find(|s| s.style_id(self) == id)
    }

    /// The style whose *display name* (`w:name`) is `name`, preferring a case-sensitive exact
    /// match and falling back to a case-insensitive one.
    ///
    /// python-docx looks styles up by display name (e.g. `"Heading 1"`), so this mirrors that
    /// affordance. The case-insensitive fallback tolerates the builtin-name casing quirks
    /// (Word stores `"heading 1"` lowercase). Returns `None` when no style's name matches
    /// either way.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let h1 = doc.style_by_name("heading 1").expect("template defines heading 1");
    /// assert_eq!(h1.style_id(&doc), "Heading1");
    /// ```
    pub fn style_by_name(&mut self, name: &str) -> Option<Style> {
        let styles = self.styles();
        if let Some(s) = styles
            .iter()
            .find(|s| s.display_name(self).as_deref() == Some(name))
        {
            return Some(*s);
        }
        styles.into_iter().find(|s| {
            s.display_name(self)
                .is_some_and(|n| n.eq_ignore_ascii_case(name))
        })
    }

    /// Create a new style and return a handle to it; idempotent by id.
    ///
    /// Appends a `w:style` carrying `w:type`, `w:styleId`, a `w:name` (the display name), and
    /// a `w:qFormat` (so the style shows in Word's gallery) to `word/styles.xml`. When a style
    /// with `id` already exists it is returned unchanged (the append is skipped), so calling
    /// this twice is safe. Set formatting and inheritance with the [`Style`] setters
    /// afterwards; apply the style with
    /// [`Paragraph::set_style_id`](crate::Paragraph::set_style_id) or
    /// [`Run::set_style_id`](crate::Run::set_style_id).
    ///
    /// If the document has no styles part, one is created (a minimal `w:styles` root with the
    /// namespace declarations copied from the main document root, an `[Content_Types].xml`
    /// `Override`, and a `styles` relationship), mirroring numbering-/settings-part creation.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, StyleType, Pt, Alignment};
    ///
    /// let mut doc = Document::new();
    /// let s = doc.create_style("FirmTitle", "Firm Title", StyleType::Paragraph);
    /// s.set_based_on(&mut doc, "Normal")
    ///     .set_bold(&mut doc, true)
    ///     .set_size(&mut doc, Pt(14.0))
    ///     .set_alignment(&mut doc, Alignment::Center);
    /// assert_eq!(doc.create_style("FirmTitle", "x", StyleType::Paragraph).style_id(&doc), "FirmTitle");
    /// ```
    ///
    /// # Panics
    ///
    /// Panics only if the styles part must be created and `[Content_Types].xml` or the
    /// relationships part cannot be edited — which for a [`Document::new`] or a valid opened
    /// package does not happen.
    pub fn create_style(&mut self, id: &str, display_name: &str, style_type: StyleType) -> Style {
        self.try_create_style(id, display_name, style_type)
            .expect("styles part is present or can be created")
    }

    /// Set the document's default font (`w:styles/w:docDefaults/w:rPrDefault/w:rPr`), the
    /// lowest-priority formatting every run inherits.
    ///
    /// Writes `w:rFonts` `w:ascii`/`w:hAnsi` and `w:sz`/`w:szCs`, creating the
    /// `w:docDefaults` → `w:rPrDefault` → `w:rPr` chain when absent — `w:docDefaults` is the
    /// **first** child of `w:styles`, per `CT_Styles`.
    ///
    /// Setting explicit `w:ascii`/`w:hAnsi` here is the production "defeat theme fonts"
    /// pattern: Word's default document uses *theme* fonts (`w:asciiTheme="minorHAnsi"`),
    /// which follow whatever theme the document (or a receiving machine's Normal.dotm)
    /// carries; writing concrete font names into the run defaults pins the typeface so it
    /// renders identically everywhere. The existing theme attributes are left in place — the
    /// explicit `w:ascii`/`w:hAnsi` override them.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, Pt};
    ///
    /// let mut doc = Document::new();
    /// doc.set_default_font("Century Schoolbook", Pt(13.0));
    /// ```
    ///
    /// # Panics
    ///
    /// Panics only if the styles part must be created and cannot be (see
    /// [`create_style`](Self::create_style)).
    pub fn set_default_font(&mut self, name: &str, size: Pt) {
        self.try_set_default_font(name, size)
            .expect("styles part is present or can be created");
    }

    /// Ensure the styles part is present and parsed, returning its [`PartId`]. Read-only in
    /// spirit: parsing does not mark the part modified.
    pub(crate) fn styles_part(&mut self) -> Option<PartId> {
        let name = self.part_by_rel_type(&STYLES_REL_TYPES)?;
        self.ensure_part(&name)
    }

    /// Fallible core of [`create_style`](Self::create_style).
    fn try_create_style(
        &mut self,
        id: &str,
        display_name: &str,
        style_type: StyleType,
    ) -> Result<Style> {
        if let Some(existing) = self.style_by_id(id) {
            return Ok(existing);
        }
        let part = self.ensure_styles_part()?;

        let style_name = self.qn(part, "style");
        let type_attr = self.qn(part, "type");
        let styleid_attr = self.qn(part, "styleId");
        let name_name = self.qn(part, "name");
        let qformat_name = self.qn(part, "qFormat");
        let val_attr = self.qn(part, "val");

        let root = self.tree(part).root();
        let tree = self.tree_mut(part);
        let style = tree.create_element(style_name);
        // CT_Style attribute order: w:type then w:styleId.
        tree.set_attr(style, type_attr, style_type.to_val());
        tree.set_attr(style, styleid_attr, id);

        // CT_Style child order: w:name leads, w:qFormat later (basedOn/next slot between,
        // via the ordered-insert setters).
        let name_el = tree.create_element(name_name);
        tree.set_attr(name_el, val_attr, display_name);
        tree.append_child(style, name_el);
        let qformat = tree.create_element(qformat_name);
        tree.append_child(style, qformat);

        tree.append_child(root, style);
        Ok(Style::from_node(part, style))
    }

    /// Fallible core of [`set_default_font`](Self::set_default_font).
    fn try_set_default_font(&mut self, name: &str, size: Pt) -> Result<()> {
        let part = self.ensure_styles_part()?;
        let rpr = self.ensure_docdefaults_rpr(part);
        let ascii = self.qn(part, "ascii");
        let hansi = self.qn(part, "hAnsi");
        let rfonts = super::run::ensure_rpr_child_in(self, part, rpr, "rFonts");
        self.tree_mut(part).set_attr(rfonts, ascii, name);
        self.tree_mut(part).set_attr(rfonts, hansi, name);
        set_size_in(self, part, rpr, size);
        Ok(())
    }

    /// Ensure `w:styles/w:docDefaults/w:rPrDefault/w:rPr` exists, creating the chain if
    /// absent (`w:docDefaults` inserted as the first child of `w:styles`), and return the
    /// `w:rPr`.
    fn ensure_docdefaults_rpr(&mut self, part: PartId) -> NodeId {
        let root = self.tree(part).root();
        let docdefaults = match style_child(self.tree(part), root, "docDefaults") {
            Some(d) => d,
            None => {
                let name = self.qn(part, "docDefaults");
                let el = self.tree_mut(part).create_element(name);
                self.tree_mut(part).insert_child(root, 0, el);
                el
            }
        };
        let rpr_default = match style_child(self.tree(part), docdefaults, "rPrDefault") {
            Some(r) => r,
            None => {
                let name = self.qn(part, "rPrDefault");
                let el = self.tree_mut(part).create_element(name);
                // rPrDefault precedes pPrDefault in CT_DocDefaults.
                self.tree_mut(part).insert_child(docdefaults, 0, el);
                el
            }
        };
        match style_child(self.tree(part), rpr_default, "rPr") {
            Some(r) => r,
            None => {
                let name = self.qn(part, "rPr");
                let el = self.tree_mut(part).create_element(name);
                self.tree_mut(part).append_child(rpr_default, el);
                el
            }
        }
    }

    /// Ensure the styles part is present and parsed, creating it from scratch if the document
    /// has none (mirrors numbering-/settings-part creation).
    fn ensure_styles_part(&mut self) -> Result<PartId> {
        if let Some(name) = self.part_by_rel_type(&STYLES_REL_TYPES) {
            if let Some(id) = self.ensure_part(&name) {
                return Ok(id);
            }
        }

        // Copy the namespace declarations from the main document root so the new part binds
        // the same prefixes (in particular the WordprocessingML main prefix).
        let main = PartId::MAIN;
        let decls = {
            let tree = self.tree(main);
            let root = tree.root();
            let mut s = String::new();
            for (k, v) in tree.attrs(root) {
                if k == "xmlns" || k.starts_with("xmlns:") {
                    s.push(' ');
                    s.push_str(k);
                    s.push_str("=\"");
                    s.push_str(v);
                    s.push('"');
                }
            }
            s
        };
        let root_name = self.qn(main, "styles");

        let source = self.main_part_name().to_string();
        let dir = match source.rfind('/') {
            Some(i) => source[..=i].to_string(),
            None => String::new(),
        };
        let name = format!("{dir}styles.xml");
        let minimal = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <{root_name}{decls}/>"
        );
        self.add_part(name.clone(), minimal.into_bytes());
        self.ensure_content_type_override(&format!("/{name}"), STYLES_CONTENT_TYPE)?;
        let target = name.strip_prefix(&dir).unwrap_or(&name).to_string();
        self.add_relationship(&source, STYLES_REL_TYPE, &target, false)?;

        self.ensure_part(&name)
            .ok_or_else(|| Error::InvalidPackage("created styles part does not parse".into()))
    }
}

/// A direct child element of `parent` with WML local name `local`, if present. (The same
/// finder [`rpr_child_in`] provides, named for its use on `w:style` / `w:docDefaults`.)
fn style_child(tree: &XmlTree, parent: NodeId, local: &str) -> Option<NodeId> {
    rpr_child_in(tree, parent, local)
}
