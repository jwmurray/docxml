//! The [`HeaderFooter`] handle, the [`HeaderFooterType`] selector, and the [`Section`]
//! accessors that resolve *and create* a section's header and footer parts.
//!
//! Headers and footers live in their own package parts (`word/header1.xml`,
//! `word/footer1.xml`, …), referenced from a section's `w:sectPr` by relationship id.
//! Resolving one is a three-step hop — `w:headerReference/@r:id` → the part-level
//! relationships (`word/_rels/document.xml.rels`) → the header/footer part — after which
//! the part is parsed lazily and cached in the [`Document`]. Because the returned
//! [`Paragraph`]s carry the header/footer part's id, the ordinary [`Paragraph`] / [`Run`]
//! read-and-edit API works on them unchanged.
//!
//! Creating one (milestone 10) is the inverse: allocate a fresh `word/headerN.xml`
//! part, register its content type, add a `header`/`footer` relationship from the document
//! part, and insert the `w:headerReference`/`w:footerReference` into the `w:sectPr`.

use crate::xml::{NodeId, XmlTree};

use super::{Document, Paragraph, PartId, Section, is_wml_element, split_qname};

/// The relationships-namespace URIs — transitional (what Word writes) and strict — used
/// to identify the `r:id` attribute regardless of the prefix a document binds to them.
const REL_URIS: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
];

/// Content type registered for a freshly created header part.
const HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
/// Content type registered for a freshly created footer part.
const FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
/// Relationship type written when creating a header part (the transitional URI Word uses).
const HEADER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
/// Relationship type written when creating a footer part.
const FOOTER_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

/// Which page a header or footer applies to — the `w:type` of a `w:headerReference` /
/// `w:footerReference`.
///
/// Maps one-to-one to the `w:type` attribute values: [`Default`](Self::Default) →
/// `"default"` (odd/every page unless overridden), [`First`](Self::First) → `"first"`
/// (the section's first page, shown only when [`Section::different_first_page`] is set),
/// [`Even`](Self::Even) → `"even"` (even pages, shown only when
/// [`Document::even_and_odd_headers`] is set).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterType {
    /// The default header/footer (`w:type="default"`), used on every page not covered by a
    /// more specific type.
    Default,
    /// The first-page header/footer (`w:type="first"`). Word displays it only when the
    /// section's `w:titlePg` is set — see [`Section::set_different_first_page`].
    First,
    /// The even-page header/footer (`w:type="even"`). Word displays it only when
    /// `w:evenAndOddHeaders` is set in `word/settings.xml` — see
    /// [`Document::set_even_and_odd_headers`]. Creating an [`Even`](Self::Even) header does
    /// *not* set that flag automatically; do both (the examples show it).
    Even,
}

impl HeaderFooterType {
    /// The `w:type` attribute value for this page kind.
    fn type_val(self) -> &'static str {
        match self {
            HeaderFooterType::Default => "default",
            HeaderFooterType::First => "first",
            HeaderFooterType::Even => "even",
        }
    }
}

/// Whether a header or footer is being created/resolved — selects the reference element,
/// part basename, root element, content type, and relationship type that differ between
/// the two otherwise-identical flows.
#[derive(Debug, Clone, Copy)]
enum Kind {
    Header,
    Footer,
}

impl Kind {
    /// The `w:sectPr` reference element local name (`headerReference` / `footerReference`).
    fn ref_local(self) -> &'static str {
        match self {
            Kind::Header => "headerReference",
            Kind::Footer => "footerReference",
        }
    }

    /// The created part's root element local name (`hdr` / `ftr`).
    fn root_local(self) -> &'static str {
        match self {
            Kind::Header => "hdr",
            Kind::Footer => "ftr",
        }
    }

    /// The created part's file-name stem, i.e. `header` → `word/header1.xml`.
    fn basename(self) -> &'static str {
        match self {
            Kind::Header => "header",
            Kind::Footer => "footer",
        }
    }

    /// The content type registered for a created part of this kind.
    fn content_type(self) -> &'static str {
        match self {
            Kind::Header => HEADER_CONTENT_TYPE,
            Kind::Footer => FOOTER_CONTENT_TYPE,
        }
    }

    /// The relationship type written for a created part of this kind.
    fn rel_type(self) -> &'static str {
        match self {
            Kind::Header => HEADER_REL_TYPE,
            Kind::Footer => FOOTER_REL_TYPE,
        }
    }
}

/// A lightweight handle to a header or footer part (`w:hdr` / `w:ftr`).
///
/// Obtained from [`Section::header`] / [`Section::footer`] (and the typed
/// [`header_of_type`](Section::header_of_type) / [`create_header`](Section::create_header)
/// variants). Like the other handles it is `Copy` and borrows nothing — it carries the
/// [`PartId`] of the lazily parsed header or footer part, so its
/// [`paragraphs`](Self::paragraphs) read and edit that part's tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderFooter {
    part: PartId,
}

impl HeaderFooter {
    /// Wrap a header/footer part id.
    pub(crate) fn from_part(part: PartId) -> Self {
        HeaderFooter { part }
    }

    /// The body-level paragraphs of the header/footer: the direct `w:p` children of the
    /// part's root `w:hdr` / `w:ftr`.
    ///
    /// The returned [`Paragraph`]s carry this part's id, so their text reads — and their
    /// runs edit — against the header/footer tree through the ordinary API. Paragraphs
    /// nested inside a header/footer table are not included, matching the body's
    /// `paragraphs` behavior.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let mut doc = docxml::Document::open("basic.docx")?;
    /// let section = doc.sections()[0];
    /// if let Some(header) = section.header(&mut doc) {
    ///     let text: String = header
    ///         .paragraphs(&doc)
    ///         .iter()
    ///         .map(|p| p.text(&doc))
    ///         .collect();
    ///     println!("{text}");
    /// }
    /// # Ok::<(), docxml::Error>(())
    /// ```
    pub fn paragraphs(&self, doc: &Document) -> Vec<Paragraph> {
        let tree = doc.tree(self.part);
        let root = tree.root();
        tree.children(root)
            .iter()
            .copied()
            .filter(|&c| is_wml_element(tree, c, "p"))
            .map(|c| Paragraph::from_node(self.part, c))
            .collect()
    }

    /// Append a paragraph carrying `text` to the header/footer, returning a handle to it.
    ///
    /// The new `w:p` is added as the last child of the part's root `w:hdr` / `w:ftr`; when
    /// `text` is non-empty a single run carrying it is added (reusing
    /// [`Paragraph::add_run`], so `xml:space="preserve"` is applied when needed). Edit the
    /// returned [`Paragraph`] / its [`Run`](crate::Run)s further through the ordinary API.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, HeaderFooterType};
    ///
    /// let mut doc = Document::new();
    /// let section = doc.sections()[0];
    /// let header = section.create_header(&mut doc, HeaderFooterType::Default);
    /// header.add_paragraph(&mut doc, "Confidential");
    /// assert_eq!(header.paragraphs(&doc).last().unwrap().text(&doc), "Confidential");
    /// ```
    pub fn add_paragraph(&self, doc: &mut Document, text: &str) -> Paragraph {
        let name = doc.qn(self.part, "p");
        let root = doc.tree(self.part).root();
        let tree = doc.tree_mut(self.part);
        let p = tree.create_element(name);
        tree.append_child(root, p);
        let para = Paragraph::from_node(self.part, p);
        if !text.is_empty() {
            para.add_run(doc, text);
        }
        para
    }
}

impl Section {
    /// The section's default header, if it references one. A [`HeaderFooterType::Default`]
    /// shorthand for [`header_of_type`](Self::header_of_type); see it for the resolution
    /// rules and why this takes `&mut Document`.
    pub fn header(&self, doc: &mut Document) -> Option<HeaderFooter> {
        self.header_of_type(doc, HeaderFooterType::Default)
    }

    /// The section's default footer, if it references one. A [`HeaderFooterType::Default`]
    /// shorthand for [`footer_of_type`](Self::footer_of_type).
    pub fn footer(&self, doc: &mut Document) -> Option<HeaderFooter> {
        self.footer_of_type(doc, HeaderFooterType::Default)
    }

    /// The section's header of the given page type, if it references one.
    ///
    /// Resolves the `w:headerReference` whose `w:type` matches `kind` through the part-level
    /// relationships to a header part, parses it lazily, and caches it in `doc`. Returns
    /// `None` when the section has no such reference or the target part cannot be resolved
    /// or parsed.
    ///
    /// Takes `&mut Document` because resolving may parse and cache a new part; parsing alone
    /// does **not** mark the document modified, so a read-only header access leaves every
    /// part byte-identical on save.
    pub fn header_of_type(
        &self,
        doc: &mut Document,
        kind: HeaderFooterType,
    ) -> Option<HeaderFooter> {
        self.hdr_ftr_ref(doc, "headerReference", kind)
    }

    /// The section's footer of the given page type, if it references one. See
    /// [`header_of_type`](Self::header_of_type) for the resolution rules.
    pub fn footer_of_type(
        &self,
        doc: &mut Document,
        kind: HeaderFooterType,
    ) -> Option<HeaderFooter> {
        self.hdr_ftr_ref(doc, "footerReference", kind)
    }

    /// Create (or return the existing) header of the given page type for this section.
    ///
    /// If the section already references a header of `kind`, that one is returned unchanged
    /// — no duplicate part or reference is made. Otherwise a fresh header is created: the
    /// next free `word/headerN.xml` part (with a minimal `w:hdr` root carrying one empty
    /// `w:p` and the document's full namespace set, so it is self-contained), an
    /// `[Content_Types].xml` `Override`, a `header` relationship from the document part, and
    /// a `w:headerReference` (namespace-correct `r:id`, `w:type`) inserted first in the
    /// `w:sectPr` per `CT_SectPr` order.
    ///
    /// A [`First`](HeaderFooterType::First) header is shown by Word only when
    /// [`set_different_first_page`](Self::set_different_first_page) is also set; an
    /// [`Even`](HeaderFooterType::Even) header only when
    /// [`Document::set_even_and_odd_headers`] is — creating the header does not set those
    /// flags for you.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, HeaderFooterType};
    ///
    /// let mut doc = Document::new();
    /// let section = doc.sections()[0];
    ///
    /// // A distinct first-page header: create it *and* turn on the title-page flag.
    /// let first = section.create_header(&mut doc, HeaderFooterType::First);
    /// first.add_paragraph(&mut doc, "Cover page");
    /// section.set_different_first_page(&mut doc, true);
    ///
    /// assert!(section.header_of_type(&mut doc, HeaderFooterType::First).is_some());
    /// ```
    ///
    /// # Panics
    ///
    /// Panics only if the created part cannot be registered — which for a document built
    /// from [`Document::new`] or opened from a valid package does not happen.
    pub fn create_header(&self, doc: &mut Document, kind: HeaderFooterType) -> HeaderFooter {
        self.create_ref(doc, Kind::Header, kind)
    }

    /// Create (or return the existing) footer of the given page type for this section. The
    /// footer twin of [`create_header`](Self::create_header); see it for the full behavior.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, HeaderFooterType};
    ///
    /// let mut doc = Document::new();
    /// let section = doc.sections()[0];
    /// let footer = section.create_footer(&mut doc, HeaderFooterType::Default);
    /// footer.add_paragraph(&mut doc, "Page footer");
    /// assert!(section.footer(&mut doc).is_some());
    /// ```
    pub fn create_footer(&self, doc: &mut Document, kind: HeaderFooterType) -> HeaderFooter {
        self.create_ref(doc, Kind::Footer, kind)
    }

    /// Remove this section's header reference of the given page type, returning whether one
    /// was removed.
    ///
    /// Only the `w:headerReference` element is deleted from the `w:sectPr`; the header
    /// *part* is deliberately left in the package, orphaned. This matches python-docx's
    /// `is_linked_to_previous = True` semantics and is safe: another section may still
    /// reference the same part, and an OPC part that nothing references is harmless
    /// pass-through content (it round-trips byte-for-byte and Word ignores it). Reclaiming
    /// orphaned parts is a package-level concern, not this operation's job.
    pub fn remove_header_reference(&self, doc: &mut Document, kind: HeaderFooterType) -> bool {
        self.remove_ref(doc, "headerReference", kind)
    }

    /// Remove this section's footer reference of the given page type, returning whether one
    /// was removed. The footer twin of
    /// [`remove_header_reference`](Self::remove_header_reference); the footer part is
    /// likewise left orphaned (see that method for why).
    pub fn remove_footer_reference(&self, doc: &mut Document, kind: HeaderFooterType) -> bool {
        self.remove_ref(doc, "footerReference", kind)
    }

    /// Shared header/footer resolution: find the reference of type `kind` under the
    /// `w:sectPr`, read its relationship id, resolve it to a part, and lazily parse it.
    fn hdr_ftr_ref(
        &self,
        doc: &mut Document,
        ref_local: &str,
        kind: HeaderFooterType,
    ) -> Option<HeaderFooter> {
        let part = self.part();
        let r_id = {
            let tree = doc.tree(part);
            let type_attr = doc.qn(part, "type");
            let sect = self.node();
            tree.children(sect)
                .iter()
                .copied()
                .filter(|&c| is_wml_element(tree, c, ref_local))
                .find(|&c| tree.attr(c, &type_attr) == Some(kind.type_val()))
                .and_then(|rf| rel_id_attr(tree, rf))
                .map(str::to_owned)
        }?;

        let source = doc.main_part_name().to_string();
        let target = doc.resolve_rel_target(&source, &r_id)?;
        let hf_part = doc.ensure_part(&target)?;
        Some(HeaderFooter::from_part(hf_part))
    }

    /// Shared create path for headers and footers (see [`create_header`](Self::create_header)).
    fn create_ref(&self, doc: &mut Document, kind: Kind, page: HeaderFooterType) -> HeaderFooter {
        // Idempotent: an existing reference of this type wins — no duplicate part/reference.
        if let Some(existing) = self.hdr_ftr_ref(doc, kind.ref_local(), page) {
            return existing;
        }

        // The document part's directory (`word/`), where the new part and its relative
        // relationship target live.
        let source = doc.main_part_name().to_string();
        let dir = match source.rfind('/') {
            Some(i) => source[..=i].to_string(),
            None => String::new(),
        };

        // A fresh `word/headerN.xml` (or `footerN.xml`) with a self-contained root.
        let part_name = next_part_name(doc, &dir, kind.basename());
        let xml = build_hdr_ftr_xml(doc, kind.root_local());
        doc.add_part(part_name.clone(), xml);
        doc.ensure_content_type_override(&format!("/{part_name}"), kind.content_type())
            .expect("[Content_Types].xml is editable");
        let target = part_name
            .strip_prefix(&dir)
            .unwrap_or(&part_name)
            .to_string();
        let r_id = doc
            .add_relationship(&source, kind.rel_type(), &target, false)
            .expect("document relationships part is editable");

        // Parse the new part so the returned handle can read/edit it.
        let hf_part = doc
            .ensure_part(&part_name)
            .expect("created header/footer part parses");

        self.insert_hdr_ftr_ref(doc, kind.ref_local(), page, &r_id);
        HeaderFooter::from_part(hf_part)
    }

    /// Insert a new `w:headerReference` / `w:footerReference` (with `w:type` and a
    /// namespace-correct `r:id`) into the `w:sectPr`, first in `CT_SectPr` order.
    fn insert_hdr_ftr_ref(
        &self,
        doc: &mut Document,
        ref_local: &str,
        page: HeaderFooterType,
        r_id: &str,
    ) {
        let part = self.part();
        let ref_name = doc.qn(part, ref_local);
        let type_attr = doc.qn(part, "type");
        let id_attr = rel_id_attr_name(doc.tree(part));
        let index = self.sect_insert_index(doc.tree(part), ref_local);

        let tree = doc.tree_mut(part);
        let el = tree.create_element(ref_name);
        // Attribute order matches Word's output: w:type then r:id.
        tree.set_attr(el, type_attr, page.type_val());
        tree.set_attr(el, id_attr, r_id);
        tree.insert_child(self.node(), index, el);
    }

    /// Remove the reference of type `kind` (leaving the part orphaned); see
    /// [`remove_header_reference`](Self::remove_header_reference).
    fn remove_ref(&self, doc: &mut Document, ref_local: &str, kind: HeaderFooterType) -> bool {
        let part = self.part();
        let node = {
            let tree = doc.tree(part);
            let type_attr = doc.qn(part, "type");
            tree.children(self.node())
                .iter()
                .copied()
                .filter(|&c| is_wml_element(tree, c, ref_local))
                .find(|&c| tree.attr(c, &type_attr) == Some(kind.type_val()))
        };
        match node {
            Some(n) => {
                doc.tree_mut(part).remove_from_parent(n);
                true
            }
            None => false,
        }
    }
}

/// The next free `<dir><basename>N.xml` part name: `max(existing N) + 1`, scanning parts
/// already named `<dir><basename><digits>.xml` (so headers and footers number
/// independently — `header1.xml` alongside `footer1.xml`).
fn next_part_name(doc: &Document, dir: &str, basename: &str) -> String {
    let prefix = format!("{dir}{basename}");
    let mut max = 0u32;
    for part in doc.package().parts() {
        let Some(rest) = part.name.strip_prefix(&prefix) else {
            continue;
        };
        let digits: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
        if !digits.is_empty() && rest[digits.len()..] == *".xml" {
            if let Ok(n) = digits.parse::<u32>() {
                max = max.max(n);
            }
        }
    }
    format!("{prefix}{}.xml", max + 1)
}

/// Build the raw bytes of a minimal, self-contained header/footer part: an XML declaration,
/// a `w:hdr`/`w:ftr` root carrying **every namespace declaration from the main document
/// root** (so the created part resolves `w:`, `r:`, `w14:`, … exactly as the document does,
/// without depending on any element's ancestors), and one empty `w:p` child.
///
/// The root and paragraph element names use the document's own WordprocessingML prefix
/// (via [`Document::qn`]); namespace URIs are copied verbatim from the root's `xmlns[:*]`
/// attributes (URIs contain no XML metacharacters, matching the from-scratch construction in
/// `numbering.rs`).
fn build_hdr_ftr_xml(doc: &Document, root_local: &str) -> Vec<u8> {
    let root_name = doc.qn(PartId::MAIN, root_local);
    let p_name = doc.qn(PartId::MAIN, "p");

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

    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <{root_name}{decls}><{p_name}/></{root_name}>"
    )
    .into_bytes()
}

/// The value of an element's relationship-id attribute — the attribute whose local name is
/// `id` and whose prefix resolves (via the in-scope namespace declarations) to a
/// relationships namespace URI, transitional or strict. This is the namespace-correct way
/// to read `r:id` without assuming the `r` prefix.
fn rel_id_attr(tree: &XmlTree, node: NodeId) -> Option<&str> {
    for (key, value) in tree.attrs(node) {
        let (prefix, local) = split_qname(key);
        if local != "id" {
            continue;
        }
        if let Some(uri) = tree.namespace_uri(node, prefix) {
            if REL_URIS.contains(&uri) {
                return Some(value.as_str());
            }
        }
    }
    None
}

/// The qualified attribute name to *write* a relationship id (`r:id`), using whatever prefix
/// the tree's root binds to the relationships URI — the write-side counterpart of
/// [`rel_id_attr`]. Falls back to the conventional `r` prefix when the root declares none
/// (in practice `word/document.xml` always declares `xmlns:r`).
fn rel_id_attr_name(tree: &XmlTree) -> String {
    let root = tree.root();
    for (key, value) in tree.attrs(root) {
        if !REL_URIS.contains(&value.as_str()) {
            continue;
        }
        if let Some(prefix) = key.strip_prefix("xmlns:") {
            return format!("{prefix}:id");
        }
        if key == "xmlns" {
            return "id".to_string();
        }
    }
    "r:id".to_string()
}
