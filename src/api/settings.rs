//! Document settings: the `word/settings.xml` part, and the few settings the typed API
//! reads and writes there.
//!
//! At this milestone that is [`Document::even_and_odd_headers`] — the `w:evenAndOddHeaders`
//! flag that tells Word to honor `"even"`-type header/footer references on even pages.
//!
//! Like the numbering part (see [`Document::create_numbering`](crate::Document::create_numbering)),
//! the settings part is resolved through the document relationships and parsed lazily. When
//! a document has no settings part at all, [`set_even_and_odd_headers`](Document::set_even_and_odd_headers)
//! creates one from scratch (a minimal `w:settings` root, an `[Content_Types].xml`
//! `Override`, and a `settings` relationship from the document part); the blank template and
//! every fixture already ship one, so that path is only for documents that never had it.

use crate::error::{Error, Result};

use super::{Document, PartId, is_wml_element, ordered_insert_index, rank_in};

/// Relationship types identifying the settings part (transitional and strict).
const SETTINGS_REL_TYPES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/settings",
];
/// Relationship type written when creating a settings part (the transitional URI Word uses).
const SETTINGS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
/// Content type registered for a freshly created settings part.
const SETTINGS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
/// The transitional WordprocessingML main URI, declared on a from-scratch settings root.
const WML_MAIN_URI: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// Canonical `w:settings` child order (ECMA-376 §17.15.1.78, `CT_Settings` sequence),
/// local names only — enough of the (long) sequence to slot `w:evenAndOddHeaders` correctly
/// among the properties Word actually writes. `w:evenAndOddHeaders` follows `w:defaultTabStop`
/// / `w:defaultTableStyle` and precedes `w:characterSpacingControl`, `w:compat`, `w:rsids`,
/// etc. Unlisted children rank last, so a from-template settings part keeps its trailing
/// (`m:mathPr`, `w:clrSchemeMapping`, …) content after the inserted flag.
const SETTINGS_ORDER: &[&str] = &[
    "writeProtection",
    "view",
    "zoom",
    "removePersonalInformation",
    "doNotDisplayPageBoundaries",
    "displayBackgroundShape",
    "printPostScriptOverText",
    "printFractionalCharacterWidth",
    "printFormsData",
    "embedTrueTypeFonts",
    "embedSystemFonts",
    "saveSubsetFonts",
    "saveFormsData",
    "mirrorMargins",
    "alignBordersAndEdges",
    "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter",
    "gutterAtTop",
    "hideSpellingErrors",
    "hideGrammaticalErrors",
    "activeWritingStyle",
    "proofState",
    "formsDesign",
    "attachedTemplate",
    "linkStyles",
    "stylePaneFormatFilter",
    "stylePaneSortMethod",
    "documentType",
    "mailMerge",
    "revisionView",
    "trackChanges",
    "doNotTrackMoves",
    "doNotTrackFormatting",
    "documentProtection",
    "autoFormatOverride",
    "styleLockTheme",
    "styleLockQFSet",
    "defaultTabStop",
    "autoHyphenation",
    "consecutiveHyphenLimit",
    "hyphenationZone",
    "doNotHyphenateCaps",
    "showEnvelope",
    "summaryLength",
    "clickAndTypeStyle",
    "defaultTableStyle",
    "evenAndOddHeaders",
    "bookFoldRevPrinting",
    "bookFoldPrinting",
    "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery",
    "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin",
    "doNotShadeFormData",
    "noPunctuationKerning",
    "characterSpacingControl",
    "printTwoOnOne",
    "strictFirstAndLastChars",
    "noLineBreaksAfter",
    "noLineBreaksBefore",
    "savePreviewPicture",
    "doNotValidateAgainstSchema",
    "saveInvalidXml",
    "ignoreMixedContent",
    "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml",
    "saveXmlDataOnly",
    "useXSLTWhenSaving",
    "saveThroughXslt",
    "showXMLTags",
    "alwaysMergeEmptyNamespace",
    "updateFields",
    "hdrShapeDefaults",
    "footnotePr",
    "endnotePr",
    "compat",
    "rsids",
    "mathPr",
    "attachedSchema",
    "themeFontLang",
    "clrSchemeMapping",
    "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures",
    "forceUpgrade",
    "captions",
    "readModeInkLockDown",
    "smartTagType",
    "schemaLibrary",
    "shapeDefaults",
    "doNotEmbedSmartTags",
    "decimalSymbol",
    "listSeparator",
];

impl Document {
    /// Whether even/odd-page header and footer differentiation is on
    /// (`word/settings.xml/w:evenAndOddHeaders`).
    ///
    /// `w:evenAndOddHeaders` is an on/off setting (`CT_OnOff`): present and not explicitly
    /// `w:val="0"`/`"false"` means on. When on, Word honors `"even"`-type header/footer
    /// references (see [`HeaderFooterType::Even`](crate::HeaderFooterType::Even)) on even
    /// pages; when off, the default header/footer is used on every page.
    ///
    /// Takes `&mut self` because the settings part is parsed lazily on first access (like
    /// header/footer resolution) — parsing does **not** mark it modified, so a read leaves
    /// every part byte-identical on save. Returns `false` when the document has no settings
    /// part or the flag is absent.
    pub fn even_and_odd_headers(&mut self) -> bool {
        let Some(name) = self.part_by_rel_type(&SETTINGS_REL_TYPES) else {
            return false;
        };
        let Some(part) = self.ensure_part(&name) else {
            return false;
        };
        let tree = self.tree(part);
        let root = tree.root();
        match tree
            .children(root)
            .iter()
            .copied()
            .find(|&c| is_wml_element(tree, c, "evenAndOddHeaders"))
        {
            Some(el) => match tree.attr(el, &self.qn(part, "val")) {
                Some(v) => !matches!(v, "0" | "false"),
                None => true,
            },
            None => false,
        }
    }

    /// Turn even/odd-page header/footer differentiation on or off
    /// (`word/settings.xml/w:evenAndOddHeaders`).
    ///
    /// On ensures a bare `w:evenAndOddHeaders` element (in `CT_Settings` order, clearing any
    /// explicit `w:val="0"/"false"`); off removes it. Only `word/settings.xml` is modified —
    /// the write is surgical, so on a document that already has a settings part every other
    /// part stays byte-identical on save.
    ///
    /// Creating an [`Even`](crate::HeaderFooterType::Even) header/footer does *not* set this
    /// flag; do both. If the document has no settings part, one is created (part, content-type
    /// `Override`, and `settings` relationship), mirroring numbering-part creation.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::{Document, HeaderFooterType};
    ///
    /// let mut doc = Document::new();
    /// let section = doc.sections()[0];
    /// // An even-page footer is only shown once the flag is set — so do both.
    /// section.create_footer(&mut doc, HeaderFooterType::Even);
    /// doc.set_even_and_odd_headers(true);
    /// assert!(doc.even_and_odd_headers());
    /// ```
    ///
    /// # Panics
    ///
    /// Panics only if the settings part must be created and `[Content_Types].xml` or the
    /// relationships part cannot be edited — which for a [`Document::new`] or a valid opened
    /// package does not happen.
    pub fn set_even_and_odd_headers(&mut self, on: bool) {
        self.try_set_even_and_odd_headers(on)
            .expect("settings part is present or can be created");
    }

    /// Fallible core of [`set_even_and_odd_headers`](Self::set_even_and_odd_headers).
    fn try_set_even_and_odd_headers(&mut self, on: bool) -> Result<()> {
        if !on {
            // Off: remove the flag if present. If there is no settings part, nothing to do.
            if let Some(name) = self.part_by_rel_type(&SETTINGS_REL_TYPES) {
                if let Some(part) = self.ensure_part(&name) {
                    let existing = {
                        let tree = self.tree(part);
                        let root = tree.root();
                        tree.children(root)
                            .iter()
                            .copied()
                            .find(|&c| is_wml_element(tree, c, "evenAndOddHeaders"))
                    };
                    if let Some(el) = existing {
                        self.tree_mut(part).remove_from_parent(el);
                    }
                }
            }
            return Ok(());
        }

        let part = self.ensure_settings_part()?;
        let val = self.qn(part, "val");

        // Existing flag: just clear any explicit off value.
        let existing = {
            let tree = self.tree(part);
            let root = tree.root();
            tree.children(root)
                .iter()
                .copied()
                .find(|&c| is_wml_element(tree, c, "evenAndOddHeaders"))
        };
        if let Some(el) = existing {
            let is_off = self
                .tree(part)
                .attr(el, &val)
                .is_some_and(|v| matches!(v, "0" | "false"));
            if is_off {
                self.tree_mut(part).remove_attr(el, &val);
            }
            return Ok(());
        }

        // Create the flag in CT_Settings order.
        let name = self.qn(part, "evenAndOddHeaders");
        let root = self.tree(part).root();
        let index = ordered_insert_index(
            self.tree(part),
            root,
            rank_in(SETTINGS_ORDER, "evenAndOddHeaders"),
            SETTINGS_ORDER,
        );
        let tree = self.tree_mut(part);
        let el = tree.create_element(name);
        tree.insert_child(root, index, el);
        Ok(())
    }

    /// Ensure the settings part is present and parsed, returning its [`PartId`].
    ///
    /// Uses the existing part if the document references one via a `settings` relationship;
    /// otherwise creates it (a minimal `w:settings` root part, an `[Content_Types].xml`
    /// `Override`, and a `settings` relationship from the document part).
    fn ensure_settings_part(&mut self) -> Result<PartId> {
        if let Some(name) = self.part_by_rel_type(&SETTINGS_REL_TYPES) {
            if let Some(id) = self.ensure_part(&name) {
                return Ok(id);
            }
        }

        let source = self.main_part_name().to_string();
        let dir = match source.rfind('/') {
            Some(i) => source[..=i].to_string(),
            None => String::new(),
        };
        let name = format!("{dir}settings.xml");
        let minimal = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
             <w:settings xmlns:w=\"{WML_MAIN_URI}\"/>"
        );
        self.add_part(name.clone(), minimal.into_bytes());
        self.ensure_content_type_override(&format!("/{name}"), SETTINGS_CONTENT_TYPE)?;
        let target = name.strip_prefix(&dir).unwrap_or(&name).to_string();
        self.add_relationship(&source, SETTINGS_REL_TYPE, &target, false)?;

        self.ensure_part(&name)
            .ok_or_else(|| Error::InvalidPackage("created settings part does not parse".into()))
    }
}
