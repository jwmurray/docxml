//! Field-code authoring: [`Paragraph::add_field`] and the [`Paragraph::add_page_number_field`]
//! convenience.
//!
//! A Word field is not a single element but a *run sequence* delimited by `w:fldChar`
//! markers: a `begin` char, the field's `w:instrText` instruction (e.g. `PAGE`, or a `TOC`
//! switch string), optionally a `separate` char followed by the cached result text, and an
//! `end` char. Word recomputes the result when the document opens; the cached text is what
//! shows until then (and what a non-Word reader sees).
//!
//! This is a write-only API — it emits the sequence but does not parse or evaluate existing
//! fields.

use super::{Document, Paragraph, Run};

impl Paragraph {
    /// Append a Word field as its canonical run sequence, returning a run handle.
    ///
    /// The emitted runs are, in order:
    /// 1. a run holding `w:fldChar w:fldCharType="begin"`;
    /// 2. a run holding `w:instrText` (with `xml:space="preserve"`) carrying `instr`, padded
    ///    with a single leading and trailing space if it lacks them (Word expects the
    ///    instruction delimited by spaces, e.g. `" PAGE "`);
    /// 3. **when `cached_text` is `Some`**, a run holding `w:fldChar w:fldCharType="separate"`
    ///    followed by a run holding a `w:t` with the cached result text;
    /// 4. a run holding `w:fldChar w:fldCharType="end"`.
    ///
    /// Returns the cached-text run when `cached_text` is `Some` (the run a caller most likely
    /// wants to format, since it is the visible result), otherwise the `w:instrText` run.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use docxml::Document;
    ///
    /// let mut doc = Document::new();
    /// let p = doc.add_paragraph("");
    /// p.add_field(&mut doc, "PAGE", Some("1"));
    /// ```
    pub fn add_field(&self, doc: &mut Document, instr: &str, cached_text: Option<&str>) -> Run {
        self.append_fld_char(doc, "begin");
        let instr_run = self.append_instr_text(doc, instr);

        let mut result = instr_run;
        if let Some(cached) = cached_text {
            self.append_fld_char(doc, "separate");
            // A plain text run for the cached result; appended last so far, so it lands
            // after the `separate` char and before the `end` char added below.
            result = self.add_run(doc, cached);
        }
        self.append_fld_char(doc, "end");
        result
    }

    /// Append a `PAGE` field with a cached result of `"1"`, returning the cached-result run.
    ///
    /// Equivalent to `add_field(doc, "PAGE", Some("1"))`. Combined with
    /// [`Section::footer`](crate::Section::footer) (milestone 6), this is how a page-number
    /// footer is built — resolve the footer, take (or add) a paragraph, and drop a page
    /// field into it:
    ///
    /// ```rust,ignore
    /// use docxml::Document;
    ///
    /// let mut doc = Document::open("brief.docx")?;
    /// let section = doc.sections()[0];
    /// if let Some(footer) = section.footer(&mut doc) {
    ///     let para = footer.paragraphs(&doc)[0];
    ///     para.add_run(&mut doc, "Page ");
    ///     para.add_page_number_field(&mut doc);
    /// }
    /// doc.save("brief.docx")?;
    /// # Ok::<(), docxml::Error>(())
    /// ```
    pub fn add_page_number_field(&self, doc: &mut Document) -> Run {
        self.add_field(doc, "PAGE", Some("1"))
    }

    /// Append a run holding a single `w:fldChar` of the given `w:fldCharType`.
    fn append_fld_char(&self, doc: &mut Document, char_type: &str) {
        let part = self.part();
        let r_name = doc.qn(part, "r");
        let fld_name = doc.qn(part, "fldChar");
        let type_attr = doc.qn(part, "fldCharType");
        let node = self.node();

        let tree = doc.tree_mut(part);
        let r = tree.create_element(r_name);
        let fld = tree.create_element(fld_name);
        tree.set_attr(fld, type_attr, char_type);
        tree.append_child(r, fld);
        tree.append_child(node, r);
    }

    /// Append a run holding a space-preserving `w:instrText` carrying `instr` (padded with a
    /// leading/trailing space when absent). Returns that run.
    fn append_instr_text(&self, doc: &mut Document, instr: &str) -> Run {
        let part = self.part();
        let padded = pad_instruction(instr);
        let r_name = doc.qn(part, "r");
        let it_name = doc.qn(part, "instrText");
        let node = self.node();

        let tree = doc.tree_mut(part);
        let r = tree.create_element(r_name);
        let it = tree.create_element(it_name);
        tree.set_attr(it, "xml:space", "preserve");
        let text = tree.create_text(padded);
        tree.append_child(it, text);
        tree.append_child(r, it);
        tree.append_child(node, r);
        Run::from_node(part, r)
    }
}

/// Pad a field instruction with a single leading and trailing space when it lacks them, so
/// the instruction is delimited as Word expects (`"PAGE"` → `" PAGE "`, `" PAGE "` unchanged).
fn pad_instruction(instr: &str) -> String {
    let mut out = String::new();
    if !instr.starts_with(' ') {
        out.push(' ');
    }
    out.push_str(instr);
    if !instr.ends_with(' ') {
        out.push(' ');
    }
    out
}

#[cfg(test)]
mod tests {
    use super::pad_instruction;

    #[test]
    fn pads_only_when_needed() {
        assert_eq!(pad_instruction("PAGE"), " PAGE ");
        assert_eq!(pad_instruction(" PAGE "), " PAGE ");
        assert_eq!(pad_instruction(" PAGE"), " PAGE ");
        assert_eq!(pad_instruction("PAGE "), " PAGE ");
        assert_eq!(pad_instruction("TOC \\o \"1-3\""), " TOC \\o \"1-3\" ");
    }
}
