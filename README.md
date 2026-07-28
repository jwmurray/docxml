# docxml

**Create and edit `.docx` files in Rust — a [python-docx](https://python-docx.readthedocs.io/) for Rust.**

> **Status: functional, pre-1.0 — python-docx parity plus the court-formatting layer
> python-docx never had.** Open, create, edit, and save real documents with full
> round-trip fidelity — paragraphs, runs, character/paragraph formatting, tables
> (including table and cell borders), sections, headers/footers, inline images,
> numbering/lists, hyperlinks, line numbering/frames/hidden text, styles (catalog read,
> authoring, `docDefaults`, effective formatting reads), footnotes/endnotes, the typed
> section grid (`w:docGrid` line pitch — exact lines-per-page), typed `w:pgNumType`,
> and comments/tracked-changes authoring all work (see the roadmap below). The API may
> change between minor versions until 1.0.

## Why another docx crate?

Existing Rust docx crates are write-focused: they generate documents from typed structs.
That works for creation, but editing an *existing* document — a template, a contract, a
court filing — silently drops everything the structs don't model.

`docxml` takes the opposite architecture, the one python-docx got right:

- **Lossless core.** Every part of the package is parsed into a mutable, namespace-aware
  XML tree. Anything the library doesn't understand passes through byte-for-byte on save.
  Open → save is a faithful round trip, always.
- **Typed API on top.** `Document`, `Paragraph`, `Run`, `Table` are lightweight handles
  (arena node ids) into that tree — ergonomic accessors without `Rc<RefCell<>>` soup.
- **One code path.** Creating a document is editing an embedded blank one, exactly like
  python-docx's `default.docx`.

## Example

```rust,ignore
use docxml::Document;

let mut doc = Document::open("contract.docx")?;

for para in doc.paragraphs() {
    println!("{}", para.text(&doc));
}

let p = doc.add_paragraph("Signed and agreed:");
p.add_run(&mut doc, "John Murray").bold(&mut doc, true);

let table = doc.add_table(2, 2);
for (r, row) in table.rows(&doc).into_iter().enumerate() {
    for (c, cell) in row.cells(&doc).into_iter().enumerate() {
        cell.set_text(&mut doc, &format!("r{r}c{c}"));
    }
}

doc.save("contract-signed.docx")?;
```

## Roadmap

- [x] OPC packaging layer (zip, relationships, byte-for-byte round-trip test)
- [x] Lossless mutable XML tree with semantic round-trip tests against real-world documents
- [x] `Document` / `Paragraph` / `Run` — text read/edit, bold/italic, embedded blank template (create)
- [x] Character/paragraph formatting (underline, size, color, font, alignment, styles read)
- [x] Tables (read rows/cells/text, merge awareness, create, `add_row`, cell `set_text`)
- [x] Sections, headers/footers (page geometry read/set via `Length`; header/footer text read + edit through lazily parsed parts)
- [x] Images (inline pictures — read, add with EMU geometry, media part + content-type + relationship)
- [x] Paragraph formatting (line spacing, space before/after, indents, tab stops, keep/page-break-before), breaks, field codes (PAGE/TOC)
- [x] Numbering / lists (numbering.xml authoring; List Bullet / List Number, `create_numbering`, restartable lists)
- [x] Header/footer part creation; first/even-page headers (`create_header`/`create_footer` with `HeaderFooterType`, `different_first_page`/`even_and_odd_headers` flags)
- [x] Table column widths, merge creation, grid-based cell addressing (`set_column_widths`/`Cell::width`/`set_fixed_layout`, grid-based `Table::cell`, `Table::merge`)
- [x] Hyperlinks (read `hyperlinks()`; `add_hyperlink`/`add_anchor_hyperlink` with External relationship creation, `add_bookmark`)
- [x] Section line numbering, paragraph frames/borders, hidden text (`Section::set_line_numbering` with `LineNumbering`/`LineNumberRestart`, `Paragraph::suppress_line_numbers`, `Paragraph::set_frame`/`set_borders` with `FrameOptions`/`BorderEdge`, `Run::set_vanish`)
- [x] Styles authoring, `docDefaults`, style-aware formatting reads (`Style`/`StyleType` catalog via `Document::styles`/`style_by_id`/`style_by_name`, `create_style` with `set_based_on`/`set_next` and reused bold/size/color/font/alignment/spacing setters, `Paragraph::style_name`, `Run::set_style_id`, `Document::set_default_font`, and `Run::effective_bold`/`effective_italic`/`effective_size`/`effective_font`)
- [x] Table and cell borders (`Table::set_borders`/`borders` writing `w:tblBorders` outer/inside edges, `Cell::set_borders`/`borders` writing `w:tcBorders`, reusing the paragraph `BorderEdge`/`BorderStyle` edge model)
- [x] Footnotes and endnotes (`Paragraph::add_footnote`/`add_endnote` creating `word/footnotes.xml`/`word/endnotes.xml` with the required separator/continuation stub notes, superscripted `w:footnoteReference`/`w:endnoteReference` marks, per-note style + point-size control through the ordinary `Paragraph`/`Run` API, and the `Document::footnotes`/`endnotes` read-back — count, `id`/`text`/`point_sizes`, `Run::footnote_reference_id`)
- [x] Typed section grid and page numbering (`Section::set_doc_grid` with `DocGrid`/`DocGridType` — `w:docGrid` incl. `linePitch`, the lines-per-page control that pins the 28-line pleading page; `Section::set_page_numbering` with `PageNumbering` — `w:pgNumType` format/start)
- [x] Per-paragraph grid binding (`Paragraph::set_snap_to_grid`/`snap_to_grid` — `w:pPr/w:snapToGrid` as a tri-state: bare element states "snap", `w:val="0"` exempts the paragraph, absent inherits; the paragraph-side half of `w:docGrid`, so one document can carry grid-bound body text over freely led headers and block quotes)
- [x] Measurement ergonomics (`Pt` is `PartialOrd` — type-floor checks compare directly; `Length` carries full `Ord`/`Hash`; ordered-insertion audit: every property container writes through a canonical `CT_*` order list)
- [x] Comments and tracked changes (`Paragraph::add_comment` creating `word/comments.xml` with `w:commentRangeStart`/`End` + `w:commentReference` anchors and the `Document::comments` read-back; `Paragraph::add_inserted_run`/`add_deleted_run` wrapping runs in attributed `w:ins`/`w:del` with `w:delText` preservation; `Paragraph::revisions` read-back; `Paragraph::text` reads the *final* view — insertions included, deletions excluded)
- [ ] Anchored (floating) images — design note: `Picture` today is inline-only (`wp:inline`). Anchoring means replacing the `wp:inline` wrapper with `wp:anchor` carrying `behindDoc`/`locked`/`layoutInCell`/`allowOverlap` flags, `wp:simplePos`, `wp:positionH`/`wp:positionV` (each `relativeFrom` + `wp:posOffset` in EMU), and a wrap element (`wp:wrapNone`/`wrapSquare`/`wrapTopAndBottom`); the inner `a:graphic` subtree is unchanged, so the natural API is `Picture::anchor_at(&mut doc, AnchorOptions { h: (HRelativeFrom, Length), v: (VRelativeFrom, Length), wrap: WrapKind, behind_text: bool })` converting in place plus an `is_anchored` read. Court filings rarely float images (exhibits are inline), so this waits for a consumer with a measured exemplar — implement only with round-trip tests and a rendered-position measurement.

## License

Licensed under either of

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE))
- MIT license ([LICENSE-MIT](LICENSE-MIT))

at your option.
