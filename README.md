# docxml

**Create and edit `.docx` files in Rust — a [python-docx](https://python-docx.readthedocs.io/) for Rust.**

> **Status: functional, pre-1.0.** Open, create, edit, and save real documents with full
> round-trip fidelity: paragraphs, runs, character/paragraph formatting, tables, sections,
> headers/footers, and inline images all work (see the roadmap below). The API may change
> between minor versions until 1.0.

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
- [ ] Styles (read + pass-through first, authoring later)

## License

Licensed under either of

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE))
- MIT license ([LICENSE-MIT](LICENSE-MIT))

at your option.
