# docxml — Design Specification

Create and edit `.docx` files in Rust: a python-docx for Rust, with editing existing
documents as a first-class use case, not an afterthought.

## Goals

1. **Edit and create.** Open any real-world `.docx` (templates, contracts, court
   filings), modify it, save it — and never corrupt or silently drop content.
2. **python-docx-level ergonomics**, adapted to Rust idioms rather than literally
   transliterated.
3. **Zero Python.** Pure Rust, single static binary, no runtime dependencies.

## Non-goals

- Rendering, layout, or PDF conversion.
- Full OOXML schema coverage as typed structs (see Fidelity Contract for why).
- `.doc` (binary Word), `.odt`, or spreadsheet/presentation formats.

## The Fidelity Contract

**Anything the library does not understand passes through unchanged on save.**

This is the load-bearing decision. Existing Rust docx crates (docx-rs, docx-rust)
model the document as typed structs; whatever the structs don't cover is lost on
write. That is fine for generation and fatal for editing. python-docx wins at editing
because lxml keeps the full XML tree intact and the API mutates it in place.

docxml adopts the same architecture, enforced by test at every layer:

| Layer | Guarantee | Test |
|---|---|---|
| OPC package | Untouched parts are byte-identical on round trip | `tests/roundtrip.rs` (implemented) |
| XML tree | Parse → serialize is semantically lossless: every element, attribute, namespace declaration, and text node preserved, in order | semantic diff against real-world fixtures (next) |
| Typed API | Mutations touch only the nodes they target | per-feature tests |

Byte-identical XML serialization is *not* promised once a part is parsed and
re-serialized (attribute quoting, self-closing forms may normalize); semantic
equivalence is. Parts are parsed lazily — a part never touched is never re-serialized,
so it stays byte-identical.

## Architecture

Three layers, bottom to top:

### 1. OPC packaging (`src/opc/`) — implemented

- `Package`: ordered `Vec<Part>` read from the zip; order preserved for save.
- `Part`: entry name + raw bytes.
- Relationships parsed from `.rels`; main document located via the officeDocument
  relationship (transitional and strict URIs), never a hard-coded path.

### 2. Lossless XML tree (`src/xml/`) — next

A mutable, namespace-aware DOM — the lxml role. No existing crate is a drop-in
(quick-xml is event-based; typed-serde crates violate the fidelity contract), so this
is a small owned layer over quick-xml events:

- **Arena storage**: nodes live in a `Vec<Node>` inside the tree; `NodeId(u32)`
  indexes into it. Handles are `Copy`. No `Rc<RefCell<>>`.
- **Node data**: qualified name (namespace URI + local name + recorded prefix),
  attributes in document order, children (elements, text, comments, PIs) in document
  order. Unknown content is a first-class citizen, not an error.
- **Namespace handling**: prefixes recorded as written so serialization round-trips
  Word's conventional `w:`, `r:`, `wp:` prefixes; lookup happens by URI.
- **Mutation API**: insert/remove/replace children, set attributes, splice subtrees.

### 3. Typed API (`src/api/` or crate root) — after the tree

python-docx's proxy pattern, Rust-flavored:

- `Document` owns the `Package` and parsed trees.
- `Paragraph`, `Run`, `Table`, `Cell`, `Section` are lightweight handles: a `NodeId`
  plus phantom typing. They borrow nothing, so the borrow checker stays out of the way;
  all reads/mutations go through `&Document` / `&mut Document`.
- Creation = editing an embedded blank document (`include_bytes!` default template),
  exactly python-docx's `default.docx` approach. One code path for create and edit.

```rust,ignore
let mut doc = Document::open("contract.docx")?;
for para in doc.paragraphs() {
    println!("{}", para.text(&doc));
}
let p = doc.add_paragraph("Signed and agreed:");
p.add_run(&mut doc, "John Murray").bold(&mut doc, true);
doc.save("contract-signed.docx")?;
```

## Concurrency

**Synchronous core; no async runtime; parallelism belongs to the caller.**

- **No tokio.** The workload is CPU + local file I/O; async buys nothing, and forcing
  a runtime on every consumer is an adoption killer. Format crates (`zip`, `calamine`,
  `lopdf`) are conventionally sync; async callers wrap operations in `spawn_blocking`.
  An optional `tokio` feature for async open/save may land later if demand shows —
  the core API stays sync regardless.
- **No rayon in the core.** Documents are small and parts parse lazily; single-threaded
  quick-xml handles a multi-MB document in low milliseconds. There is nothing worth
  parallelizing inside one document.
- **The real win is across documents** (batch-process hundreds of files), and that is
  the caller's job. The contract that enables it: `Package`, `Document`, and the tree
  types are `Send` — guaranteed by the arena design (`Vec<Node>` + `NodeId`, no
  `Rc`/`RefCell`). A `static_assertions`-style compile-time `Send` check accompanies
  the tree layer when it lands.

## Dependencies

Kept deliberately minimal: `zip`, `quick-xml`, `thiserror`. Dev: `tempfile`.
`rust-version` pinned (currently 1.85); dependency bumps must respect it.

## Testing strategy

- **Fixtures are real documents.** Primary fixture generated by python-docx itself
  (`tests/fixtures/basic.docx`); Word- and Google-Docs-authored fixtures to be added
  as the tree layer lands. python-docx is also the behavioral reference: where its
  semantics are sane, match them.
- **Round-trip tests gate every layer** (see Fidelity Contract table).
- CI: `cargo fmt --check`, `clippy -D warnings`, `cargo test` on every PR.

## Milestones

1. [x] OPC packaging layer + byte-level round-trip test (PR #1)
2. [x] Lossless XML tree + semantic round-trip tests against real-world fixtures (PR #4)
3. [x] `Document` / `Paragraph` / `Run` + text read/edit; embedded blank template (create) (PR #6)
4. [x] Character/paragraph formatting (bold, italic, size, color, alignment, styles read) (PR #7)
5. [x] Tables (PR #8)
6. [x] Sections, headers/footers (PR #9)
7. [x] Images (inline pictures, EMU geometry) (PR #10)
8. [x] Paragraph formatting (spacing/indent/line-spacing/tabs), breaks, field codes (PAGE/TOC) (PR #12)
9. [x] Numbering / lists (numbering.xml authoring; List Bullet / List Number) (PR #13)
10. [x] Header/footer part creation; first/even-page headers (PR #14)
11. [x] Table column widths, merge creation, grid-based cell addressing (PR #15)
12. [x] Hyperlinks (read + write API with relationship creation; bookmarks) (PR #16)
13. [x] Section line numbering (w:lnNumType), paragraph frames (w:framePr, w:pBdr), hidden text (w:vanish) (PR #17)
14. [x] Styles authoring (styles.xml, docDefaults), style-aware formatting reads (PR #18)

All queued milestones are complete: the library now covers the measured production feature
set at python-docx parity for creating and editing real documents.

Versioning: 0.x during the milestone sequence; API may break between minors until 1.0.

## Licensing

Dual MIT OR Apache-2.0 (Rust ecosystem convention). python-docx (MIT) is the API
inspiration; no code is translated from it — this is a clean-room reimplementation of
the concepts. All dependencies are MIT/Apache-2.0; no copyleft anywhere in the graph.
