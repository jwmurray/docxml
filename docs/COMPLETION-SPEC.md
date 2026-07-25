# docxml completion spec (2026-07-25)

Owner-approved plan to finish the crate. One branch + PR per item, CI green
before the next; additive API only (docgen_mcp pins this crate by git rev —
breaking changes require a coordinated pin-bump PR in docgen_mcp that runs
its full suite plus docgen-bench compare). Exemplar-driven: every
rendering-visible feature gets a golden derived from filed-PDF measurements
or a cited court rule (the .docx files formerly in docgen_mcp/samples were
firm-generated conversions and were purged — PDFs are the only ground
truth). No firm/vendor/client names in code or fixtures.

## Priorities

1. **Footnote/endnote authoring** — the blocking gap for brief writing
   (python-docx cannot do this either; courts regulate footnotes, e.g. CRC
   2.108 footnote point size). Scope: footnotes.xml/endnotes.xml part
   creation incl. required separator/continuation stub notes; Run-level
   footnote reference marks; per-note paragraphs with style + point-size
   control; read-back accessors (note count, text, point sizes) so
   docgen-bench probes can assert them; author→save→reopen round-trip test.
   Consumer follow-up in docgen_mcp: render a footnote-bearing pleading per
   jurisdiction pack with goldens measured from the footnote-bearing
   exemplar PDFs.
2. **Typed `w:docGrid` (incl. `linePitch`) + typed `w:pgNumType`** — buys
   the exact 28-lines-per-page geometry the AZ/CA exemplars show (see
   docgen_mcp docs/jurisdictions/ca.md §9 and §docxml). Optional pack field
   in docgen-core + a golden proving 28 lines/page against the CA exemplar
   measurements.
3. **Ergonomics** from the CA build: derive `PartialOrd`/`Ord` for `Pt` and
   sibling measurement types; audit element-ordered insertion helpers.
4. **Optional, gated on 1–3 landing green**: comments + tracked-changes
   authoring (redline support); inline→anchored image positioning. Design
   notes in README-checklist style first; implement only with test
   coverage.

## Definition of done

README checklist updated ([x] per landed item); report per priority: API
additions, docgen_mcp pin bumps made, remaining known gaps. Known
OOXML-inherent limitations (decorative evenly-spaced number grids, vertical
separator rules) stay documented in the jurisdiction specs, not worked
around here.

Context: docgen_mcp docs/jurisdictions/*.md, docgen-bench property probes,
verify_mcp docs/ROLLOUT-UT-AZ.md ground rules.
