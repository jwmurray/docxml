# Benchmarks: docxml vs python-docx

No criterion, no external benchmark harness — plain `std::time::Instant` on
the Rust side, `time.perf_counter()` on the Python side. Each operation is
timed N times per fixture; the reported number is the **median** of the
individual timings, not the mean, so a handful of outliers (page faults on
the first touch, OS scheduling noise, etc.) don't skew the number that
matters: the typical cost of one call.

- Rust: `cargo run --example bench --release` — N=200 per operation.
- Python: `uv run --with python-docx python benches/bench_python_docx.py` —
  N=50 per document operation (python-docx is slow enough that 200 iterations
  isn't necessary to get a stable median), plus a separate startup
  measurement (see below).

## Hardware and versions

- Apple Silicon M-series Mac (arm64), macOS 26.5.2
- rustc 1.96.1, cargo 1.96.1, release profile (`opt-level = 3` default)
- Python 3.13.7 (via `uv run --with python-docx`), python-docx 1.2.0
- docxml 0.0.1, dependencies: quick-xml 0.41.0, zip 7.2.0, thiserror 2.0.19

## Results

### docxml (Rust), median microseconds/op, N=200

```
fixture                     open (us)     open+save (us)   xml parse (us)
------------------------------------------------------------------------
basic.docx                     322.29            1644.58            13.33
hyperlinks_images.docx         233.94            1590.19             6.69
styles_toc.docx                236.12            1606.08             6.88
tables_merged.docx             231.52            1584.92            17.62
```

- **open**: `Package::open(path)` — read the zip, extract every part into
  memory. No XML parsing happens yet (parts are parsed lazily, per the
  fidelity contract).
- **open+save**: the above, plus `Package::write` into an in-memory
  `Cursor<Vec<u8>>` (no disk I/O in the save, to isolate packaging cost from
  filesystem noise).
- **xml parse**: `XmlTree::parse` of `word/document.xml`'s bytes alone
  (already in memory), the cost of building the lossless DOM.

`basic.docx` is the outlier on `open` and `xml parse` because it carries a
much larger `word/document.xml` (and a full `stylesWithEffects.xml`) than the
three generated fixtures — more zip entries to inflate and more elements to
walk.

### python-docx, median milliseconds/op, N=50

```
fixture                   open (ms)   open+save (ms)
----------------------------------------------------
basic.docx                    2.251            5.743
hyperlinks_images.docx        1.969            5.675
styles_toc.docx               1.976            5.670
tables_merged.docx            1.949            5.640

interpreter + `import docx` startup: 39.91 ms (median of 10 subprocess runs)
```

- **open**: `docx.Document(path)` — reads the zip and eagerly parses every
  XML part into an lxml tree (python-docx has no lazy parsing).
- **open+save**: the above, plus `document.save()` into an in-memory
  `io.BytesIO`.
- **startup**: median wall time of a fresh `python -c "import docx"`
  subprocess — interpreter boot plus importing python-docx and its lxml/PIL
  dependency chain. Measured separately because it's a one-time,
  per-process cost that would otherwise dominate (and misrepresent) the
  per-document numbers above; docxml has no equivalent since it ships as a
  library linked into the caller's own binary, not a fresh interpreter per run.

## Head-to-head

| Operation | docxml (Rust) | python-docx | Ratio |
|---|---:|---:|---:|
| open, `basic.docx` | 322 µs | 2,251 µs | ~7.0x faster |
| open, generated fixtures (avg) | ~234 µs | ~1,965 µs | ~8.4x faster |
| open+save, `basic.docx` | 1,645 µs | 5,743 µs | ~3.5x faster |
| open+save, generated fixtures (avg) | ~1,594 µs | ~5,662 µs | ~3.6x faster |
| interpreter/import startup | n/a (compiled binary) | 39.9 ms | — |

Two things stand out:

1. **`open` is where docxml's lazy-parsing design pays off most**: python-docx
   parses every part into an lxml tree up front; `Package::open` only unzips
   and hands back raw bytes, so it's roughly an order of magnitude faster.
2. **The gap narrows for open+save** because docxml's round trip still
   re-serializes the zip archive and (for parts actually touched, none here)
   would re-serialize XML — the dominant cost converges to "write a zip",
   which both implementations must do.
3. **Startup dwarfs everything for short-lived scripts.** A one-shot Python
   process pays ~40 ms just to `import docx` before touching a single
   document — more than 100x any single docxml operation above. A caller
   invoking python-docx once per file (e.g. from a shell loop) pays this
   every time; a long-running Python process amortizes it, but docxml (or any
   compiled binary/library) never pays it at all.

## Reproducing

```fish
# Rust side
cargo run --example bench --release

# Python side
uv run --with python-docx python benches/bench_python_docx.py
```

Numbers will vary run to run (typically single-digit percent for the Rust
side, more for Python due to GC and import caching effects) but the relative
ordering and rough magnitude above are stable.
