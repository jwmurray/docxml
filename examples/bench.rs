//! Micro-benchmark of the three hot paths a caller batch-processing `.docx`
//! files exercises: opening a package, a full open+save round trip, and
//! parsing the main document XML.
//!
//! No criterion, no extra dependencies — plain `std::time::Instant`, run
//! straight from the crate's own dev-dependencies. Each op is timed N times
//! and reported as the median (not mean) of the individual timings, in
//! microseconds, so a handful of outliers (first-run page faults, GC-less-Rust
//! notwithstanding, OS scheduling noise) don't skew the number callers care
//! about: the typical cost of one call.
//!
//! Run with:
//!     cargo run --example bench --release

use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use docxml::opc::Package;
use docxml::xml::XmlTree;

/// Iterations per operation. 200 keeps total run time well under a second per
/// fixture while giving a stable median.
const N: usize = 200;

fn main() {
    let fixtures = all_fixtures();
    println!(
        "docxml bench — {} fixture(s), N={N} per operation, reporting median microseconds/op\n",
        fixtures.len()
    );

    println!(
        "{:<22} {:>14} {:>18} {:>16}",
        "fixture", "open (us)", "open+save (us)", "xml parse (us)"
    );
    println!("{}", "-".repeat(72));

    for fixture in &fixtures {
        let name = fixture.file_name().unwrap().to_string_lossy();
        let open_us = bench_open(fixture);
        let roundtrip_us = bench_open_save_roundtrip(fixture);
        let parse_us = bench_xml_parse(fixture);
        println!("{name:<22} {open_us:>14.2} {roundtrip_us:>18.2} {parse_us:>16.2}");
    }
}

/// Every `.docx` fixture in `tests/fixtures/`, sorted for stable output.
fn all_fixtures() -> Vec<PathBuf> {
    let dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut fixtures: Vec<PathBuf> = std::fs::read_dir(&dir)
        .unwrap_or_else(|e| panic!("reading {}: {e}", dir.display()))
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .filter(|p| p.extension().is_some_and(|ext| ext == "docx"))
        .collect();
    fixtures.sort();
    fixtures
}

/// Median of a slice of durations, in microseconds. `durations` is sorted
/// in place; for an even count the median is the mean of the two middle
/// values, matching the usual definition.
fn median_micros(durations: &mut [Duration]) -> f64 {
    durations.sort();
    let len = durations.len();
    let mid = len / 2;
    let micros = |d: Duration| d.as_secs_f64() * 1_000_000.0;
    if len % 2 == 0 {
        (micros(durations[mid - 1]) + micros(durations[mid])) / 2.0
    } else {
        micros(durations[mid])
    }
}

/// (a) `Package::open` from disk, N times.
fn bench_open(fixture: &Path) -> f64 {
    let mut durations = Vec::with_capacity(N);
    for _ in 0..N {
        let start = Instant::now();
        let pkg = Package::open(fixture).unwrap();
        durations.push(start.elapsed());
        std::hint::black_box(&pkg);
    }
    median_micros(&mut durations)
}

/// (b) Full open+save round trip, N times. Saves to an in-memory buffer
/// (`Cursor<Vec<u8>>`) rather than disk so the timing reflects the library's
/// own packaging cost rather than filesystem noise.
fn bench_open_save_roundtrip(fixture: &Path) -> f64 {
    let mut durations = Vec::with_capacity(N);
    for _ in 0..N {
        let start = Instant::now();
        let pkg = Package::open(fixture).unwrap();
        let mut buf = Cursor::new(Vec::new());
        pkg.write(&mut buf).unwrap();
        durations.push(start.elapsed());
        std::hint::black_box(&buf);
    }
    median_micros(&mut durations)
}

/// (c) `XmlTree::parse` of `word/document.xml`, N times. The part bytes are
/// read once outside the timing loop so only parsing is measured.
fn bench_xml_parse(fixture: &Path) -> f64 {
    let pkg = Package::open(fixture).unwrap();
    let doc = pkg.main_document_part().unwrap();
    let bytes = doc.data.clone();

    let mut durations = Vec::with_capacity(N);
    for _ in 0..N {
        let start = Instant::now();
        let tree = XmlTree::parse(&bytes).unwrap();
        durations.push(start.elapsed());
        std::hint::black_box(&tree);
    }
    median_micros(&mut durations)
}
