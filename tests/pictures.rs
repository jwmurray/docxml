//! Inline-picture tests: reading real python-docx pictures, creating new ones with EMU
//! geometry, and the package plumbing (media part, content type, relationship) that goes
//! with them. See the milestone-7 fidelity requirements in `docs/DESIGN.md`.

use std::io::Cursor;

use docxml::opc::Package;
use docxml::{Document, Length};

const HYPERLINKS_IMAGES: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/fixtures/hyperlinks_images.docx"
);

/// Serialize a document to an in-memory `.docx` and return the bytes.
fn to_bytes(doc: &Document) -> Vec<u8> {
    let mut buf = Vec::new();
    doc.write(Cursor::new(&mut buf)).unwrap();
    buf
}

/// Read a document back from in-memory `.docx` bytes.
fn from_bytes(bytes: &[u8]) -> Document {
    Document::read(Cursor::new(bytes.to_vec())).unwrap()
}

// --- Test 1: read existing pictures ------------------------------------------------

#[test]
fn reads_inline_picture_from_fixture() {
    let doc = Document::open(HYPERLINKS_IMAGES).unwrap();
    let pics = doc.inline_pictures();
    assert!(
        !pics.is_empty(),
        "hyperlinks_images.docx should have at least one inline picture"
    );
    let pic = pics[0];
    assert!(pic.width(&doc).emu() > 0, "width should be positive EMU");
    assert!(pic.height(&doc).emu() > 0, "height should be positive EMU");
    // The fixture image was inserted at width = 1 inch (see the generator).
    assert_eq!(pic.width(&doc), Length::from_inches(1.0));
}

// --- Test 2: create at native size, full round trip --------------------------------

#[test]
fn add_native_size_picture_roundtrips_with_all_plumbing() {
    let png = make_png(4, 4);

    let mut doc = Document::new();
    let paras_before = doc.paragraphs().len();
    let pic = doc.add_picture(&png, "red.png", None, None).unwrap();

    // Native size: 4 px at the default 96 dpi == 4/96 inch.
    let expected = Length::from_emu(4 * 914_400 / 96);
    assert_eq!(pic.width(&doc), expected);
    assert_eq!(pic.height(&doc), expected);

    // A new paragraph carrying the drawing was appended.
    assert_eq!(doc.paragraphs().len(), paras_before + 1);

    // Save → reopen and check the picture survives.
    let bytes = to_bytes(&doc);
    let reopened = from_bytes(&bytes);
    let pics = reopened.inline_pictures();
    assert_eq!(pics.len(), 1);
    assert_eq!(pics[0].width(&reopened), expected);
    assert_eq!(pics[0].height(&reopened), expected);

    // Package-level plumbing: media part, content type, and the image relationship.
    let pkg = Package::read(Cursor::new(bytes.clone())).unwrap();
    assert!(
        pkg.part("word/media/image1.png").is_some(),
        "media part word/media/image1.png must exist"
    );

    let ct = pkg.part("[Content_Types].xml").unwrap();
    let ct_xml = std::str::from_utf8(&ct.data).unwrap();
    assert!(
        ct_xml.contains(r#"Extension="png""#) && ct_xml.contains(r#"ContentType="image/png""#),
        "content types must declare the png Default: {ct_xml}"
    );

    let rels = pkg.part("word/_rels/document.xml.rels").unwrap();
    let rels_xml = std::str::from_utf8(&rels.data).unwrap();
    assert!(
        rels_xml.contains("relationships/image") && rels_xml.contains("media/image1.png"),
        "document rels must contain the image relationship: {rels_xml}"
    );

    // Fidelity: an untouched open → save of the created file is byte-identical per part.
    let saved_again = {
        let d = from_bytes(&bytes);
        to_bytes(&d)
    };
    let a = Package::read(Cursor::new(bytes)).unwrap();
    let b = Package::read(Cursor::new(saved_again)).unwrap();
    let a_names: Vec<_> = a.parts().iter().map(|p| p.name.as_str()).collect();
    let b_names: Vec<_> = b.parts().iter().map(|p| p.name.as_str()).collect();
    assert_eq!(
        a_names, b_names,
        "part names/order changed on untouched round trip"
    );
    for (x, y) in a.parts().iter().zip(b.parts()) {
        assert_eq!(
            x.data, y.data,
            "part {} changed on untouched round trip",
            x.name
        );
    }
}

// --- Test 3: width only → aspect-preserving height ---------------------------------

#[test]
fn width_only_scales_height_by_aspect_ratio() {
    // A 2:1 image (width : height). Requesting a width should halve it for the height.
    let png = make_png(4, 2);
    let mut doc = Document::new();
    let width = Length::from_inches(2.0);
    let pic = doc
        .add_picture(&png, "wide.png", Some(width), None)
        .unwrap();

    assert_eq!(pic.width(&doc), width);
    assert_eq!(
        pic.height(&doc),
        Length::from_inches(1.0),
        "2:1 image at 2in wide should be 1in tall"
    );
}

// --- Test 4: set_width / set_height touch both extents ------------------------------

#[test]
fn set_width_and_height_update_extent_and_xfrm() {
    let png = make_png(10, 10);
    let mut doc = Document::new();
    let pic = doc.add_picture(&png, "sq.png", None, None).unwrap();

    let new_w = Length::from_inches(3.0);
    let new_h = Length::from_inches(1.5);
    pic.set_width(&mut doc, new_w);
    pic.set_height(&mut doc, new_h);

    assert_eq!(pic.width(&doc), new_w);
    assert_eq!(pic.height(&doc), new_h);

    // Both the wp:extent and the a:xfrm/a:ext must carry the new values. Check the raw XML.
    let bytes = to_bytes(&doc);
    let pkg = Package::read(Cursor::new(bytes)).unwrap();
    let xml = std::str::from_utf8(&pkg.part("word/document.xml").unwrap().data).unwrap();

    let cx = new_w.emu().to_string();
    let cy = new_h.emu().to_string();
    // wp:extent with both new dimensions.
    assert!(
        xml.contains(&format!(r#"<wp:extent cx="{cx}" cy="{cy}""#)),
        "wp:extent not updated: {xml}"
    );
    // a:ext with both new dimensions (inside a:xfrm).
    assert!(
        xml.contains(&format!(r#"<a:ext cx="{cx}" cy="{cy}""#)),
        "a:xfrm/a:ext not updated: {xml}"
    );
}

// --- Test 5: two pictures → distinct ids and media names ---------------------------

#[test]
fn two_pictures_get_distinct_rids_media_and_docpr_ids() {
    let png = make_png(4, 4);
    let mut doc = Document::new();
    doc.add_picture(&png, "a.png", None, None).unwrap();
    doc.add_picture(&png, "b.png", None, None).unwrap();

    let bytes = to_bytes(&doc);
    let pkg = Package::read(Cursor::new(bytes)).unwrap();

    // Distinct media parts.
    assert!(pkg.part("word/media/image1.png").is_some());
    assert!(pkg.part("word/media/image2.png").is_some());

    let doc_xml = std::str::from_utf8(&pkg.part("word/document.xml").unwrap().data).unwrap();
    let rels_xml =
        std::str::from_utf8(&pkg.part("word/_rels/document.xml.rels").unwrap().data).unwrap();

    // Two image relationships with distinct rIds, both referenced by a:blip r:embed.
    let rid_count = rels_xml.matches("relationships/image").count();
    assert_eq!(rid_count, 2, "expected two image relationships: {rels_xml}");

    let embeds: Vec<&str> = collect_attr(doc_xml, "r:embed=\"");
    assert_eq!(embeds.len(), 2, "expected two a:blip r:embed refs");
    assert_ne!(embeds[0], embeds[1], "the two rIds must differ");

    // Distinct docPr ids.
    let docpr_ids: Vec<&str> = collect_attr(doc_xml, "<wp:docPr id=\"");
    assert_eq!(docpr_ids.len(), 2);
    assert_ne!(docpr_ids[0], docpr_ids[1], "docPr ids must differ");

    // Both rIds actually resolve to a media target in the rels.
    for rid in &embeds {
        assert!(
            rels_xml.contains(&format!(r#"Id="{rid}""#)),
            "embed {rid} has no matching relationship"
        );
    }
}

/// Collect the values that follow each occurrence of `prefix` up to the next `"`.
fn collect_attr<'a>(xml: &'a str, prefix: &str) -> Vec<&'a str> {
    let mut out = Vec::new();
    let mut rest = xml;
    while let Some(i) = rest.find(prefix) {
        let after = &rest[i + prefix.len()..];
        let end = after.find('"').unwrap();
        out.push(&after[..end]);
        rest = &after[end..];
    }
    out
}

// --- PNG builder (no external deps) ------------------------------------------------

/// Build a minimal, valid truecolor PNG of the given pixel size (solid red), hand-rolling
/// the chunks + CRC + a stored-block zlib stream — the same approach as
/// `tests/fixtures/generators/hyperlinks_images.py`, in Rust.
fn make_png(width: u32, height: u32) -> Vec<u8> {
    let mut out = vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a];

    let mut ihdr = Vec::new();
    ihdr.extend_from_slice(&width.to_be_bytes());
    ihdr.extend_from_slice(&height.to_be_bytes());
    ihdr.extend_from_slice(&[8, 2, 0, 0, 0]); // 8-bit, truecolor, no filter/interlace
    out.extend(chunk(b"IHDR", &ihdr));

    // One scanline: a filter-type byte (0) followed by `width` red pixels.
    let mut row = vec![0u8];
    for _ in 0..width {
        row.extend_from_slice(&[220, 20, 20]);
    }
    let mut raw = Vec::with_capacity(row.len() * height as usize);
    for _ in 0..height {
        raw.extend_from_slice(&row);
    }
    out.extend(chunk(b"IDAT", &zlib_store(&raw)));
    out.extend(chunk(b"IEND", &[]));
    out
}

fn chunk(tag: &[u8], data: &[u8]) -> Vec<u8> {
    let mut body = Vec::with_capacity(tag.len() + data.len());
    body.extend_from_slice(tag);
    body.extend_from_slice(data);

    let mut out = Vec::new();
    out.extend_from_slice(&(data.len() as u32).to_be_bytes());
    out.extend_from_slice(&body);
    out.extend_from_slice(&crc32(&body).to_be_bytes());
    out
}

/// A zlib stream carrying `data` in uncompressed (stored) DEFLATE blocks.
fn zlib_store(data: &[u8]) -> Vec<u8> {
    let mut out = vec![0x78, 0x01]; // zlib header
    let mut remaining = data;
    loop {
        let block = remaining.len().min(0xFFFF);
        let is_final = block == remaining.len();
        out.push(if is_final { 1 } else { 0 }); // BFINAL + BTYPE=00 (stored)
        out.extend_from_slice(&(block as u16).to_le_bytes());
        out.extend_from_slice(&(!(block as u16)).to_le_bytes());
        out.extend_from_slice(&remaining[..block]);
        remaining = &remaining[block..];
        if is_final {
            break;
        }
    }
    out.extend_from_slice(&adler32(data).to_be_bytes());
    out
}

fn adler32(data: &[u8]) -> u32 {
    let (mut a, mut b) = (1u32, 0u32);
    for &byte in data {
        a = (a + byte as u32) % 65521;
        b = (b + a) % 65521;
    }
    (b << 16) | a
}

fn crc32(data: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for &byte in data {
        crc ^= byte as u32;
        for _ in 0..8 {
            crc = if crc & 1 != 0 {
                (crc >> 1) ^ 0xEDB8_8320
            } else {
                crc >> 1
            };
        }
    }
    crc ^ 0xFFFF_FFFF
}
