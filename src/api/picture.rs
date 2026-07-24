//! Inline pictures: reading and creating DrawingML inline images.
//!
//! An inline picture is a `w:drawing` wrapping a `wp:inline` shape — the run-level image
//! form python-docx creates with `Document.add_picture` and reads with
//! `Document.inline_shapes`. [`Document::add_picture`] adds one (creating the media part,
//! its content-type registration, and a package relationship), and
//! [`Document::inline_pictures`] reads every one back. Geometry is in
//! [EMU](crate::Length); the sizing rules mirror python-docx (see [`Document::add_picture`]).

use crate::error::{Error, Result};
use crate::xml::{NodeId, XmlTree};

use super::{Document, Length, PartId, is_element_in};

/// `wp` — WordprocessingDrawing namespace (transitional + strict), for matching `wp:*`.
const WP_URIS: [&str; 2] = [
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
];
/// `a` — DrawingML main namespace (transitional + strict), for matching `a:*`.
const A_URIS: [&str; 2] = [
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/main",
];

// Canonical (transitional) URIs used when *creating* the inline shape. They are declared
// as `xmlns:*` attributes directly on the `wp:inline` element (see `build_inline`) rather
// than assumed at the document root, so the shape is self-contained even in a document
// whose root does not declare `a`/`pic` (the embedded blank template declares neither).
const WP_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_URI: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// Relationship type for an embedded image part.
const IMAGE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// A lightweight handle to an inline picture — a `wp:inline` shape inside a `w:drawing`.
///
/// Like the other handles, `Picture` is `Copy` and borrows nothing: it is the arena node
/// id of the `wp:inline` element plus the id of the part it lives in. Pass a [`Document`]
/// back to it to read or set the picture's geometry.
///
/// # Geometry
///
/// [`width`](Self::width) / [`height`](Self::height) read the `wp:extent` display size in
/// [EMU](crate::Length). [`set_width`](Self::set_width) / [`set_height`](Self::set_height)
/// update *both* the `wp:extent` and the `pic:spPr/a:xfrm/a:ext` extents (python-docx keeps
/// the two in lock-step), preserving the other axis unless it is also set.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Picture {
    part: PartId,
    node: NodeId,
}

impl Picture {
    /// Wrap a known-`wp:inline` node id living in `part`.
    pub(crate) fn from_node(part: PartId, node: NodeId) -> Self {
        Picture { part, node }
    }

    /// The picture's underlying `wp:inline` tree node id.
    pub fn node(&self) -> NodeId {
        self.node
    }

    /// The picture's display width, from `wp:extent/@cx` (EMU). Defaults to a zero length
    /// if the extent is somehow absent.
    pub fn width(&self, doc: &Document) -> Length {
        self.extent_dim(doc, "cx")
    }

    /// The picture's display height, from `wp:extent/@cy` (EMU).
    pub fn height(&self, doc: &Document) -> Length {
        self.extent_dim(doc, "cy")
    }

    /// Set the display width, updating both the `wp:extent` and the `a:xfrm/a:ext` `cx`
    /// (height preserved).
    pub fn set_width(&self, doc: &mut Document, width: Length) -> Picture {
        self.set_dim(doc, "cx", width);
        *self
    }

    /// Set the display height, updating both the `wp:extent` and the `a:xfrm/a:ext` `cy`
    /// (width preserved).
    pub fn set_height(&self, doc: &mut Document, height: Length) -> Picture {
        self.set_dim(doc, "cy", height);
        *self
    }

    /// The direct `wp:extent` child, if present.
    fn extent(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.children(self.node)
            .iter()
            .copied()
            .find(|&c| is_element_in(tree, c, &WP_URIS, "extent"))
    }

    /// The `pic:spPr/a:xfrm/a:ext` extent (first `a:ext` descendant), if present.
    fn a_ext(&self, tree: &XmlTree) -> Option<NodeId> {
        tree.descendants(self.node)
            .find(|&d| is_element_in(tree, d, &A_URIS, "ext"))
    }

    /// Read one axis (`cx`/`cy`) of the `wp:extent` as a [`Length`].
    fn extent_dim(&self, doc: &Document, attr: &str) -> Length {
        let tree = doc.tree(self.part);
        self.extent(tree)
            .and_then(|e| tree.attr(e, attr))
            .and_then(|v| v.trim().parse::<i64>().ok())
            .map(Length::from_emu)
            .unwrap_or(Length::from_emu(0))
    }

    /// Set one axis (`cx`/`cy`) on both the `wp:extent` and the `a:xfrm/a:ext`.
    fn set_dim(&self, doc: &mut Document, attr: &str, len: Length) {
        let value = len.emu().to_string();
        if let Some(extent) = self.extent(doc.tree(self.part)) {
            doc.tree_mut(self.part)
                .set_attr(extent, attr, value.clone());
        }
        if let Some(a_ext) = self.a_ext(doc.tree(self.part)) {
            doc.tree_mut(self.part).set_attr(a_ext, attr, value);
        }
    }
}

impl Document {
    /// Every inline picture in the main document part, in document order.
    ///
    /// Matches every `wp:inline` shape (namespace-correct: the wordprocessingDrawing URI,
    /// transitional or strict), mirroring python-docx's `Document.inline_shapes` restricted
    /// to pictures. Shapes in headers/footers are not included at this milestone.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use docxml::Document;
    ///
    /// let doc = Document::open("with-image.docx")?;
    /// for pic in doc.inline_pictures() {
    ///     println!("{} x {} EMU", pic.width(&doc).emu(), pic.height(&doc).emu());
    /// }
    /// # Ok::<(), docxml::Error>(())
    /// ```
    pub fn inline_pictures(&self) -> Vec<Picture> {
        let tree = self.tree(PartId::MAIN);
        tree.descendants(tree.root())
            .filter(|&d| is_element_in(tree, d, &WP_URIS, "inline"))
            .map(|d| Picture::from_node(PartId::MAIN, d))
            .collect()
    }

    /// Add an inline picture and return a handle to it.
    ///
    /// The image bytes are sniffed for format (PNG or JPEG) and pixel dimensions; the
    /// display size follows python-docx's rules:
    ///
    /// - neither `width` nor `height`: the native size, `pixels / dpi` inches (dpi read
    ///   from the image, defaulting to 96);
    /// - one of the two: that value, with the other scaled to preserve the aspect ratio;
    /// - both: exactly those values.
    ///
    /// Adding a picture creates the media part (`word/media/imageN.<ext>`), registers the
    /// extension's content type in `[Content_Types].xml`, adds an image relationship from
    /// the document part, and appends a new paragraph whose single run carries the
    /// `w:drawing`. `filename_hint` names the `pic:cNvPr` shape (python-docx uses the
    /// source filename); it does not affect the media part name.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Image`] when the bytes are not a supported (PNG/JPEG) image.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// use docxml::{Document, Length};
    ///
    /// let mut doc = Document::new();
    /// let png: &[u8] = /* image bytes */ &[];
    /// let pic = doc.add_picture(png, "logo.png", Some(Length::from_inches(2.0)), None)?;
    /// # Ok::<(), docxml::Error>(())
    /// ```
    pub fn add_picture(
        &mut self,
        image: &[u8],
        filename_hint: &str,
        width: Option<Length>,
        height: Option<Length>,
    ) -> Result<Picture> {
        let info = sniff_image(image)?;

        // Native size in EMU, then apply the python-docx scaling rules.
        let native_w = px_to_emu(info.px_width, info.horz_dpi);
        let native_h = px_to_emu(info.px_height, info.vert_dpi);
        let (cx, cy) = scaled_dimensions(native_w, native_h, width, height);

        // Media part + content-type registration + relationship.
        let media_name = self.next_media_name(info.ext);
        self.add_part(media_name.clone(), image.to_vec());
        self.ensure_content_type_default(info.ext, info.content_type)?;

        let source = self.main_part_name().to_string();
        let dir = match source.rfind('/') {
            Some(i) => &source[..=i],
            None => "",
        };
        let target = media_name
            .strip_prefix(dir)
            .unwrap_or(&media_name)
            .to_string();
        let rid = self.add_relationship(&source, IMAGE_REL_TYPE, &target, false)?;

        let docpr_id = self.next_docpr_id();
        let shape_name = if filename_hint.is_empty() {
            "image"
        } else {
            filename_hint
        };

        // New paragraph → run → w:drawing → wp:inline (built with explicit xmlns decls).
        let para = self.add_paragraph("");
        let p_node = para.node();
        let r_name = self.qn(PartId::MAIN, "r");
        let drawing_name = self.qn(PartId::MAIN, "drawing");

        let tree = self.tree_mut(PartId::MAIN);
        let run = tree.create_element(r_name);
        let drawing = tree.create_element(drawing_name);
        let inline = build_inline(tree, cx, cy, docpr_id, shape_name, &rid);
        tree.append_child(drawing, inline);
        tree.append_child(run, drawing);
        tree.append_child(p_node, run);

        Ok(Picture::from_node(PartId::MAIN, inline))
    }

    /// The next free `word/media/imageN.<ext>` name: `max(existing N) + 1`, over parts
    /// already named `word/media/image<digits>.*`.
    fn next_media_name(&self, ext: &str) -> String {
        let mut max = 0u32;
        for part in self.package().parts() {
            if let Some(rest) = part.name.strip_prefix("word/media/image") {
                let digits: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
                if let Ok(n) = digits.parse::<u32>() {
                    max = max.max(n);
                }
            }
        }
        format!("word/media/image{}.{ext}", max + 1)
    }

    /// The next free `wp:docPr/@id` in the main document part: `max(existing) + 1`.
    fn next_docpr_id(&self) -> u32 {
        let tree = self.tree(PartId::MAIN);
        let mut max = 0u32;
        for d in tree.descendants(tree.root()) {
            if is_element_in(tree, d, &WP_URIS, "docPr") {
                if let Some(id) = tree
                    .attr(d, "id")
                    .and_then(|v| v.trim().parse::<u32>().ok())
                {
                    max = max.max(id);
                }
            }
        }
        max + 1
    }
}

/// Build a `wp:inline` picture shape subtree and return its node id (detached; the caller
/// attaches it under a `w:drawing`).
///
/// The namespaces `wp`, `a`, `pic`, and `r` are declared as `xmlns:*` attributes on the
/// `wp:inline` element itself, so every descendant resolves without relying on the
/// document root — matching python-docx's `CT_Inline`/`CT_Picture` templates, which carry
/// their own namespace declarations.
fn build_inline(
    tree: &mut XmlTree,
    cx: i64,
    cy: i64,
    docpr_id: u32,
    shape_name: &str,
    rid: &str,
) -> NodeId {
    let inline = tree.create_element("wp:inline");
    tree.set_attr(inline, "xmlns:wp", WP_URI);
    tree.set_attr(inline, "xmlns:a", A_URI);
    tree.set_attr(inline, "xmlns:pic", PIC_URI);
    tree.set_attr(inline, "xmlns:r", R_URI);

    let extent = tree.create_element("wp:extent");
    tree.set_attr(extent, "cx", cx.to_string());
    tree.set_attr(extent, "cy", cy.to_string());
    tree.append_child(inline, extent);

    let docpr = tree.create_element("wp:docPr");
    tree.set_attr(docpr, "id", docpr_id.to_string());
    tree.set_attr(docpr, "name", format!("Picture {docpr_id}"));
    tree.append_child(inline, docpr);

    let cnv_frame = tree.create_element("wp:cNvGraphicFramePr");
    let locks = tree.create_element("a:graphicFrameLocks");
    tree.set_attr(locks, "noChangeAspect", "1");
    tree.append_child(cnv_frame, locks);
    tree.append_child(inline, cnv_frame);

    let graphic = tree.create_element("a:graphic");
    let graphic_data = tree.create_element("a:graphicData");
    tree.set_attr(graphic_data, "uri", PIC_URI);

    let pic = tree.create_element("pic:pic");

    let nv_pic_pr = tree.create_element("pic:nvPicPr");
    let cnv_pr = tree.create_element("pic:cNvPr");
    tree.set_attr(cnv_pr, "id", "0");
    tree.set_attr(cnv_pr, "name", shape_name);
    tree.append_child(nv_pic_pr, cnv_pr);
    let cnv_pic_pr = tree.create_element("pic:cNvPicPr");
    tree.append_child(nv_pic_pr, cnv_pic_pr);
    tree.append_child(pic, nv_pic_pr);

    let blip_fill = tree.create_element("pic:blipFill");
    let blip = tree.create_element("a:blip");
    tree.set_attr(blip, "r:embed", rid);
    tree.append_child(blip_fill, blip);
    let stretch = tree.create_element("a:stretch");
    let fill_rect = tree.create_element("a:fillRect");
    tree.append_child(stretch, fill_rect);
    tree.append_child(blip_fill, stretch);
    tree.append_child(pic, blip_fill);

    let sp_pr = tree.create_element("pic:spPr");
    let xfrm = tree.create_element("a:xfrm");
    let off = tree.create_element("a:off");
    tree.set_attr(off, "x", "0");
    tree.set_attr(off, "y", "0");
    tree.append_child(xfrm, off);
    let a_ext = tree.create_element("a:ext");
    tree.set_attr(a_ext, "cx", cx.to_string());
    tree.set_attr(a_ext, "cy", cy.to_string());
    tree.append_child(xfrm, a_ext);
    tree.append_child(sp_pr, xfrm);
    let prst_geom = tree.create_element("a:prstGeom");
    tree.set_attr(prst_geom, "prst", "rect");
    tree.append_child(sp_pr, prst_geom);
    tree.append_child(pic, sp_pr);

    tree.append_child(graphic_data, pic);
    tree.append_child(graphic, graphic_data);
    tree.append_child(inline, graphic);
    inline
}

/// The display extent for an image, in EMU, applying python-docx's sizing rules:
/// native size when neither dimension is given; the given one plus an aspect-preserving
/// scale of the other when exactly one is given; both verbatim when both are given.
fn scaled_dimensions(
    native_w: i64,
    native_h: i64,
    width: Option<Length>,
    height: Option<Length>,
) -> (i64, i64) {
    match (width, height) {
        (None, None) => (native_w, native_h),
        (Some(w), None) => {
            let cx = w.emu();
            let cy = scale(cx, native_h, native_w).unwrap_or(native_h);
            (cx, cy)
        }
        (None, Some(h)) => {
            let cy = h.emu();
            let cx = scale(cy, native_w, native_h).unwrap_or(native_w);
            (cx, cy)
        }
        (Some(w), Some(h)) => (w.emu(), h.emu()),
    }
}

/// `round(given * other_native / this_native)`, or `None` when `this_native` is zero.
fn scale(given: i64, other_native: i64, this_native: i64) -> Option<i64> {
    if this_native == 0 {
        return None;
    }
    Some((given as f64 * other_native as f64 / this_native as f64).round() as i64)
}

/// Pixels to EMU at `dpi`: `round(px * 914400 / dpi)`.
fn px_to_emu(px: u32, dpi: u32) -> i64 {
    let dpi = if dpi == 0 { 96 } else { dpi };
    (px as f64 * 914_400.0 / dpi as f64).round() as i64
}

/// Sniffed image facts: pixel dimensions, resolution, and the target part format.
struct ImageInfo {
    px_width: u32,
    px_height: u32,
    horz_dpi: u32,
    vert_dpi: u32,
    /// Canonical file extension for the media part (`"png"` / `"jpg"`).
    ext: &'static str,
    /// Content type registered in `[Content_Types].xml`.
    content_type: &'static str,
}

/// Detect the image format from a magic-number prefix and delegate to the format parser.
fn sniff_image(data: &[u8]) -> Result<ImageInfo> {
    if data.starts_with(&[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]) {
        sniff_png(data)
    } else if data.starts_with(&[0xFF, 0xD8]) {
        sniff_jpeg(data)
    } else {
        Err(Error::Image(
            "unsupported image format (only PNG and JPEG are supported)".into(),
        ))
    }
}

/// Read the width/height (IHDR) and optional resolution (pHYs) of a PNG.
fn sniff_png(data: &[u8]) -> Result<ImageInfo> {
    // Signature (8) + IHDR length (4) + "IHDR" (4) + width (4) + height (4) = 24 bytes.
    if data.len() < 24 || &data[12..16] != b"IHDR" {
        return Err(Error::Image("PNG is truncated or missing IHDR".into()));
    }
    let px_width = be_u32(&data[16..20]);
    let px_height = be_u32(&data[20..24]);
    let (horz_dpi, vert_dpi) = png_dpi(data).unwrap_or((96, 96));
    Ok(ImageInfo {
        px_width,
        px_height,
        horz_dpi,
        vert_dpi,
        ext: "png",
        content_type: "image/png",
    })
}

/// The (horz, vert) DPI from a PNG pHYs chunk, if present and expressed in pixels/metre.
fn png_dpi(data: &[u8]) -> Option<(u32, u32)> {
    let mut pos = 8; // past the signature
    while pos + 8 <= data.len() {
        let len = be_u32(&data[pos..pos + 4]) as usize;
        let ctype = &data[pos + 4..pos + 8];
        let body = pos + 8;
        if ctype == b"pHYs" {
            if body + 9 > data.len() {
                return None;
            }
            let x_ppu = be_u32(&data[body..body + 4]);
            let y_ppu = be_u32(&data[body + 4..body + 8]);
            let unit = data[body + 8];
            // unit == 1 means pixels per metre; anything else is aspect-ratio only.
            return (unit == 1).then(|| (ppu_to_dpi(x_ppu), ppu_to_dpi(y_ppu)));
        }
        if ctype == b"IDAT" || ctype == b"IEND" {
            return None; // pHYs precedes IDAT; no point scanning image data
        }
        pos = body + len + 4; // skip chunk data + CRC
    }
    None
}

/// Pixels-per-metre to DPI: `round(ppu * 0.0254)`.
fn ppu_to_dpi(ppu: u32) -> u32 {
    (ppu as f64 * 0.0254).round() as u32
}

/// Read the width/height (SOFn) and optional resolution (JFIF APP0 density) of a JPEG.
fn sniff_jpeg(data: &[u8]) -> Result<ImageInfo> {
    let (mut horz_dpi, mut vert_dpi) = (96u32, 96u32);
    let mut dims: Option<(u32, u32)> = None;

    let mut i = 2; // past the SOI marker (FFD8)
    while i + 1 < data.len() {
        if data[i] != 0xFF {
            i += 1;
            continue;
        }
        // Skip any fill bytes (a run of 0xFF) preceding the marker code.
        let mut j = i + 1;
        while j < data.len() && data[j] == 0xFF {
            j += 1;
        }
        if j >= data.len() {
            break;
        }
        let marker = data[j];
        i = j + 1;

        // Standalone markers with no length/payload.
        if marker == 0x01 || (0xD0..=0xD9).contains(&marker) {
            continue;
        }
        if i + 1 >= data.len() {
            break;
        }
        let seg_len = ((data[i] as usize) << 8) | data[i + 1] as usize;
        if seg_len < 2 || i + seg_len > data.len() {
            break;
        }
        let seg = &data[i + 2..i + seg_len];

        // APP0 JFIF density (comes before the SOF frame header).
        if marker == 0xE0 && seg.len() >= 12 && &seg[0..5] == b"JFIF\0" {
            let units = seg[7];
            let x_density = ((seg[8] as u32) << 8) | seg[9] as u32;
            let y_density = ((seg[10] as u32) << 8) | seg[11] as u32;
            match units {
                1 => {
                    // dots per inch
                    if x_density > 0 {
                        horz_dpi = x_density;
                    }
                    if y_density > 0 {
                        vert_dpi = y_density;
                    }
                }
                2 => {
                    // dots per centimetre
                    if x_density > 0 {
                        horz_dpi = (x_density as f64 * 2.54).round() as u32;
                    }
                    if y_density > 0 {
                        vert_dpi = (y_density as f64 * 2.54).round() as u32;
                    }
                }
                _ => {}
            }
        }

        // SOFn frame header: [precision:1][height:2][width:2] …
        if is_sof_marker(marker) && seg.len() >= 5 {
            let h = ((seg[1] as u32) << 8) | seg[2] as u32;
            let w = ((seg[3] as u32) << 8) | seg[4] as u32;
            dims = Some((w, h));
            break;
        }
        i += seg_len;
    }

    let (px_width, px_height) =
        dims.ok_or_else(|| Error::Image("JPEG has no SOF frame header".into()))?;
    Ok(ImageInfo {
        px_width,
        px_height,
        horz_dpi,
        vert_dpi,
        ext: "jpg",
        content_type: "image/jpeg",
    })
}

/// Whether a JPEG marker is a start-of-frame (SOF0–3, 5–7, 9–11, 13–15) — the ones that
/// carry frame dimensions, excluding DHT (C4), JPG (C8), and DAC (CC).
fn is_sof_marker(m: u8) -> bool {
    matches!(m, 0xC0..=0xC3 | 0xC5..=0xC7 | 0xC9..=0xCB | 0xCD..=0xCF)
}

/// Big-endian `u32` from a 4-byte slice.
fn be_u32(b: &[u8]) -> u32 {
    u32::from_be_bytes([b[0], b[1], b[2], b[3]])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scaled_dimensions_rules() {
        // Native when neither given.
        assert_eq!(scaled_dimensions(400, 200, None, None), (400, 200));
        // Width only: height scales to preserve 2:1 aspect.
        assert_eq!(
            scaled_dimensions(400, 200, Some(Length::from_emu(1000)), None),
            (1000, 500)
        );
        // Height only: width scales.
        assert_eq!(
            scaled_dimensions(400, 200, None, Some(Length::from_emu(500))),
            (1000, 500)
        );
        // Both: verbatim.
        assert_eq!(
            scaled_dimensions(
                400,
                200,
                Some(Length::from_emu(11)),
                Some(Length::from_emu(22))
            ),
            (11, 22)
        );
    }

    #[test]
    fn px_to_emu_at_96_dpi() {
        // 96 px at 96 dpi == 1 inch == 914400 EMU.
        assert_eq!(px_to_emu(96, 96), 914_400);
        assert_eq!(px_to_emu(4, 96), 38_100);
    }

    #[test]
    fn sniff_png_reads_ihdr_and_defaults_dpi() {
        // Signature + a 13-byte IHDR chunk (8x8) + an IDAT chunk header; enough for the
        // sniffer, which reads IHDR and stops the pHYs scan at IDAT. It never inflates.
        let mut png = vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a];
        png.extend_from_slice(&13u32.to_be_bytes());
        png.extend_from_slice(b"IHDR");
        png.extend_from_slice(&8u32.to_be_bytes()); // width
        png.extend_from_slice(&8u32.to_be_bytes()); // height
        png.extend_from_slice(&[8, 2, 0, 0, 0]); // bit depth, color type, etc.
        png.extend_from_slice(&[0, 0, 0, 0]); // (bogus) CRC — never checked
        png.extend_from_slice(&0u32.to_be_bytes());
        png.extend_from_slice(b"IDAT");

        let info = sniff_image(&png).unwrap();
        assert_eq!((info.px_width, info.px_height), (8, 8));
        assert_eq!(info.ext, "png");
        assert_eq!(info.content_type, "image/png");
        assert_eq!((info.horz_dpi, info.vert_dpi), (96, 96)); // no pHYs
    }

    #[test]
    fn sniff_rejects_unknown_format() {
        assert!(sniff_image(b"not an image at all").is_err());
    }
}
