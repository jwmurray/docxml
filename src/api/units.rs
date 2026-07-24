//! Value types for character and paragraph formatting: [`Pt`] (font size), [`RgbColor`],
//! [`Alignment`], and [`Length`] (page geometry).
//!
//! Each type owns the conversion between its Rust form and the WordprocessingML
//! attribute string it serializes to, so the run and paragraph accessors stay free of
//! format trivia. The conversions match python-docx's semantics (see each type).

/// English Metric Units per inch — the base of the OOXML measurement system.
const EMU_PER_INCH: i64 = 914_400;
/// EMU per point: `914400 / 72`.
const EMU_PER_PT: i64 = 12_700;
/// EMU per twip (twentieth of a point): `914400 / 1440`.
const EMU_PER_TWIP: i64 = 635;
/// EMU per centimetre: `914400 / 2.54`.
const EMU_PER_CM: i64 = 360_000;

/// A length, stored internally as [English Metric Units][emu] (EMU) — a signed integer
/// count so no precision is lost between unit systems.
///
/// The OOXML measurement hierarchy is `1 inch = 914400 EMU = 1440 twips = 72 pt`
/// (and `1 cm = 360000 EMU`). Construct a `Length` in whichever unit is natural and read
/// it back in any other:
///
/// ```rust
/// use docxml::Length;
///
/// let m = Length::from_inches(1.25);
/// assert_eq!(m.twips(), 1800); // page-geometry attributes are in twips
/// assert_eq!(m.emu(), 1_143_000);
/// assert_eq!(Length::from_twips(1440), Length::from_inches(1.0));
/// ```
///
/// Page-geometry XML attributes (`w:pgSz`, `w:pgMar`) are expressed in twips; the
/// [`Section`](crate::Section) accessors read and write them through this type.
///
/// # Naming
///
/// Constructors are `from_*` and accessors are the plain unit name — the same split
/// [`std::time::Duration`] uses (`Duration::from_secs` / `Duration::as_secs`). Rust has
/// no name overloading, so a single type cannot expose both a `twips(i64)` constructor
/// *and* a `twips(&self)` accessor; the `from_` prefix keeps both directions available
/// and unambiguous.
///
/// [emu]: https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Length(i64);

impl Length {
    /// A length of exactly `emu` English Metric Units.
    pub const fn from_emu(emu: i64) -> Length {
        Length(emu)
    }

    /// A length of `twips` (twentieths of a point) — the unit of page-geometry attributes.
    pub const fn from_twips(twips: i64) -> Length {
        Length(twips * EMU_PER_TWIP)
    }

    /// A length of `inches`.
    pub fn from_inches(inches: f64) -> Length {
        Length((inches * EMU_PER_INCH as f64).round() as i64)
    }

    /// A length of `points`.
    pub fn from_pt(points: f64) -> Length {
        Length((points * EMU_PER_PT as f64).round() as i64)
    }

    /// A length of `centimetres`.
    pub fn from_cm(centimetres: f64) -> Length {
        Length((centimetres * EMU_PER_CM as f64).round() as i64)
    }

    /// This length in English Metric Units (the exact stored value).
    pub const fn emu(self) -> i64 {
        self.0
    }

    /// This length in twips, rounded to the nearest whole twip.
    pub fn twips(self) -> i64 {
        (self.0 as f64 / EMU_PER_TWIP as f64).round() as i64
    }

    /// This length in inches.
    pub fn inches(self) -> f64 {
        self.0 as f64 / EMU_PER_INCH as f64
    }

    /// This length in points.
    pub fn pt(self) -> f64 {
        self.0 as f64 / EMU_PER_PT as f64
    }

    /// This length in centimetres.
    pub fn cm(self) -> f64 {
        self.0 as f64 / EMU_PER_CM as f64
    }

    /// Parse a twips-valued XML attribute (e.g. `w:pgSz/@w:w`) into a `Length`. Returns
    /// `None` when the value is not an integer.
    pub(crate) fn from_twips_str(val: &str) -> Option<Length> {
        val.trim().parse::<i64>().ok().map(Length::from_twips)
    }

    /// This length as a twips XML-attribute string (rounded to the nearest twip).
    pub(crate) fn to_twips_string(self) -> String {
        self.twips().to_string()
    }
}

/// A measurement in points, used for font size (`w:sz`).
///
/// WordprocessingML stores font size in *half-points*, so [`Pt`] serializes to `w:val`
/// as `round(points * 2)` — e.g. `Pt(14.0)` → `"28"`, `Pt(10.5)` → `"21"`. Reading is
/// tolerant: an integer or decimal half-point string parses back to points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Pt(pub f64);

impl Pt {
    /// This size as a `w:sz`/`w:szCs` `w:val` string: half-points, rounded to the
    /// nearest integer.
    pub(crate) fn to_half_points_string(self) -> String {
        (self.0 * 2.0).round().to_string()
    }

    /// Parse a `w:sz` `w:val` (half-points, integer or decimal) back into points.
    /// Returns `None` when the value is not a number.
    pub(crate) fn from_half_points_str(val: &str) -> Option<Pt> {
        val.trim().parse::<f64>().ok().map(|hp| Pt(hp / 2.0))
    }
}

/// An sRGB color (`w:color`), three 8-bit channels.
///
/// Serializes to six uppercase hex digits (`RgbColor(0x1F, 0x4E, 0x79)` → `"1F4E79"`).
/// Parsing accepts any case; `w:val="auto"` is read as `None` (no explicit color), per
/// python-docx.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RgbColor(pub u8, pub u8, pub u8);

impl RgbColor {
    /// This color as a `w:color` `w:val` string: six uppercase hex digits.
    pub(crate) fn to_hex(self) -> String {
        format!("{:02X}{:02X}{:02X}", self.0, self.1, self.2)
    }

    /// Parse a `w:color` `w:val`. Accepts any case; `"auto"` (case-insensitive) and any
    /// non-six-hex-digit value read as `None`.
    pub(crate) fn from_hex(val: &str) -> Option<RgbColor> {
        let val = val.trim();
        if val.eq_ignore_ascii_case("auto") || val.len() != 6 {
            return None;
        }
        let r = u8::from_str_radix(&val[0..2], 16).ok()?;
        let g = u8::from_str_radix(&val[2..4], 16).ok()?;
        let b = u8::from_str_radix(&val[4..6], 16).ok()?;
        Some(RgbColor(r, g, b))
    }
}

/// Paragraph horizontal alignment (`w:jc`).
///
/// Maps to `w:val` `left` / `center` / `right` / `both`. Parsing also accepts the strict
/// aliases `start` (→ [`Left`](Alignment::Left)) and `end` (→ [`Right`](Alignment::Right));
/// unrecognized values read as `None`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    /// Left-aligned (`w:val="left"`).
    Left,
    /// Centered (`w:val="center"`).
    Center,
    /// Right-aligned (`w:val="right"`).
    Right,
    /// Justified (`w:val="both"`).
    Justify,
}

impl Alignment {
    /// This alignment as a `w:jc` `w:val` string.
    pub(crate) fn to_val(self) -> &'static str {
        match self {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
            Alignment::Justify => "both",
        }
    }

    /// Parse a `w:jc` `w:val`. Recognizes `left`/`center`/`right`/`both` plus the strict
    /// aliases `start`→`Left` and `end`→`Right`; anything else is `None`.
    pub(crate) fn from_val(val: &str) -> Option<Alignment> {
        match val.trim() {
            "left" | "start" => Some(Alignment::Left),
            "center" => Some(Alignment::Center),
            "right" | "end" => Some(Alignment::Right),
            "both" => Some(Alignment::Justify),
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pt_serializes_to_half_points() {
        assert_eq!(Pt(14.0).to_half_points_string(), "28");
        assert_eq!(Pt(10.5).to_half_points_string(), "21");
        assert_eq!(Pt(11.0).to_half_points_string(), "22");
    }

    #[test]
    fn pt_parses_integer_and_decimal() {
        assert_eq!(Pt::from_half_points_str("28"), Some(Pt(14.0)));
        assert_eq!(Pt::from_half_points_str("28.0"), Some(Pt(14.0)));
        assert_eq!(Pt::from_half_points_str("21"), Some(Pt(10.5)));
        assert_eq!(Pt::from_half_points_str("nan-ish"), None);
    }

    #[test]
    fn rgb_hex_roundtrip_and_case() {
        assert_eq!(RgbColor(0x1F, 0x4E, 0x79).to_hex(), "1F4E79");
        assert_eq!(
            RgbColor::from_hex("1f4e79"),
            Some(RgbColor(0x1F, 0x4E, 0x79))
        );
        assert_eq!(
            RgbColor::from_hex("1F4E79"),
            Some(RgbColor(0x1F, 0x4E, 0x79))
        );
    }

    #[test]
    fn rgb_auto_and_garbage_read_none() {
        assert_eq!(RgbColor::from_hex("auto"), None);
        assert_eq!(RgbColor::from_hex("AUTO"), None);
        assert_eq!(RgbColor::from_hex("xyz"), None);
        assert_eq!(RgbColor::from_hex("12345"), None);
    }

    #[test]
    fn alignment_val_mapping() {
        assert_eq!(Alignment::Justify.to_val(), "both");
        assert_eq!(Alignment::from_val("both"), Some(Alignment::Justify));
        assert_eq!(Alignment::from_val("start"), Some(Alignment::Left));
        assert_eq!(Alignment::from_val("end"), Some(Alignment::Right));
        assert_eq!(Alignment::from_val("distribute"), None);
    }

    #[test]
    fn length_unit_conversions_are_exact_at_the_hierarchy() {
        let inch = Length::from_inches(1.0);
        assert_eq!(inch.emu(), 914_400);
        assert_eq!(inch.twips(), 1440);
        assert_eq!(inch.pt(), 72.0);
        assert_eq!(inch, Length::from_twips(1440));
        assert_eq!(inch, Length::from_pt(72.0));
        assert_eq!(inch, Length::from_emu(914_400));
    }

    #[test]
    fn length_inch_and_a_quarter_is_1800_twips() {
        let m = Length::from_inches(1.25);
        assert_eq!(m.twips(), 1800);
        assert_eq!(m.emu(), 1_143_000);
    }

    #[test]
    fn length_cm_conversion() {
        assert_eq!(Length::from_cm(2.54).emu(), 914_400);
        assert!((Length::from_emu(360_000).cm() - 1.0).abs() < 1e-9);
    }

    #[test]
    fn length_twips_string_roundtrip() {
        assert_eq!(Length::from_twips(12240).to_twips_string(), "12240");
        assert_eq!(
            Length::from_twips_str("1800"),
            Some(Length::from_twips(1800))
        );
        assert_eq!(
            Length::from_twips_str("  720 "),
            Some(Length::from_twips(720))
        );
        assert_eq!(Length::from_twips_str("auto"), None);
    }
}
