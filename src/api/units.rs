//! Value types for character and paragraph formatting: [`Pt`] (font size), [`RgbColor`],
//! and [`Alignment`].
//!
//! Each type owns the conversion between its Rust form and the WordprocessingML
//! attribute string it serializes to, so the run and paragraph accessors stay free of
//! format trivia. The conversions match python-docx's semantics (see each type).

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
}
