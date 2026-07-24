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

    /// This size as a twentieths-of-a-point string (the unit of `w:spacing/@w:before`,
    /// `@w:after`, and exact/at-least line values), rounded to the nearest twentieth.
    pub(crate) fn to_twentieths_string(self) -> String {
        ((self.0 * 20.0).round() as i64).to_string()
    }

    /// Parse a twentieths-of-a-point value (integer or decimal) back into points. Returns
    /// `None` when the value is not a number.
    pub(crate) fn from_twentieths_str(val: &str) -> Option<Pt> {
        val.trim().parse::<f64>().ok().map(|t| Pt(t / 20.0))
    }
}

/// Paragraph line spacing (`w:spacing/@w:line` + `@w:lineRule`).
///
/// The three named multiples and [`Multiple`](LineSpacing::Multiple) are *auto* multiples
/// of a line, expressed in 240ths of a line (`w:lineRule="auto"`): [`Single`] is `240`,
/// [`OnePointFive`] is `360`, [`Double`] is `480`, and `Multiple(x)` is `round(240 * x)`.
/// [`Exactly`](LineSpacing::Exactly) and [`AtLeast`](LineSpacing::AtLeast) pin an absolute
/// line height in points, stored in twentieths of a point with `w:lineRule="exact"` /
/// `"atLeast"` respectively — matching python-docx's `WD_LINE_SPACING` semantics.
///
/// Reading is tolerant: an `"auto"` rule (or a missing rule alongside a `w:line` value) is
/// read as the corresponding multiple, with the three round values recovered as their named
/// variants (`480` → [`Double`], not `Multiple(2.0)`).
///
/// [`Single`]: LineSpacing::Single
/// [`OnePointFive`]: LineSpacing::OnePointFive
/// [`Double`]: LineSpacing::Double
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineSpacing {
    /// Single spacing (`w:line="240" w:lineRule="auto"`).
    Single,
    /// One-and-a-half spacing (`w:line="360" w:lineRule="auto"`).
    OnePointFive,
    /// Double spacing (`w:line="480" w:lineRule="auto"`).
    Double,
    /// An arbitrary multiple of single spacing (`w:line="round(240 * x)" w:lineRule="auto"`).
    Multiple(f64),
    /// An exact line height in points (`w:lineRule="exact"`).
    Exactly(Pt),
    /// A minimum line height in points (`w:lineRule="atLeast"`).
    AtLeast(Pt),
}

impl LineSpacing {
    /// This spacing as the `(w:line value, w:lineRule value)` pair to write.
    pub(crate) fn to_line_and_rule(self) -> (String, &'static str) {
        match self {
            LineSpacing::Single => ("240".to_string(), "auto"),
            LineSpacing::OnePointFive => ("360".to_string(), "auto"),
            LineSpacing::Double => ("480".to_string(), "auto"),
            LineSpacing::Multiple(x) => (((x * 240.0).round() as i64).to_string(), "auto"),
            LineSpacing::Exactly(pt) => (pt.to_twentieths_string(), "exact"),
            LineSpacing::AtLeast(pt) => (pt.to_twentieths_string(), "atLeast"),
        }
    }

    /// Parse a `(w:line, w:lineRule)` pair. `line` must be an integer; the rule selects the
    /// interpretation (`"exact"`/`"atLeast"` are absolute point heights, `"auto"` or a
    /// missing rule is an auto multiple). Returns `None` when `line` is not an integer or
    /// the rule is present but unrecognized.
    pub(crate) fn from_line_and_rule(line: &str, rule: Option<&str>) -> Option<LineSpacing> {
        let n: i64 = line.trim().parse().ok()?;
        match rule.map(str::trim) {
            Some("exact") => Some(LineSpacing::Exactly(Pt(n as f64 / 20.0))),
            Some("atLeast") => Some(LineSpacing::AtLeast(Pt(n as f64 / 20.0))),
            Some("auto") | None => Some(match n {
                240 => LineSpacing::Single,
                360 => LineSpacing::OnePointFive,
                480 => LineSpacing::Double,
                _ => LineSpacing::Multiple(n as f64 / 240.0),
            }),
            _ => None,
        }
    }
}

/// A tab-stop alignment (`w:tab/@w:val`).
///
/// Maps to `left` / `center` / `right` / `decimal`. Reading also accepts the strict aliases
/// `start` (→ [`Left`](TabAlignment::Left)) and `end` (→ [`Right`](TabAlignment::Right));
/// other values (`bar`, `num`, `clear`, …) are not modeled and read as `None`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabAlignment {
    /// Left-aligned tab (`w:val="left"`).
    Left,
    /// Centered tab (`w:val="center"`).
    Center,
    /// Right-aligned tab (`w:val="right"`).
    Right,
    /// Decimal-aligned tab (`w:val="decimal"`).
    Decimal,
}

impl TabAlignment {
    /// This alignment as a `w:tab` `w:val` string.
    pub(crate) fn to_val(self) -> &'static str {
        match self {
            TabAlignment::Left => "left",
            TabAlignment::Center => "center",
            TabAlignment::Right => "right",
            TabAlignment::Decimal => "decimal",
        }
    }

    /// Parse a `w:tab` `w:val`. Recognizes `left`/`center`/`right`/`decimal` plus the
    /// strict aliases `start`→`Left` and `end`→`Right`; anything else is `None`.
    pub(crate) fn from_val(val: &str) -> Option<TabAlignment> {
        match val.trim() {
            "left" | "start" => Some(TabAlignment::Left),
            "center" => Some(TabAlignment::Center),
            "right" | "end" => Some(TabAlignment::Right),
            "decimal" => Some(TabAlignment::Decimal),
            _ => None,
        }
    }
}

/// A tab-stop leader — the character drawn to fill the space a tab spans (`w:tab/@w:leader`).
///
/// [`None`](TabLeader::None) means no leader; the attribute is omitted entirely when writing
/// it and a `w:leader="none"` (or a missing attribute) reads back as `None`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabLeader {
    /// No leader (`w:leader` omitted, or `"none"`).
    None,
    /// A dotted leader (`w:leader="dot"`) — the common "table of contents" dots.
    Dots,
    /// A dashed leader (`w:leader="hyphen"`).
    Dashes,
    /// A solid underline leader (`w:leader="underscore"`).
    Underscore,
}

impl TabLeader {
    /// This leader as a `w:tab` `w:leader` string, or `None` for [`TabLeader::None`] (whose
    /// caller omits the attribute rather than writing `"none"`).
    pub(crate) fn to_val(self) -> Option<&'static str> {
        match self {
            TabLeader::None => Option::None,
            TabLeader::Dots => Some("dot"),
            TabLeader::Dashes => Some("hyphen"),
            TabLeader::Underscore => Some("underscore"),
        }
    }

    /// Parse a `w:tab` `w:leader`. `dot`→`Dots`, `hyphen`→`Dashes`, `underscore`→
    /// `Underscore`, and `none` (or anything unrecognized) → [`TabLeader::None`].
    pub(crate) fn from_val(val: &str) -> TabLeader {
        match val.trim() {
            "dot" => TabLeader::Dots,
            "hyphen" => TabLeader::Dashes,
            "underscore" => TabLeader::Underscore,
            _ => TabLeader::None,
        }
    }
}

/// A run-level break (`w:br`) kind.
///
/// [`Page`](BreakType::Page) and [`Column`](BreakType::Column) write `w:br w:type="page"` /
/// `"column"`; [`Line`](BreakType::Line) is a bare `w:br` (the default type). All three read
/// back as a newline in [`Run::text`](crate::Run::text) / [`Paragraph::text`](crate::Paragraph::text).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BreakType {
    /// A page break (`w:br w:type="page"`).
    Page,
    /// A column break (`w:br w:type="column"`).
    Column,
    /// A line break (a bare `w:br`, i.e. the default `textWrapping` type).
    Line,
}

impl BreakType {
    /// The `w:type` attribute value to write, or `None` for a bare `w:br`
    /// ([`Line`](BreakType::Line), the default type).
    pub(crate) fn type_val(self) -> Option<&'static str> {
        match self {
            BreakType::Page => Some("page"),
            BreakType::Column => Some("column"),
            BreakType::Line => None,
        }
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
