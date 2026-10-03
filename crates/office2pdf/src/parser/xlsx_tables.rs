//! The paint a built-in table style lays over an Excel table (`xl/tables/*.xml`).
//!
//! A table's styling lives only in `<tableStyleInfo>`. None of it is recoverable
//! from the cell records: every body cell of the audited table carries a
//! `cellXfs` entry with `fillId="0"` and no border, so reading styles alone
//! leaves all 171 rows white (issue #532) and every rule missing (issue #1080).

use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::Event;

use super::xlsx_cells::parse_column_letters;
use super::xlsx_drawing::{parse_rels_by_type, read_zip_entry_string, resolve_relative_xl_path};
use crate::ir::{BorderLineStyle, BorderSide, CellBorder, Color, LineCap, LineJoin};
use crate::parser::xml_util::get_attr_str;

/// One table's range together with the paint its built-in style lays over it.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct TableStyleRange {
    start_col: u32,
    end_col: u32,
    /// The table's first row, which is its first header row when it has one.
    start_row: u32,
    /// First body row — the table's first row plus its header rows.
    first_body_row: u32,
    end_row: u32,
    totals_row_count: u32,
    /// The fill every body row takes, where the style paints one.
    body: Option<Color>,
    stripe: Option<Color>,
    total_row: Option<Color>,
    rule: Option<TableRule>,
    header: Option<HeaderPaint>,
}

impl TableStyleRange {
    /// Last column the table style can visibly paint.
    ///
    /// The table range may contain value-less cells whose band fill or rules
    /// still print. Printed-range inference must therefore retain its full
    /// width whenever the resolved style contributes any paint at all.
    pub(super) fn painted_end_col(&self) -> Option<u32> {
        (self.body.is_some()
            || self.stripe.is_some()
            || self.total_row.is_some()
            || self.rule.is_some()
            || self.header.is_some())
        .then_some(self.end_col)
    }

    /// The fill at `(col, row)`, or `None` where the table paints none.
    ///
    /// A header row takes the style's own fill. Declared totals rows take
    /// their separately measured fill; the remaining body rows shade the
    /// first row and every second one after it, alternating from the top of
    /// the body rather than from row 1. Later Medium bands also fill every
    /// body row in two tints (issue #1189).
    pub(super) fn fill_at(&self, col: u32, row: u32) -> Option<Color> {
        if !(self.start_col..=self.end_col).contains(&col)
            || !(self.start_row..=self.end_row).contains(&row)
        {
            // A header row sits above `first_body_row`, so this bound has to
            // be checked before the offset below is computed.
            return None;
        }
        if row < self.first_body_row {
            return self.header.map(|header| header.fill);
        }
        if self.totals_row_count > 0 && row > self.end_row.saturating_sub(self.totals_row_count) {
            return self.total_row;
        }
        if (row - self.first_body_row).is_multiple_of(2) {
            self.stripe.or(self.body)
        } else {
            self.body
        }
    }

    /// The style's own rules at `(col, row)`.
    ///
    /// Which boundaries carry one, and at what width, is the band's to say —
    /// see [`RuleExtent`]. A 1pt rule is the paint issue #619 measured every
    /// thin Excel border to be; the heavier ones are the 2pt and 3pt bands the
    /// later Medium bands draw (issue #1189).
    pub(super) fn border_at(&self, col: u32, row: u32) -> Option<CellBorder> {
        let rule: TableRule = self.rule?;
        if !(self.start_col..=self.end_col).contains(&col)
            || !(self.start_row..=self.end_row).contains(&row)
        {
            return None;
        }
        let side = |width: f64| BorderSide {
            width,
            color: rule.color,
            style: BorderLineStyle::Solid,
            join: LineJoin::Round,
            cap: LineCap::Flat,
        };
        let extent: RuleExtent = rule.extent;
        // The header rule closes the last header row, so a table declaring
        // `headerRowCount="0"` — where `first_body_row` is `start_row` — has
        // no row to hang it on. It is looked up before the foot so that a
        // table whose only row is its header still prints the header seam.
        let bottom_width: Option<f64> = if row + 1 == self.first_body_row {
            extent.under_header
        } else if row == self.end_row {
            extent.foot
        } else if row >= self.first_body_row {
            extent.between_body
        } else {
            // Between two header rows, which no measured style rules.
            None
        };
        // Each interior column boundary is ruled once, by the column to its
        // right, so the two neighbours never lay two bands on it.
        let left_width: Option<f64> = if col == self.start_col {
            extent.left_edge
        } else {
            extent.between_columns
        };
        let border = CellBorder {
            top: (row == self.start_row)
                .then_some(extent.top)
                .flatten()
                .map(side),
            bottom: bottom_width.map(side),
            left: left_width.map(side),
            right: (col == self.end_col)
                .then_some(extent.right_edge)
                .flatten()
                .map(side),
        };
        let CellBorder {
            top,
            bottom,
            left,
            right,
        } = &border;
        (top.is_some() || bottom.is_some() || left.is_some() || right.is_some()).then_some(border)
    }

    /// Whether `(col, row)` sits in a header row this style prints bold.
    ///
    /// Gated on the same `rule` as [`Self::border_at`], so a style resolved
    /// from banding alone would leave its header alone entirely rather than
    /// bolding it into a header that is otherwise unpainted. Every measured
    /// band rules something, so the gate only guards a future one.
    pub(super) fn bolds_header_at(&self, col: u32, row: u32) -> bool {
        self.rule.is_some() && self.is_header_cell(col, row)
    }

    /// The ink a header row's runs take, where the style states one.
    ///
    /// Only a style that fills the header states one: the Medium family's
    /// first three bands print white on their accent band and its fourth
    /// prints black on the far lighter one it fills instead. A Light header
    /// sits on the sheet's own background and carries no fill, so the export
    /// leaves its runs black by default.
    pub(super) fn header_text_color_at(&self, col: u32, row: u32) -> Option<Color> {
        self.header
            .filter(|_| self.is_header_cell(col, row))
            .map(|header| header.text)
    }

    fn is_header_cell(&self, col: u32, row: u32) -> bool {
        (self.start_col..=self.end_col).contains(&col)
            && (self.start_row..self.first_body_row).contains(&row)
    }
}

/// The style paint for each sheet that declares one, keyed by sheet name.
pub(super) type SheetTableStyles = HashMap<String, Vec<TableStyleRange>>;

/// The theme slots a built-in table style resolves its colours against.
struct StylePalette {
    /// `accent1`..`accent6`, in order.
    accents: Vec<Color>,
    /// `lt1`, the background the accent-less styles shade their band out of.
    light: Color,
    /// `dk1`, which those same styles rule in.
    dark: Color,
}

/// What a built-in table style paints over its range.
///
/// A family whose paint is unmeasured — the Light family's later bands and the
/// Dark family — resolves to no `TableStylePaint` at all rather than to a
/// guess, which leaves its table exactly as the cell records style it.
#[derive(Debug, Clone, Copy, PartialEq)]
struct TableStylePaint {
    /// The fill every body row takes, under the banding. The later Medium
    /// bands fill their whole table and band a second tint over alternate
    /// rows; the first band fills nothing under its stripe (issue #1189).
    body: Option<Color>,
    /// The banded-row fill, painted where `showRowStripes` asks for it.
    stripe: Option<Color>,
    /// The fill on a totals row, where a measured style defines one.
    total_row: Option<Color>,
    /// The rules the style lays over the table's boundaries.
    rule: Option<TableRule>,
    /// The header row's own paint, where the style gives it one.
    header: Option<HeaderPaint>,
}

/// A style's rules: one colour, and the extent it draws them over.
#[derive(Debug, Clone, Copy, PartialEq)]
struct TableRule {
    color: Color,
    extent: RuleExtent,
}

/// How far a style's rules run: the width in points it draws at each of the
/// table's boundaries, `None` at every boundary it leaves clear.
#[derive(Debug, Clone, Copy, PartialEq)]
struct RuleExtent {
    /// Above the table's first row.
    top: Option<f64>,
    /// Under its last header row.
    under_header: Option<f64>,
    /// Between two of its body rows.
    between_body: Option<f64>,
    /// Under its last row.
    foot: Option<f64>,
    /// Down its own left edge.
    left_edge: Option<f64>,
    /// Down its own right edge.
    right_edge: Option<f64>,
    /// Between two of its columns.
    between_columns: Option<f64>,
}

/// No rule anywhere, which every measured extent below is written against.
const NO_RULES: RuleExtent = RuleExtent {
    top: None,
    under_header: None,
    between_body: None,
    foot: None,
    left_edge: None,
    right_edge: None,
    between_columns: None,
};

/// `TableStyleLight1`..`7`: one 1pt rule above the header row, one under it
/// and one at the table's foot — nothing between the body rows, and no
/// verticals (issue #1080).
const LIGHT_BAND_ONE_RULES: RuleExtent = RuleExtent {
    top: Some(1.0),
    under_header: Some(1.0),
    foot: Some(1.0),
    ..NO_RULES
};

/// `TableStyleMedium1`..`7`: a 1pt rule at every row boundary and one down
/// each outer edge — five horizontals over a four-row table and two verticals
/// spanning its full height (issue #1125).
const MEDIUM_BAND_ONE_RULES: RuleExtent = RuleExtent {
    top: Some(1.0),
    under_header: Some(1.0),
    between_body: Some(1.0),
    foot: Some(1.0),
    left_edge: Some(1.0),
    right_edge: Some(1.0),
    ..NO_RULES
};

/// `TableStyleMedium8`..`14`: a 3pt seam under the header, a 1pt seam between
/// the body rows and one down each boundary *between* the columns. The table's
/// own top, foot and outer edges stay clear, so the band reads as filled rows
/// held apart by the sheet's background (issue #1189).
const MEDIUM_BAND_TWO_RULES: RuleExtent = RuleExtent {
    under_header: Some(3.0),
    between_body: Some(1.0),
    between_columns: Some(1.0),
    ..NO_RULES
};

/// `TableStyleMedium15`, the accent-less member of band 3: the 2pt box its
/// accented siblings draw, plus a 1pt rule between the body rows and one down
/// every column boundary (issue #1189).
const MEDIUM_BAND_THREE_NEUTRAL_RULES: RuleExtent = RuleExtent {
    top: Some(2.0),
    under_header: Some(2.0),
    between_body: Some(1.0),
    foot: Some(2.0),
    left_edge: Some(1.0),
    right_edge: Some(1.0),
    between_columns: Some(1.0),
};

/// `TableStyleMedium16`..`21`: a 2pt rule above the header, under it and at
/// the table's foot, and nothing else — no verticals and no interior rules
/// (issue #1189).
const MEDIUM_BAND_THREE_ACCENT_RULES: RuleExtent = RuleExtent {
    top: Some(2.0),
    under_header: Some(2.0),
    foot: Some(2.0),
    ..NO_RULES
};

/// `TableStyleMedium22`..`28`: a 1pt rule at every row boundary and at every
/// column boundary, interior verticals included — band 1's extent with the
/// interiors ruled too, and the same for every member of the band (issue
/// #1189).
const MEDIUM_BAND_FOUR_RULES: RuleExtent = RuleExtent {
    between_columns: Some(1.0),
    ..MEDIUM_BAND_ONE_RULES
};

/// The paint a style lays over its header row.
///
/// Bands 1..3 of the Medium family fill the row in the style's own colour and
/// set the runs on it white; band 4 fills it at the 80% tint its body rows
/// band off and prints them black. The Light band fills nothing, so it carries
/// no `HeaderPaint` at all and its header keeps the sheet's own ink.
#[derive(Debug, Clone, Copy, PartialEq)]
struct HeaderPaint {
    fill: Color,
    text: Color,
}

/// The tint a banded row takes off its style's accent.
const BAND_TINT: f64 = 0.8;
/// The darker of the two tints a band that fills every row alternates between
/// — `#B8CCE4` against a `4F81BD` accent 1 (issue #1189).
const DARK_BAND_TINT: f64 = 0.6;
/// The tint the Medium band's rules take off the same accent.
const RULE_TINT: f64 = 0.4;
/// The shade an accent-less style bands with, off `lt1`. Negative, because a
/// neutral style darkens its background where an accented one lightens.
const NEUTRAL_BAND_TINT: f64 = -0.15;
/// The deeper shade the accent-less member of a band that fills every row
/// takes — `#A6A6A6` off a white `lt1`, which is where its accented siblings
/// take [`DARK_BAND_TINT`] (issue #1189).
const NEUTRAL_DARK_BAND_TINT: f64 = -0.35;

/// The colours one member of a style band resolves its paint from.
///
/// Each band runs one accent-less member followed by accent 1 through 6. An
/// accented member walks a single ramp off its accent; the accent-less one
/// walks the same positions off `lt1` and `dk1` instead, giving `#D9D9D9` and
/// `#A6A6A6` where an accent gives `#DCE6F1` and `#B8CCE4`.
#[derive(Debug, Clone, Copy, PartialEq)]
struct MemberColors {
    /// The style's own colour, which bands 1..3 fill their header row with.
    strong: Color,
    /// The lighter of the two body tints, and the only one a band that leaves
    /// alternate rows clear paints.
    light_band: Color,
    /// The darker body tint, banded over the light one by the bands that fill
    /// every row.
    dark_band: Color,
    /// The colour a band rules in when it rules in its own.
    rule: Color,
}

impl MemberColors {
    fn resolve(accent: Option<Color>, palette: &StylePalette) -> Self {
        match accent {
            Some(accent) => Self {
                strong: accent,
                light_band: tint(accent, BAND_TINT),
                dark_band: tint(accent, DARK_BAND_TINT),
                rule: tint(accent, RULE_TINT),
            },
            None => Self {
                strong: palette.dark,
                light_band: tint(palette.light, NEUTRAL_BAND_TINT),
                dark_band: tint(palette.light, NEUTRAL_DARK_BAND_TINT),
                rule: palette.dark,
            },
        }
    }
}

/// Excel's built-in table styles derive their paint from the theme palette.
///
/// Each family runs in bands of seven: one accent-less style followed by
/// accent 1 through 6. Every band below is measured from native Excel-for-Mac
/// exports of one table restyled through each member of it — the later Medium
/// bands off `tests/fixtures/xlsx/table-multiple.xlsx`, whose three columns
/// and four body rows tell an interior vertical apart from an outer edge.
///
/// `TableStyleLight1`..`7` band the body, rule above the header row, under it
/// and at the table's foot, and print the header bold. The accent-less member
/// bands out of `lt1` shaded 15% and rules in `dk1` — `#D9D9D9` over black —
/// while `Light2` bands in `#DCE6F1` and rules in the `4F81BD` accent itself
/// (issue #1080).
///
/// `TableStyleMedium1`..`7` band identically but paint far more around it: the
/// header row is filled in the accent with white bold runs on it, and the
/// rules — the accent at a 40% tint, `#95B3D7` against `4F81BD` — run at every
/// row boundary and down both outer edges. `Medium1` opens that band on the
/// dark style, filling and ruling in `dk1` (issue #1125).
///
/// The Medium family's three later bands paint something else again, and each
/// of them fills or rules more than band 1 does (issue #1189):
///
/// - `Medium8`..`14` keep the accent header and fill **every** body row, in
///   two tints — the accent at 60% over a whole-table 80% — then seam the rows
///   apart in `lt1`: 3pt under the header, 1pt between the body rows and 1pt
///   between the columns.
/// - `Medium15`..`21` keep the accent header, band `lt1` shaded 15% over
///   alternate rows whichever member they are, and box the table in 2pt `dk1`
///   rules above the header, under it and at its foot. The accent-less member
///   adds a 1pt rule between the body rows and one down every column boundary.
/// - `Medium22`..`28` fill every row in the same two tints as band 2 but hand
///   the header the light one, print its runs black, and rule 1pt at every row
///   boundary and every column boundary.
///
/// TODO(unverified): the Light family's later bands paint something else
/// again — `Light8` fills its header row solid and rules every row, `Light15`
/// boxes the whole table — and the Dark family is not measured at all. Both
/// stay unpainted rather than guessed at.
fn built_in_table_style(style_name: &str, palette: &StylePalette) -> Option<TableStylePaint> {
    let (index, is_medium): (usize, bool) = match style_name.strip_prefix("TableStyleMedium") {
        Some(digits) => (digits.parse().ok()?, true),
        None => (
            style_name.strip_prefix("TableStyleLight")?.parse().ok()?,
            false,
        ),
    };

    let band: usize = (index.checked_sub(1)?) / 7;
    let within_band: usize = (index - 1) % 7;
    let accent: Option<Color> = match within_band.checked_sub(1) {
        Some(accent_index) => Some(*palette.accents.get(accent_index)?),
        // The band opens on the accent-less style: the neutral one in the
        // Light family, the dark one in the Medium family.
        None => None,
    };
    let colors: MemberColors = MemberColors::resolve(accent, palette);

    if !is_medium {
        // Only the Light family's first band is measured.
        if band > 0 {
            return None;
        }
        return Some(TableStylePaint {
            body: None,
            stripe: Some(colors.light_band),
            total_row: None,
            // A Light table rules in its own colour, not at the 40% tint the
            // Medium band takes.
            rule: Some(TableRule {
                color: colors.strong,
                extent: LIGHT_BAND_ONE_RULES,
            }),
            header: None,
        });
    }

    let white_header = |fill: Color| {
        Some(HeaderPaint {
            fill,
            text: Color::white(),
        })
    };
    match band {
        0 => Some(TableStylePaint {
            body: None,
            stripe: Some(colors.light_band),
            total_row: None,
            rule: Some(TableRule {
                color: colors.rule,
                extent: MEDIUM_BAND_ONE_RULES,
            }),
            header: white_header(colors.strong),
        }),
        1 => Some(TableStylePaint {
            body: Some(colors.light_band),
            stripe: Some(colors.dark_band),
            // The native Excel export for issue #1970 shows Medium9's total
            // row in its accent color. Keep other total-row mappings unset
            // until their built-in styles are measured as well.
            total_row: (index == 9).then_some(colors.strong),
            // The seams are the sheet's own background showing between the
            // filled rows, which both measured members print `#FFFFFF`.
            rule: Some(TableRule {
                color: palette.light,
                extent: MEDIUM_BAND_TWO_RULES,
            }),
            header: white_header(colors.strong),
        }),
        2 => Some(TableStylePaint {
            body: None,
            // Band 3 bands in `lt1` shaded 15% under an accent header as well
            // as under a dark one: both measured members print `#D9D9D9`.
            stripe: Some(tint(palette.light, NEUTRAL_BAND_TINT)),
            total_row: None,
            rule: Some(TableRule {
                color: palette.dark,
                extent: if accent.is_some() {
                    MEDIUM_BAND_THREE_ACCENT_RULES
                } else {
                    MEDIUM_BAND_THREE_NEUTRAL_RULES
                },
            }),
            header: white_header(colors.strong),
        }),
        3 => Some(TableStylePaint {
            body: Some(colors.light_band),
            stripe: Some(colors.dark_band),
            total_row: None,
            rule: Some(TableRule {
                color: colors.rule,
                extent: MEDIUM_BAND_FOUR_RULES,
            }),
            header: Some(HeaderPaint {
                fill: colors.light_band,
                text: Color::black(),
            }),
        }),
        // Excel's Medium family ends at 28, so a higher index is a name no
        // built-in style carries.
        _ => None,
    }
}

/// Excel's integer HLS range.
///
/// The OOXML note on `tint` names 255, but only 240 — the range Win32's
/// `RGBToHLS` has always used — reproduces a native export: a `4F81BD`
/// accent 1 tinted 0.8 prints `DCE6F1`, which 255 misses by a level
/// (`DBE5F1`) and a per-channel RGB tint misses the other way (`DCE6F2`,
/// issue #1125).
const HLS_MAX: i32 = 240;
const RGB_MAX: i32 = 255;

/// Move `color`'s luminance `amount` of the way to white, or to black where
/// `amount` is negative, the way OOXML's `tint` attribute does.
///
/// Measured against native Excel-for-Mac exports of one table restyled under
/// two Office themes: `4F81BD` gives `DCE6F1`/`95B3D7`, `5B9BD5` gives
/// `DDEBF7`/`9BC2E6` and `C0504D` gives `F2DCDB`/`DA9694` at 0.8 and 0.4, and
/// white gives `D9D9D9` at -0.15. The moved luminance truncates rather than
/// rounds — rounding misses four of those seven by a level.
fn tint(color: Color, amount: f64) -> Color {
    let (hue, luminance, saturation): (i32, i32, i32) = to_hls(color);
    let moved: f64 = if amount < 0.0 {
        f64::from(luminance) * (1.0 + amount)
    } else {
        f64::from(luminance) * (1.0 - amount) + f64::from(HLS_MAX) * amount
    };
    from_hls(hue, moved as i32, saturation)
}

/// Win32's `RGBToHLS`, in the integer arithmetic it is defined with.
fn to_hls(color: Color) -> (i32, i32, i32) {
    let (red, green, blue): (i32, i32, i32) =
        (i32::from(color.r), i32::from(color.g), i32::from(color.b));
    let brightest: i32 = red.max(green).max(blue);
    let darkest: i32 = red.min(green).min(blue);
    let sum: i32 = brightest + darkest;
    let span: i32 = brightest - darkest;
    let luminance: i32 = (sum * HLS_MAX + RGB_MAX) / (2 * RGB_MAX);
    if span == 0 {
        return (0, luminance, 0);
    }
    let saturation: i32 = if luminance <= HLS_MAX / 2 {
        (span * HLS_MAX + sum / 2) / sum
    } else {
        (span * HLS_MAX + (2 * RGB_MAX - sum) / 2) / (2 * RGB_MAX - sum)
    };
    let distance =
        |channel: i32| -> i32 { ((brightest - channel) * (HLS_MAX / 6) + span / 2) / span };
    let hue: i32 = if red == brightest {
        distance(blue) - distance(green)
    } else if green == brightest {
        HLS_MAX / 3 + distance(red) - distance(blue)
    } else {
        2 * HLS_MAX / 3 + distance(green) - distance(red)
    };
    (hue.rem_euclid(HLS_MAX), luminance, saturation)
}

/// Win32's `HLSToRGB`.
///
/// Its zero-saturation shortcut is left out on purpose: it truncates where the
/// general path rounds, which lands `D8D8D8` for the `lt1` shade a native
/// export bands `D9D9D9`. With no saturation the general path collapses to the
/// same single level anyway, so dropping the shortcut only changes the
/// rounding.
fn from_hls(hue: i32, luminance: i32, saturation: i32) -> Color {
    let upper: i32 = if luminance <= HLS_MAX / 2 {
        (luminance * (HLS_MAX + saturation) + HLS_MAX / 2) / HLS_MAX
    } else {
        luminance + saturation - (luminance * saturation + HLS_MAX / 2) / HLS_MAX
    };
    let lower: i32 = 2 * luminance - upper;
    let channel = |hue_offset: i32| -> u8 {
        let level: i32 = hue_level(lower, upper, hue + hue_offset);
        ((level * RGB_MAX + HLS_MAX / 2) / HLS_MAX).clamp(0, RGB_MAX) as u8
    };
    Color::new(channel(HLS_MAX / 3), channel(0), channel(-HLS_MAX / 3))
}

/// One channel's level, interpolated between `lower` and `upper` across the
/// hue wheel — Win32's `HueToRGB`.
fn hue_level(lower: i32, upper: i32, hue: i32) -> i32 {
    let hue: i32 = hue.rem_euclid(HLS_MAX);
    let sixth: i32 = HLS_MAX / 6;
    if hue < sixth {
        lower + ((upper - lower) * hue + HLS_MAX / 12) / sixth
    } else if hue < HLS_MAX / 2 {
        upper
    } else if hue < 2 * HLS_MAX / 3 {
        lower + ((upper - lower) * (2 * HLS_MAX / 3 - hue) + HLS_MAX / 12) / sixth
    } else {
        lower
    }
}

/// Read a table part's style paint.
///
/// `None` when the range is unreadable, or when the style resolves to nothing
/// the table would print — an unmapped style family, or a mapped one whose
/// only paint is a stripe the table turned off.
fn parse_table_part(xml: &str, palette: &StylePalette) -> Option<TableStyleRange> {
    let mut reader = Reader::from_str(xml);
    let mut range: Option<(u32, u32, u32, u32)> = None;
    let mut header_rows: u32 = 0;
    let mut totals_row_count: u32 = 0;
    let mut paint: Option<TableStylePaint> = None;
    let mut shows_row_stripes: bool = false;

    loop {
        let event = reader.read_event();
        let element = match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => e.clone(),
            Ok(Event::Eof) | Err(_) => break,
            _ => continue,
        };
        match element.local_name().as_ref() {
            b"table" => {
                range = get_attr_str(&element, b"ref")
                    .as_deref()
                    .and_then(parse_range);
                header_rows = get_attr_str(&element, b"headerRowCount")
                    .and_then(|value| value.parse::<u32>().ok())
                    // ECMA-376 defaults headerRowCount to 1.
                    .unwrap_or(1);
                totals_row_count = get_attr_str(&element, b"totalsRowCount")
                    .and_then(|value| value.parse::<u32>().ok())
                    .unwrap_or(0);
            }
            b"tableStyleInfo" => {
                // `showRowStripes` governs the banding alone: a `Light1` table
                // that turns it off still prints its three rules and its bold
                // header, as the native export of that variant shows.
                shows_row_stripes =
                    get_attr_str(&element, b"showRowStripes").as_deref() == Some("1");
                paint = get_attr_str(&element, b"name")
                    .as_deref()
                    .and_then(|name| built_in_table_style(name, palette));
            }
            _ => {}
        }
    }

    let (start_col, start_row, end_col, end_row) = range?;
    let paint: TableStylePaint = paint?;
    let first_body_row: u32 = start_row + header_rows;
    let table_data_rows: u32 = end_row
        .checked_sub(first_body_row)
        .map_or(0, |rows_before_last| rows_before_last.saturating_add(1));
    totals_row_count = totals_row_count.min(table_data_rows);
    let body_end_row: u32 = end_row.saturating_sub(totals_row_count);
    let has_body_rows: bool = first_body_row <= body_end_row;
    // `showRowStripes` scopes the banded tint alone: a band-2 table with it
    // off keeps the whole-table fill under it, as it keeps its rules.
    let stripe: Option<Color> = paint.stripe.filter(|_| shows_row_stripes && has_body_rows);
    let body: Option<Color> = paint.body.filter(|_| has_body_rows);
    let total_row: Option<Color> = paint.total_row.filter(|_| totals_row_count > 0);
    (stripe.is_some()
        || body.is_some()
        || total_row.is_some()
        || paint.rule.is_some()
        || paint.header.is_some())
    .then_some(TableStyleRange {
        start_col,
        end_col,
        start_row,
        first_body_row,
        end_row,
        totals_row_count,
        body,
        stripe,
        total_row,
        rule: paint.rule,
        header: paint.header,
    })
}

/// Parse an `A4:H175`-style range into `(start_col, start_row, end_col, end_row)`.
fn parse_range(reference: &str) -> Option<(u32, u32, u32, u32)> {
    let (start, end) = reference.split_once(':')?;
    let (start_col, start_row) = parse_cell_reference(start)?;
    let (end_col, end_row) = parse_cell_reference(end)?;
    Some((start_col, start_row, end_col, end_row))
}

/// Parse an `A4`-style reference into a 1-based `(column, row)`.
fn parse_cell_reference(reference: &str) -> Option<(u32, u32)> {
    let split: usize = reference.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = reference.split_at(split);
    Some((parse_column_letters(letters)?, digits.parse().ok()?))
}

/// The theme slots the built-in styles resolve against.
///
/// A workbook whose theme part is missing or unreadable still prints an
/// accent-less style, because Excel's default `lt1`/`dk1` are plain white and
/// black; only the accented members go unresolved.
fn style_palette(theme_xml: &str) -> StylePalette {
    let scheme = crate::parser::drawingml::parse_theme_color_scheme(theme_xml);
    StylePalette {
        accents: (1..=6)
            .map_while(|index| scheme.get(&format!("accent{index}")).copied())
            .collect(),
        light: scheme.get("lt1").copied().unwrap_or(Color::white()),
        dark: scheme.get("dk1").copied().unwrap_or(Color::black()),
    }
}

/// Collect every sheet's table-style paint from the workbook.
pub(super) fn extract_table_styles(data: &[u8]) -> SheetTableStyles {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return SheetTableStyles::new();
    };
    let workbook_xml: String = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let workbook_rels_xml: String =
        read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = super::xlsx_drawing::parse_rels_targets(&workbook_rels_xml);

    let theme_path: String = parse_rels_by_type(&workbook_rels_xml, "theme")
        .first()
        .map(|target| resolve_relative_xl_path("xl", target))
        .unwrap_or_else(|| "xl/theme/theme1.xml".to_string());
    let theme_xml: String = read_zip_entry_string(&mut archive, &theme_path);
    let palette: StylePalette = style_palette(&theme_xml);

    let mut styles: SheetTableStyles = HashMap::new();
    for (sheet_name, sheet_rid) in super::xlsx_drawing::parse_workbook_sheet_rids(&workbook_xml) {
        let Some(sheet_target) = rid_to_target.get(&sheet_rid) else {
            continue;
        };
        let sheet_path: String = format!("xl/{sheet_target}");
        let filename: &str = sheet_path.rsplit('/').next().unwrap_or(&sheet_path);
        let rels_path: String = format!("xl/worksheets/_rels/{filename}.rels");
        let rels_xml: String = read_zip_entry_string(&mut archive, &rels_path);
        if rels_xml.is_empty() {
            continue;
        }
        let mut sheet_styles: Vec<TableStyleRange> = Vec::new();
        for target in parse_rels_by_type(&rels_xml, "table") {
            let table_path: String = resolve_relative_xl_path("xl/worksheets", &target);
            let table_xml: String = read_zip_entry_string(&mut archive, &table_path);
            if table_xml.is_empty() {
                continue;
            }
            if let Some(parsed) = parse_table_part(&table_xml, &palette) {
                sheet_styles.push(parsed);
            }
        }
        if !sheet_styles.is_empty() {
            styles.insert(sheet_name, sheet_styles);
        }
    }
    styles
}

#[cfg(test)]
#[path = "xlsx_tables_tests.rs"]
mod tests;
