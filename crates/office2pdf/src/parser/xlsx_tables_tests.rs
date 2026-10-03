use super::*;

/// Office's default palette, as the audited workbooks' themes declare it —
/// six accents over a white `lt1` and a black `dk1`.
fn palette() -> StylePalette {
    StylePalette {
        accents: vec![
            Color::new(0x4f, 0x81, 0xbd), // accent1
            Color::new(0xc0, 0x50, 0x4d), // accent2
            Color::new(0x9b, 0xbb, 0x59), // accent3
            Color::new(0x80, 0x64, 0xa2), // accent4
            Color::new(0x4b, 0xac, 0xc6), // accent5
            Color::new(0xf7, 0x96, 0x46), // accent6
        ],
        light: Color::white(),
        dark: Color::black(),
    }
}

/// The audited workbook's table part, trimmed to the elements that matter.
fn module_inventory_table(style_info: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1"
               name="ModuleInventory" displayName="ModuleInventory" ref="A4:H175"
               headerRowCount="1" totalsRowCount="0">
          <autoFilter ref="A4:H175"/>
          {style_info}
        </table>"#
    )
}

/// The one-factor probe's table part: the `G1:I4` range of
/// `tests/fixtures/xlsx/ExcelTables.xlsx`, whose `headerRowCount` defaults to
/// 1, restyled through each built-in style in turn (issue #1189).
fn probe_table(style_name: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1"
               name="Probe" displayName="Probe" ref="G1:I4" totalsRowCount="0">
          <tableStyleInfo name="{style_name}" showRowStripes="1" showColumnStripes="0"/>
        </table>"#
    )
}

/// The band a style lays over the probe table's three body rows, top to
/// bottom, and the fill it lays over its header row.
fn probe_fills(style_name: &str) -> (Option<Color>, Vec<Option<Color>>) {
    let style = parse_table_part(&probe_table(style_name), &palette())
        .unwrap_or_else(|| panic!("{style_name} paints its table"));
    (
        style.fill_at(7, 1),
        (2..=4).map(|row| style.fill_at(7, row)).collect(),
    )
}

/// The colours a native Excel-for-Mac export prints for a theme colour tinted
/// by each amount the built-in table styles use.
///
/// Three accents from two Office themes plus the neutral shade, measured off
/// `mutool draw -F trace` of the same table restyled through Excel 16
/// (issue #1125). Together they pin the whole tint model: nothing but the
/// 240-step HLS range with a truncated luminance reproduces all seven.
fn measured_tints() -> [(Color, f64, Color); 7] {
    [
        // accent 1 of the Office 2007 theme, banded and ruled
        (
            Color::new(0x4f, 0x81, 0xbd),
            0.8,
            Color::new(0xdc, 0xe6, 0xf1),
        ),
        (
            Color::new(0x4f, 0x81, 0xbd),
            0.4,
            Color::new(0x95, 0xb3, 0xd7),
        ),
        // accent 2 of the same theme, from the TableStyleMedium3 variant
        (
            Color::new(0xc0, 0x50, 0x4d),
            0.8,
            Color::new(0xf2, 0xdc, 0xdb),
        ),
        (
            Color::new(0xc0, 0x50, 0x4d),
            0.4,
            Color::new(0xda, 0x96, 0x94),
        ),
        // accent 1 of the current Office theme, from `SH001-Table.xlsx`
        (
            Color::new(0x5b, 0x9b, 0xd5),
            0.8,
            Color::new(0xdd, 0xeb, 0xf7),
        ),
        (
            Color::new(0x5b, 0x9b, 0xd5),
            0.4,
            Color::new(0x9b, 0xc2, 0xe6),
        ),
        // `lt1` shaded, which is the band an accent-less style prints
        (Color::white(), -0.15, Color::new(0xd9, 0xd9, 0xd9)),
    ]
}

#[test]
fn tinting_reproduces_every_measured_native_export_colour() {
    for (source, amount, expected) in measured_tints() {
        assert_eq!(
            tint(source, amount),
            expected,
            "{source:?} tinted {amount} should print {expected:?}"
        );
    }
}

#[test]
fn a_medium2_table_stripes_in_accent1_at_a_twenty_percent_tint() {
    // Excel paints DCE6F1 for TableStyleMedium2 against a 4F81BD accent 1,
    // which is the accent moved 80% of the way to white on Excel's HLS range.
    let xml = module_inventory_table(
        r#"<tableStyleInfo name="TableStyleMedium2" showRowStripes="1" showColumnStripes="0"/>"#,
    );

    let stripes = parse_table_part(&xml, &palette()).expect("stripes are declared");

    assert_eq!(stripes.fill_at(1, 5), Some(Color::new(0xdc, 0xe6, 0xf1)));
}

#[test]
fn stripes_start_at_the_first_body_row_and_alternate() {
    // The header occupies row 4, so row 5 is the first body row — and Excel
    // shades it, then every second row after it.
    let xml =
        module_inventory_table(r#"<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>"#);

    let stripes = parse_table_part(&xml, &palette()).unwrap();

    assert_eq!(
        stripes.fill_at(1, 4),
        Some(palette().accents[0]),
        "the header row takes the style's own fill, not the band"
    );
    assert!(stripes.fill_at(1, 5).is_some());
    assert!(stripes.fill_at(1, 6).is_none());
    assert!(stripes.fill_at(1, 7).is_some());
    assert!(
        stripes.fill_at(1, 175).is_some(),
        "the last body row is 175"
    );
    assert!(stripes.fill_at(1, 176).is_none(), "past the table range");
}

#[test]
fn medium9_totals_row_uses_its_own_fill() {
    let xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                   ref="G1:I4" headerRowCount="1" totalsRowCount="1">
      <tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/>
    </table>"#;

    let style = parse_table_part(xml, &palette()).unwrap();

    assert_eq!(style.fill_at(7, 3), Some(accent_light_band()));
    assert_eq!(
        style.fill_at(7, 4),
        Some(accent()),
        "the declared totals row uses its measured total-row fill"
    );
}

#[test]
fn stripes_stay_inside_the_tables_columns() {
    let xml =
        module_inventory_table(r#"<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>"#);

    let stripes = parse_table_part(&xml, &palette()).unwrap();

    assert!(stripes.fill_at(8, 5).is_some(), "column H is the last one");
    assert!(stripes.fill_at(9, 5).is_none(), "column I is outside");
}

#[test]
fn a_table_that_asks_for_no_row_stripes_keeps_the_rest_of_its_style() {
    // `showRowStripes` scopes the banding alone, so a Medium2 table with it
    // off still fills and rules the header row the native export fills.
    let xml = module_inventory_table(
        r#"<tableStyleInfo name="TableStyleMedium2" showRowStripes="0" showColumnStripes="1"/>"#,
    );

    let style = parse_table_part(&xml, &palette()).expect("the header paint survives");

    assert_eq!(style.fill_at(1, 5), None, "no band without row stripes");
    assert_eq!(style.fill_at(1, 4), Some(palette().accents[0]));
    assert!(style.border_at(1, 4).is_some());
}

/// The colours the probe's accent 1 resolves to, at each tint the built-in
/// table styles use, plus the two `lt1` shades their accent-less members take.
fn accent() -> Color {
    Color::new(0x4f, 0x81, 0xbd)
}
fn accent_light_band() -> Color {
    Color::new(0xdc, 0xe6, 0xf1)
}
fn accent_dark_band() -> Color {
    Color::new(0xb8, 0xcc, 0xe4)
}
fn accent_rule() -> Color {
    Color::new(0x95, 0xb3, 0xd7)
}
fn neutral_light_band() -> Color {
    Color::new(0xd9, 0xd9, 0xd9)
}
fn neutral_dark_band() -> Color {
    Color::new(0xa6, 0xa6, 0xa6)
}

/// One side's colour and width, for comparing a rule against its measurement.
fn ruled(side: &Option<BorderSide>) -> Option<(Color, f64)> {
    side.as_ref().map(|side| (side.color, side.width))
}

#[test]
fn medium_band_two_fills_every_body_row_in_two_tints() {
    // `Medium9` prints `#B8CCE4`, `#DCE6F1`, `#B8CCE4` over the probe's three
    // body rows under a solid accent header — the accent at 60% over a
    // whole-table 80%, not one 80% stripe with white between (issue #1189).
    assert_eq!(
        probe_fills("TableStyleMedium9"),
        (
            Some(accent()),
            vec![
                Some(accent_dark_band()),
                Some(accent_light_band()),
                Some(accent_dark_band())
            ]
        )
    );

    // `Medium8` opens the band on the dark style: a `dk1` header over the two
    // `lt1` shades, 35% and 15%.
    assert_eq!(
        probe_fills("TableStyleMedium8"),
        (
            Some(Color::black()),
            vec![
                Some(neutral_dark_band()),
                Some(neutral_light_band()),
                Some(neutral_dark_band())
            ]
        )
    );
}

#[test]
fn medium_band_two_seams_its_rows_in_white() {
    // A 3pt white rule under the header, 1pt between the body rows and 1pt
    // down each boundary between the columns — nothing above the header, at
    // its foot, or down either outer edge (issue #1189).
    let style = parse_table_part(&probe_table("TableStyleMedium9"), &palette()).unwrap();
    let white = Color::white();

    let header = style.border_at(7, 1).expect("the header is seamed");
    assert_eq!(ruled(&header.bottom), Some((white, 3.0)));
    assert!(header.top.is_none(), "band 2 leaves the table's top clear");
    assert!(header.left.is_none(), "band 2 rules no left edge");

    let body = style.border_at(7, 2).expect("a body row is seamed");
    assert_eq!(ruled(&body.bottom), Some((white, 1.0)));

    assert_eq!(
        style.border_at(8, 2).and_then(|border| ruled(&border.left)),
        Some((white, 1.0)),
        "the boundary between columns G and H is seamed"
    );
    assert!(
        style
            .border_at(9, 2)
            .and_then(|border| border.right)
            .is_none(),
        "band 2 leaves the table's right edge clear"
    );

    assert!(
        style.border_at(7, 4).is_none(),
        "band 2 leaves the table's foot and left edge clear"
    );
}

#[test]
fn medium_band_three_bands_in_grey_inside_a_two_point_box() {
    // `Medium15` and `Medium16` both band `#D9D9D9` over the 1st and 3rd body
    // rows and rule in `#000000`, 2pt above the header, under it and at the
    // foot (issue #1189).
    assert_eq!(
        probe_fills("TableStyleMedium16"),
        (
            Some(accent()),
            vec![Some(neutral_light_band()), None, Some(neutral_light_band())]
        ),
        "band 3 bands in grey even under an accent header"
    );
    assert_eq!(
        probe_fills("TableStyleMedium15"),
        (
            Some(Color::black()),
            vec![Some(neutral_light_band()), None, Some(neutral_light_band())]
        )
    );

    let accent_member = parse_table_part(&probe_table("TableStyleMedium16"), &palette()).unwrap();
    let black = Color::black();
    let header = accent_member.border_at(7, 1).expect("the header is boxed");
    assert_eq!(ruled(&header.top), Some((black, 2.0)));
    assert_eq!(ruled(&header.bottom), Some((black, 2.0)));
    assert!(
        header.left.is_none(),
        "the accent member draws no verticals"
    );
    assert!(
        accent_member.border_at(7, 2).is_none(),
        "nor any rule between its body rows"
    );
    assert_eq!(
        accent_member
            .border_at(7, 4)
            .and_then(|border| ruled(&border.bottom)),
        Some((black, 2.0)),
        "the foot closes the box"
    );

    // The accent-less member adds a 1pt rule between the body rows and one
    // down each outer edge.
    let neutral_member = parse_table_part(&probe_table("TableStyleMedium15"), &palette()).unwrap();
    let neutral_header = neutral_member.border_at(7, 1).expect("the header is boxed");
    assert_eq!(ruled(&neutral_header.left), Some((black, 1.0)));
    let body = neutral_member.border_at(7, 2).expect("a body row is ruled");
    assert_eq!(ruled(&body.bottom), Some((black, 1.0)));
    assert_eq!(
        neutral_member
            .border_at(8, 2)
            .and_then(|border| ruled(&border.left)),
        Some((black, 1.0)),
        "the accent-less member rules every column boundary"
    );
    assert_eq!(
        neutral_member
            .border_at(9, 2)
            .and_then(|border| ruled(&border.right)),
        Some((black, 1.0))
    );
}

#[test]
fn medium_band_four_fills_every_row_and_rules_every_boundary() {
    // `Medium23` fills its header in the accent at 80% — the same fill its
    // body rows alternate off — and rules `#95B3D7` at every row boundary and
    // every column boundary (issue #1189).
    assert_eq!(
        probe_fills("TableStyleMedium23"),
        (
            Some(accent_light_band()),
            vec![
                Some(accent_dark_band()),
                Some(accent_light_band()),
                Some(accent_dark_band())
            ]
        )
    );

    let accent_member = parse_table_part(&probe_table("TableStyleMedium23"), &palette()).unwrap();
    let header = accent_member.border_at(7, 1).expect("the header is ruled");
    assert_eq!(ruled(&header.top), Some((accent_rule(), 1.0)));
    assert_eq!(ruled(&header.bottom), Some((accent_rule(), 1.0)));
    assert_eq!(ruled(&header.left), Some((accent_rule(), 1.0)));
    assert_eq!(
        accent_member
            .border_at(7, 4)
            .and_then(|border| ruled(&border.bottom)),
        Some((accent_rule(), 1.0)),
        "the foot is ruled too"
    );
    assert_eq!(
        accent_member
            .border_at(8, 2)
            .and_then(|border| ruled(&border.left)),
        Some((accent_rule(), 1.0)),
        "band 4 rules every column boundary, interiors included"
    );

    // `Medium22` bands in the two `lt1` shades under a `#D9D9D9` header and
    // rules the same extent in `#000000`.
    assert_eq!(
        probe_fills("TableStyleMedium22"),
        (
            Some(neutral_light_band()),
            vec![
                Some(neutral_dark_band()),
                Some(neutral_light_band()),
                Some(neutral_dark_band())
            ]
        )
    );
    let neutral_member = parse_table_part(&probe_table("TableStyleMedium22"), &palette()).unwrap();
    let interior = neutral_member.border_at(8, 2).expect("column H is ruled");
    assert_eq!(
        ruled(&interior.left),
        Some((Color::black(), 1.0)),
        "the accent-less member rules in `dk1`"
    );
    assert!(
        interior.right.is_none(),
        "each interior boundary is ruled once, by the column to its right"
    );
}

#[test]
fn later_medium_bands_print_their_header_runs_in_the_measured_ink() {
    // White bold in bands 2 and 3, black bold on band 4's light header.
    for style_name in [
        "TableStyleMedium8",
        "TableStyleMedium9",
        "TableStyleMedium16",
    ] {
        let style = parse_table_part(&probe_table(style_name), &palette()).unwrap();
        assert_eq!(
            style.header_text_color_at(7, 1),
            Some(Color::white()),
            "{style_name}"
        );
        assert!(style.bolds_header_at(7, 1), "{style_name}");
    }
    for style_name in ["TableStyleMedium22", "TableStyleMedium23"] {
        let style = parse_table_part(&probe_table(style_name), &palette()).unwrap();
        assert_eq!(
            style.header_text_color_at(7, 1),
            Some(Color::black()),
            "{style_name}"
        );
        assert!(style.bolds_header_at(7, 1), "{style_name}");
    }
}

#[test]
fn a_style_past_the_last_built_in_medium_index_is_unresolved() {
    // Excel's Medium family ends at 28, so anything past it is a name we have
    // no measurement for.
    for style in ["TableStyleMedium29", "TableStyleMedium35"] {
        assert_eq!(built_in_table_style(style, &palette()), None, "{style}");
    }
}

#[test]
fn medium_styles_walk_the_accents_in_bands_of_seven() {
    // Every band opens on its accent-less member and runs accent 1 through 6:
    // Medium2..7, then Medium9..14, Medium16..21 and Medium23..28. Bands 1..3
    // fill the header in the accent itself; band 4 fills it at the 80% tint
    // its body rows band off (issue #1189).
    for (style, accent_index) in [
        ("TableStyleMedium2", 0),
        ("TableStyleMedium7", 5),
        ("TableStyleMedium9", 0),
        ("TableStyleMedium14", 5),
        ("TableStyleMedium16", 0),
        ("TableStyleMedium21", 5),
    ] {
        assert_eq!(
            built_in_table_style(style, &palette()).and_then(|paint| paint.header),
            Some(HeaderPaint {
                fill: palette().accents[accent_index],
                text: Color::white(),
            }),
            "{style} should paint in accent {}",
            accent_index + 1
        );
    }
    for (style, accent_index) in [("TableStyleMedium23", 0), ("TableStyleMedium28", 5)] {
        assert_eq!(
            built_in_table_style(style, &palette()).and_then(|paint| paint.header),
            Some(HeaderPaint {
                fill: tint(palette().accents[accent_index], 0.8),
                text: Color::black(),
            }),
            "{style} should paint in accent {}",
            accent_index + 1
        );
    }

    // The accent-less members fill their header in `dk1`, except band 4's,
    // which takes the same `lt1` shaded 15% its light rows band with.
    for (dark_style, expected) in [
        ("TableStyleMedium1", Color::black()),
        ("TableStyleMedium8", Color::black()),
        ("TableStyleMedium15", Color::black()),
        ("TableStyleMedium22", neutral_light_band()),
    ] {
        assert_eq!(
            built_in_table_style(dark_style, &palette())
                .and_then(|paint| paint.header)
                .map(|header| header.fill),
            Some(expected),
            "{dark_style} opens its band on the accent-less style"
        );
    }
}

#[test]
fn medium_band_one_opens_on_the_dark_style() {
    // Medium1 is the band's accent-less member: the native export fills its
    // header row and rules it in `dk1`, and bands the body out of `lt1` the
    // same 15% Light1 does.
    assert_eq!(
        built_in_table_style("TableStyleMedium1", &palette()),
        Some(TableStylePaint {
            body: None,
            stripe: Some(Color::new(0xd9, 0xd9, 0xd9)),
            total_row: None,
            rule: Some(TableRule {
                color: Color::black(),
                extent: MEDIUM_BAND_ONE_RULES,
            }),
            header: Some(HeaderPaint {
                fill: Color::black(),
                text: Color::white(),
            }),
        })
    );
}

#[test]
fn a_medium2_table_fills_its_header_row_in_accent1_and_prints_it_white() {
    // The native export paints G1's header band solid #4F81BD and sets its
    // run in white bold — the whole treatment issue #1125 reports missing.
    let xml =
        module_inventory_table(r#"<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>"#);

    let style = parse_table_part(&xml, &palette()).unwrap();

    assert_eq!(
        style.fill_at(1, 4),
        Some(Color::new(0x4f, 0x81, 0xbd)),
        "the header row is filled in the accent itself"
    );
    assert_eq!(style.header_text_color_at(1, 4), Some(Color::white()));
    assert!(style.bolds_header_at(1, 4));

    assert_eq!(
        style.header_text_color_at(1, 5),
        None,
        "a body row keeps its own ink"
    );
    assert_eq!(
        style.header_text_color_at(9, 4),
        None,
        "column I is outside the table"
    );
}

#[test]
fn a_light1_header_row_is_bolded_but_left_in_its_own_ink() {
    // Light1's header carries no fill, so there is nothing for white text to
    // sit on and the export prints it black.
    let xml =
        module_inventory_table(r#"<tableStyleInfo name="TableStyleLight1" showRowStripes="1"/>"#);

    let style = parse_table_part(&xml, &palette()).unwrap();

    assert!(style.bolds_header_at(1, 4));
    assert_eq!(style.header_text_color_at(1, 4), None);
    assert_eq!(style.fill_at(1, 4), None, "the header row is unfilled");
}

#[test]
fn a_medium2_table_rules_every_row_boundary_and_both_outer_edges() {
    // Against Light1's three rules, the native Medium2 export lays a #95B3D7
    // 1pt band on every row boundary and runs one down the table's left and
    // right edges (issue #1125).
    let xml =
        module_inventory_table(r#"<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>"#);

    let style = parse_table_part(&xml, &palette()).unwrap();
    let rule_color = Color::new(0x95, 0xb3, 0xd7);

    let header = style.border_at(1, 4).expect("the header row is ruled");
    let top = header.top.as_ref().expect("a rule sits above the header");
    assert_eq!(top.color, rule_color, "Medium2 rules in the accent at 40%");
    assert_eq!(top.width, 1.0, "a thin Excel border is a 1pt band");
    assert_eq!(top.style, BorderLineStyle::Solid);
    assert!(header.bottom.is_some(), "the header is closed by a rule");

    let interior = style
        .border_at(4, 100)
        .expect("an interior body row is ruled too");
    assert!(
        interior.bottom.is_some(),
        "every row boundary carries a rule"
    );
    assert!(interior.top.is_none(), "the rule hangs below each row");
    assert!(
        interior.left.is_none() && interior.right.is_none(),
        "column D is neither outer edge"
    );

    let first_column = style.border_at(1, 100).expect("column A is the left edge");
    assert_eq!(
        first_column.left.as_ref().map(|side| side.color),
        Some(rule_color)
    );
    assert!(first_column.right.is_none());

    let last_column = style.border_at(8, 100).expect("column H is the right edge");
    assert_eq!(
        last_column.right.as_ref().map(|side| side.color),
        Some(rule_color)
    );
    assert!(last_column.left.is_none());

    assert!(
        style.border_at(9, 100).is_none(),
        "column I is outside the table"
    );
    assert!(style.border_at(1, 176).is_none(), "past the last row");
}

#[test]
fn light_band_one_walks_the_accents_like_medium_does() {
    // Light1 opens the band on the neutral style: `lt1` shaded 15% over `dk1`
    // rules, which is the #D9D9D9 on black the native export prints.
    assert_eq!(
        built_in_table_style("TableStyleLight1", &palette()),
        Some(TableStylePaint {
            body: None,
            stripe: Some(Color::new(0xd9, 0xd9, 0xd9)),
            total_row: None,
            rule: Some(TableRule {
                color: Color::black(),
                extent: LIGHT_BAND_ONE_RULES,
            }),
            header: None,
        })
    );

    // Light2..7 take accent 1..6, banding at the same 20% tint as Medium and
    // ruling in the accent itself.
    for (style, accent_index) in [("TableStyleLight2", 0), ("TableStyleLight7", 5)] {
        let accent = palette().accents[accent_index];

        assert_eq!(
            built_in_table_style(style, &palette()),
            Some(TableStylePaint {
                body: None,
                stripe: Some(tint(accent, 0.8)),
                total_row: None,
                rule: Some(TableRule {
                    color: accent,
                    extent: LIGHT_BAND_ONE_RULES,
                }),
                header: None,
            }),
            "{style} should paint in accent {}",
            accent_index + 1
        );
    }
}

#[test]
fn an_unresolved_style_family_leaves_the_table_unpainted() {
    // The probe behind issue #1080 shows the Light family's later bands
    // painting something else entirely — Light8 fills its header row solid
    // and rules every row, Light15 boxes the whole table — and the Dark
    // family is not measured at all. Each must fall back to today's
    // behaviour rather than to a guess.
    for style in [
        "TableStyleLight8",
        "TableStyleLight15",
        "TableStyleDark3",
        "MyCustomStyle",
    ] {
        assert_eq!(built_in_table_style(style, &palette()), None, "{style}");
    }
}

#[test]
fn a_light1_table_rules_its_header_and_its_foot_and_nothing_between() {
    let xml =
        module_inventory_table(r#"<tableStyleInfo name="TableStyleLight1" showRowStripes="1"/>"#);

    let style = parse_table_part(&xml, &palette()).expect("Light1 paints its rules");

    let header = style
        .border_at(1, 4)
        .expect("the header row is ruled twice");
    let top: &BorderSide = header.top.as_ref().expect("a rule sits above the header");
    assert_eq!(top.color, Color::black(), "Light1 rules in `dk1`");
    assert_eq!(top.width, 1.0, "a thin Excel border is a 1pt band");
    assert_eq!(top.style, BorderLineStyle::Solid);
    assert!(header.bottom.is_some(), "the header is closed by a rule");
    assert!(
        header.left.is_none() && header.right.is_none(),
        "Light1 draws no verticals"
    );

    assert!(
        style.border_at(1, 5).is_none(),
        "the first body row carries no rule"
    );
    assert!(
        style.border_at(1, 174).is_none(),
        "nor does the second-to-last"
    );

    let foot = style.border_at(1, 175).expect("the table's foot is ruled");
    assert!(foot.bottom.is_some());
    assert!(foot.top.is_none(), "the foot rule hangs below the last row");

    assert!(
        style.border_at(9, 4).is_none(),
        "column I is outside the table"
    );
    assert!(
        style.border_at(1, 176).is_none(),
        "past the table's last row"
    );
}

#[test]
fn a_light1_table_prints_its_header_row_bold() {
    let xml =
        module_inventory_table(r#"<tableStyleInfo name="TableStyleLight1" showRowStripes="1"/>"#);

    let style = parse_table_part(&xml, &palette()).unwrap();

    assert!(style.bolds_header_at(1, 4), "row 4 is the header row");
    assert!(!style.bolds_header_at(1, 5), "row 5 is the first body row");
    assert!(
        !style.bolds_header_at(9, 4),
        "column I is outside the table"
    );
}

#[test]
fn a_light1_table_without_row_stripes_still_rules_and_bolds_its_header() {
    // Excel scopes `showRowStripes` to the banding alone: the same table
    // exported with it off keeps all three rules and the bold header.
    let xml = module_inventory_table(
        r#"<tableStyleInfo name="TableStyleLight1" showRowStripes="0" showColumnStripes="0"/>"#,
    );

    let style = parse_table_part(&xml, &palette()).expect("the rules survive");

    assert_eq!(style.fill_at(1, 5), None, "no band without row stripes");
    assert!(style.border_at(1, 4).is_some());
    assert!(style.bolds_header_at(1, 4));
}

#[test]
fn a_headerless_light1_table_rules_only_its_outer_edges() {
    // With no header row there is nothing to close, so the middle rule has no
    // row to hang on and only the table's own top and bottom print.
    let xml = r#"<?xml version="1.0"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               ref="B2:D10" headerRowCount="0">
          <tableStyleInfo name="TableStyleLight1" showRowStripes="1"/>
        </table>"#;

    let style = parse_table_part(xml, &palette()).unwrap();

    let first = style.border_at(2, 2).expect("the table's top is ruled");
    assert!(first.top.is_some());
    assert!(first.bottom.is_none(), "row 2 closes no header");
    assert!(style.border_at(2, 3).is_none());
    assert!(!style.bolds_header_at(2, 2), "the table has no header row");
}

#[test]
fn a_table_without_a_header_row_stripes_from_its_first_row() {
    let xml = r#"<?xml version="1.0"?>
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               ref="B2:D10" headerRowCount="0">
          <tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>
        </table>"#;

    let stripes = parse_table_part(xml, &palette()).unwrap();

    assert!(
        stripes.fill_at(2, 2).is_some(),
        "row 2 is the first body row"
    );
    assert!(stripes.fill_at(2, 3).is_none());
}
