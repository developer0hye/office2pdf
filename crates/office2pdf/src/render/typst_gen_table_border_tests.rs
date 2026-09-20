use super::*;

#[test]
fn test_table_all_borders() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "All borders".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
            left: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
            right: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("top:"), "Expected top border in: {result}");
    assert!(
        result.contains("bottom:"),
        "Expected bottom border in: {result}"
    );
    assert!(
        result.contains("left:"),
        "Expected left border in: {result}"
    );
    assert!(
        result.contains("right:"),
        "Expected right border in: {result}"
    );
}

#[test]
fn test_table_dashed_border_codegen() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Dashed".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Dashed,
                join: LineJoin::Round,
            }),
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::new(255, 0, 0),
                style: BorderLineStyle::Dotted,
                join: LineJoin::Round,
            }),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("dash: \"dashed\""),
        "Expected dashed dash pattern in: {result}"
    );
    assert!(
        result.contains("dash: \"dotted\""),
        "Expected dotted dash pattern in: {result}"
    );
}

#[test]
fn test_table_double_borders_render_two_oriented_rules() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Double".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 0.8,
                color: Color::new(10, 20, 30),
                style: BorderLineStyle::Double,
                join: LineJoin::Round,
            }),
            bottom: Some(BorderSide {
                width: 0.8,
                color: Color::new(10, 20, 30),
                style: BorderLineStyle::Double,
                join: LineJoin::Round,
            }),
            left: Some(BorderSide {
                width: 0.8,
                color: Color::new(10, 20, 30),
                style: BorderLineStyle::Double,
                join: LineJoin::Round,
            }),
            right: Some(BorderSide {
                width: 0.8,
                color: Color::new(10, 20, 30),
                style: BorderLineStyle::Double,
                join: LineJoin::Round,
            }),
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![TableCell::default(), TableCell::default()],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![TableCell::default(), cell],
                height: None,
            },
        ],
        column_widths: vec![50.0, 50.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let output = generate_typst(&doc).unwrap();
    let result = &output.source;

    assert_eq!(
        result.matches("stroke: 0.8pt + rgb(10, 20, 30)").count(),
        8,
        "each double side should render as two one-width rules: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5.8pt, line(length: 100% + 10pt, angle: 0deg"
        ),
        "the outer horizontal rule should sit one width above the cell edge: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -4.2pt, line(length: 100% + 10pt, angle: 0deg"
        ),
        "the inner horizontal rule should sit one width below the cell edge: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -5.8pt, dy: -5pt, line(length: 100% + 10pt, angle: 90deg"
        ),
        "the outer vertical rule should sit one width before the cell edge: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -4.2pt, dy: -5pt, line(length: 100% + 10pt, angle: 90deg"
        ),
        "the inner vertical rule should sit one width after the cell edge: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 4.2pt, line(length: 100% + 10pt, angle: 0deg"
        ),
        "the inner bottom rule should sit one width above the cell edge: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 5.8pt, line(length: 100% + 10pt, angle: 0deg"
        ),
        "the outer bottom rule should sit one width below the cell edge: {result}"
    );
    assert!(
        result.contains(
            "#place(top + right, dx: 4.2pt, dy: -5pt, line(length: 100% + 10pt, angle: 90deg"
        ),
        "the inner right rule should sit one width before the cell edge: {result}"
    );
    assert!(
        result.contains(
            "#place(top + right, dx: 5.8pt, dy: -5pt, line(length: 100% + 10pt, angle: 90deg"
        ),
        "the outer right rule should sit one width after the cell edge: {result}"
    );

    #[cfg(not(target_arch = "wasm32"))]
    {
        let pdf = crate::render::pdf::compile_to_pdf(
            &output.source,
            &output.images,
            None,
            &[],
            false,
            false,
        )
        .expect("double-border Typst should compile");
        assert!(pdf.starts_with(b"%PDF"));
    }
}

#[test]
fn test_shape_dashed_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Rectangle,
            Some(Color::new(0, 128, 255)),
            Some(BorderSide {
                width: 2.0,
                color: Color::black(),
                style: BorderLineStyle::Dashed,
                join: LineJoin::Round,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    // DrawingML dashes scale with the line width (issue #678): at w=2pt the
    // `dash` preset's 4w/3w becomes 8pt on, 6pt off. Table borders keep the
    // named patterns; only shape strokes take this rule.
    assert!(
        output.source.contains("dash: (8pt, 6pt)"),
        "Expected width-proportional dashed stroke in: {}",
        output.source
    );
}

#[test]
fn test_shape_dash_dot_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Ellipse,
            None,
            Some(BorderSide {
                width: 1.0,
                color: Color::new(0, 0, 255),
                style: BorderLineStyle::DashDot,
                join: LineJoin::Round,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    // At w=1pt the `dashDot` preset's 4w/3w/1w/3w becomes 4pt, 3pt, 1pt, 3pt.
    assert!(
        output.source.contains("dash: (4pt, 3pt, 1pt, 3pt)"),
        "Expected width-proportional dash-dotted stroke in: {}",
        output.source
    );
}

#[test]
fn test_border_line_style_to_typst_mapping() {
    assert_eq!(border_line_style_to_typst(BorderLineStyle::Solid), "solid");
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Dashed),
        "dashed"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Dotted),
        "dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::DashDot),
        "dash-dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::DashDotDot),
        "dash-dotted"
    );
    assert_eq!(border_line_style_to_typst(BorderLineStyle::Double), "solid");
    assert_eq!(border_line_style_to_typst(BorderLineStyle::None), "solid");
}

#[test]
fn test_solid_border_no_dash_param() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Solid".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
            bottom: None,
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("dash:"),
        "Solid border should not have dash parameter in: {result}"
    );
    assert!(
        result.contains("1pt + rgb(0, 0, 0)"),
        "Expected simple solid format in: {result}"
    );
}

// ---------------------------------------------------------------------------
// Boundary-anchored border bands (issue #619)
//
// Excel paints every border as a filled band anchored to the nominal grid
// boundary B (native Excel 16.111 one-factor probe + golden-mock GT traces):
// thin/hair fill [B, B+1], medium [B-1, B+1], thick [B-1, B+2], double two
// 1pt bands [B-1, B] and [B+1, B+2]. Tables flagged
// `TableBorderPaintModel::ExcelBoundaryBands` must realize these bands as offset overlay
// lines instead of Typst cell strokes, which Typst centres on the boundary.
// ---------------------------------------------------------------------------

/// One-cell paragraph content for border tests.
fn bordered_text_cell(text: &str, border: CellBorder) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(border),
        ..TableCell::default()
    }
}

fn solid_side(width: f64) -> BorderSide {
    BorderSide {
        width,
        color: Color::black(),
        style: BorderLineStyle::Solid,
        join: LineJoin::Round,
    }
}

fn boundary_band_table(rows: Vec<TableRow>, column_widths: Vec<f64>) -> Table {
    Table {
        rows,
        column_widths,
        border_paint_model: TableBorderPaintModel::ExcelBoundaryBands,
        ..Table::default()
    }
}

/// A fixed-height row — the spreadsheet default, where the cell frame height
/// is known at codegen and a vertical band can be a single concrete line.
fn fixed_row(cells: Vec<TableCell>) -> TableRow {
    TableRow {
        minimum_height: None,
        cells,
        height: Some(20.0),
    }
}

#[test]
fn test_boundary_band_thin_borders_emit_offset_overlays_not_strokes() {
    let border = CellBorder {
        top: Some(solid_side(1.0)),
        bottom: Some(solid_side(1.0)),
        left: Some(solid_side(1.0)),
        right: Some(solid_side(1.0)),
    };
    let table = boundary_band_table(
        vec![fixed_row(vec![bordered_text_cell("Thin", border)])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("stroke: ("),
        "a boundary-band table must not emit per-cell strokes: {result}"
    );
    // The layout inset still reserves the half border widths of #500/#503, so
    // no text moves relative to the stroke regime.
    assert!(
        result.contains("inset: (top: 5.5pt, right: 5pt, bottom: 5.5pt, left: 5pt)"),
        "border layout inset must be unchanged: {result}"
    );
    assert!(
        !result.contains("#move("),
        "Word's positive-axis content seat must not leak into Excel cells: {result}"
    );
    // Thin band [B, B+1]: a 1pt line whose path is centred at B + 0.5. With
    // the default 5pt padding the top boundary sits at inset.top = 5.5pt above
    // the content box, so dy = -5.5 + 0.5 = -5. Runs extend 1pt past their end
    // boundary, so horizontals span inset.left + 100% + inset.right + 1.
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "top thin band must fill [B, B+1]: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 6pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "bottom thin band must fill [B, B+1] below the boundary: {result}"
    );
    // Vertical bands use the concrete frame height (20pt row) plus the 1pt
    // run extension: a Typst-relative length inside `#place` resolves against
    // the page, not the cell, whenever a spanned row is auto-sized.
    assert!(
        result.contains(
            "#place(top + left, dx: -4.5pt, dy: -5.5pt, line(length: 21pt, angle: 90deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "left thin band must fill [B, B+1] right of the boundary: {result}"
    );
    assert!(
        result.contains(
            "#place(top + right, dx: 5.5pt, dy: -5.5pt, line(length: 21pt, angle: 90deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "right thin band must lie outside the nominal grid rect: {result}"
    );
}

#[test]
fn test_boundary_band_medium_thick_double_weights() {
    let single_top = |side: BorderSide| CellBorder {
        top: Some(side),
        bottom: None,
        left: None,
        right: None,
    };
    let medium_table = boundary_band_table(
        vec![fixed_row(vec![bordered_text_cell(
            "Med",
            single_top(solid_side(2.0)),
        )])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(medium_table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // Medium band [B-1, B+1]: 2pt centred on B; boundary at inset.top = 6pt.
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -6pt, line(length: 100% + 11pt, angle: 0deg, stroke: 2pt + rgb(0, 0, 0)))"
        ),
        "medium band must fill [B-1, B+1]: {result}"
    );

    let thick_table = boundary_band_table(
        vec![fixed_row(vec![bordered_text_cell(
            "Thick",
            single_top(solid_side(3.0)),
        )])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(thick_table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // Thick band [B-1, B+2]: 3pt centred at B + 0.5; boundary at 6.5pt.
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -6pt, line(length: 100% + 11pt, angle: 0deg, stroke: 3pt + rgb(0, 0, 0)))"
        ),
        "thick band must fill [B-1, B+2]: {result}"
    );

    let double_table = boundary_band_table(
        vec![fixed_row(vec![bordered_text_cell(
            "Double",
            single_top(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Double,
                join: LineJoin::Round,
            }),
        )])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(double_table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // Double: two 1pt bands [B-1, B] and [B+1, B+2] with the boundary strip
    // [B, B+1] as the gap; boundary at inset.top = 5.5pt.
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -6pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "outer double band must fill [B-1, B]: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -4pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "inner double band must fill [B+1, B+2]: {result}"
    );
}

/// A fitted sheet's boundary-band table: sizes already multiplied by the
/// print scale, with the scale riding on the table as the parser leaves it.
fn fitted_boundary_band_table(rows: Vec<TableRow>, scale: f64) -> Table {
    Table {
        seats_bottom_aligned_text_on_descender: true,
        print_scale: Some(scale),
        ..boundary_band_table(rows, vec![100.0])
    }
}

#[test]
fn test_fitted_sheet_scales_thin_bands_and_run_extension_with_the_print_scale() {
    // Native Excel 16.112 export of the fit-to-height budget sheet at its
    // 0.78 auto-fit scale (issue #1564): every thin rule is a 0.78pt filled
    // band `[B, B + 0.78]` running 0.78pt past its end boundary, i.e. the
    // declared 1pt band scaled with the rest of the sheet.
    let border = CellBorder {
        top: Some(solid_side(1.0)),
        bottom: Some(solid_side(1.0)),
        left: Some(solid_side(1.0)),
        right: Some(solid_side(1.0)),
    };
    let table = fitted_boundary_band_table(
        vec![fixed_row(vec![bordered_text_cell("Thin", border)])],
        0.78,
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // The layout inset keeps the declared half widths: the seat model of
    // #1545 was calibrated against it, and Excel scales only the paint.
    assert!(
        result.contains("inset: (top: 5.5pt, right: 5pt, bottom: 5.5pt, left: 5pt)"),
        "border layout inset must stay at the declared half widths: {result}"
    );
    // Top band [B, B+0.78]: centre at B + 0.39 from the boundary at
    // inset.top = 5.5pt, so dy = -5.5 + 0.39; the run extension is 0.78.
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5.11pt, line(length: 100% + 10.78pt, angle: 0deg, stroke: 0.78pt + rgb(0, 0, 0)))"
        ),
        "top thin band must fill [B, B+0.78]: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 5.89pt, line(length: 100% + 10.78pt, angle: 0deg, stroke: 0.78pt + rgb(0, 0, 0)))"
        ),
        "bottom thin band must fill [B, B+0.78] below the boundary: {result}"
    );
    // Vertical bands: the 20pt fixed row plus the scaled 0.78pt extension.
    assert!(
        result.contains(
            "#place(top + left, dx: -4.61pt, dy: -5.5pt, line(length: 20.78pt, angle: 90deg, stroke: 0.78pt + rgb(0, 0, 0)))"
        ),
        "left thin band must fill [B, B+0.78] right of the boundary: {result}"
    );
    assert!(
        result.contains(
            "#place(top + right, dx: 5.39pt, dy: -5.5pt, line(length: 20.78pt, angle: 90deg, stroke: 0.78pt + rgb(0, 0, 0)))"
        ),
        "right thin band must fill [B, B+0.78] past the boundary: {result}"
    );
    assert!(
        !result.contains("stroke: 1pt + rgb(0, 0, 0)"),
        "no band may keep the unscaled 1pt weight on a fitted sheet: {result}"
    );
}

#[test]
fn test_fitted_sheet_scales_medium_thick_double_bands_with_the_print_scale() {
    // A different scale than the reported sheet's, so the rule is the
    // multiplication and not a constant: at 0.5 the medium band [B-1, B+1]
    // becomes a 1pt rule centred on B, the thick band [B-1, B+2] a 1.5pt rule
    // centred at B + 0.25, and a double's two 1pt bands become 0.5pt rules
    // centred at B - 0.25 and B + 0.75.
    let single_top = |side: BorderSide| CellBorder {
        top: Some(side),
        bottom: None,
        left: None,
        right: None,
    };
    let render = |side: BorderSide| -> String {
        let table = fitted_boundary_band_table(
            vec![fixed_row(vec![bordered_text_cell("Fit", single_top(side))])],
            0.5,
        );
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        generate_typst(&doc).unwrap().source
    };

    let medium = render(solid_side(2.0));
    // Boundary at inset.top = 6pt (declared 1pt half width).
    assert!(
        medium.contains(
            "#place(top + left, dx: -5pt, dy: -6pt, line(length: 100% + 10.5pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "medium band must fill [B-0.5, B+0.5]: {medium}"
    );

    let thick = render(solid_side(3.0));
    // Boundary at inset.top = 6.5pt; centre at B + 0.25.
    assert!(
        thick.contains(
            "#place(top + left, dx: -5pt, dy: -6.25pt, line(length: 100% + 10.5pt, angle: 0deg, stroke: 1.5pt + rgb(0, 0, 0)))"
        ),
        "thick band must fill [B-0.5, B+1]: {thick}"
    );

    let double = render(BorderSide {
        width: 1.0,
        color: Color::black(),
        style: BorderLineStyle::Double,
        join: LineJoin::Round,
    });
    // Boundary at inset.top = 5.5pt; rules at B - 0.25 and B + 0.75.
    assert!(
        double.contains(
            "#place(top + left, dx: -5pt, dy: -5.75pt, line(length: 100% + 10.5pt, angle: 0deg, stroke: 0.5pt + rgb(0, 0, 0)))"
        ),
        "outer double band must fill [B-0.5, B]: {double}"
    );
    assert!(
        double.contains(
            "#place(top + left, dx: -5pt, dy: -4.75pt, line(length: 100% + 10.5pt, angle: 0deg, stroke: 0.5pt + rgb(0, 0, 0)))"
        ),
        "inner double band must fill [B+0.5, B+1]: {double}"
    );
}

#[test]
fn test_boundary_band_shared_edge_paints_once() {
    // Both neighbours declare the same internal boundary: it must paint once.
    let upper = bordered_text_cell(
        "Upper",
        CellBorder {
            top: None,
            bottom: Some(solid_side(1.0)),
            left: None,
            right: None,
        },
    );
    let lower = bordered_text_cell(
        "Lower",
        CellBorder {
            top: Some(solid_side(1.0)),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let table = boundary_band_table(
        vec![fixed_row(vec![upper]), fixed_row(vec![lower])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert_eq!(
        result.matches("angle: 0deg").count(),
        1,
        "an edge declared by both neighbours must paint exactly one band: {result}"
    );
    assert!(
        !result.contains("#place(bottom"),
        "on an equal-weight tie the lower cell's top declaration paints: {result}"
    );
}

#[test]
fn test_boundary_band_shared_edge_heavier_declaration_wins() {
    // Excel resolves conflicting declarations to the heavier style: a medium
    // bottom must beat the neighbour's thin top.
    let upper = bordered_text_cell(
        "Upper",
        CellBorder {
            top: None,
            bottom: Some(solid_side(2.0)),
            left: None,
            right: None,
        },
    );
    let lower = bordered_text_cell(
        "Lower",
        CellBorder {
            top: Some(solid_side(1.0)),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let table = boundary_band_table(
        vec![fixed_row(vec![upper]), fixed_row(vec![lower])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert_eq!(
        result.matches("angle: 0deg").count(),
        1,
        "a conflicting edge must still paint exactly one band: {result}"
    );
    // Medium bottom boundary sits at inset.bottom = 5 + 1 = 6pt, band centred
    // on it.
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 6pt, line(length: 100% + 11pt, angle: 0deg, stroke: 2pt + rgb(0, 0, 0)))"
        ),
        "the heavier (medium) declaration must paint the shared edge: {result}"
    );
}

#[test]
fn test_boundary_band_shared_vertical_edge_paints_once() {
    let left_cell = bordered_text_cell(
        "L",
        CellBorder {
            top: None,
            bottom: None,
            left: None,
            right: Some(solid_side(1.0)),
        },
    );
    let right_cell = bordered_text_cell(
        "R",
        CellBorder {
            top: None,
            bottom: None,
            left: Some(solid_side(1.0)),
            right: None,
        },
    );
    let table = boundary_band_table(
        vec![fixed_row(vec![left_cell, right_cell])],
        vec![100.0, 100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert_eq!(
        result.matches("angle: 90deg").count(),
        1,
        "a vertical edge declared by both neighbours must paint once: {result}"
    );
    assert!(
        !result.contains("#place(top + right"),
        "on an equal-weight tie the right cell's left declaration paints: {result}"
    );
}

#[test]
fn test_boundary_band_patterned_style_keeps_dash_dict() {
    let table = boundary_band_table(
        vec![fixed_row(vec![bordered_text_cell(
            "Dashed",
            CellBorder {
                top: Some(BorderSide {
                    width: 1.0,
                    color: Color::black(),
                    style: BorderLineStyle::Dashed,
                    join: LineJoin::Round,
                }),
                bottom: None,
                left: None,
                right: None,
            },
        )])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5pt, line(length: 100% + 11pt, angle: 0deg, stroke: (paint: rgb(0, 0, 0), thickness: 1pt, dash: \"dashed\")))"
        ),
        "a patterned band must keep its dash dict on the overlay line: {result}"
    );
    assert!(
        !result.contains("stroke: (top"),
        "a patterned band must not also emit a cell stroke: {result}"
    );
}

#[test]
fn test_unflagged_table_keeps_centred_strokes_byte_identically() {
    // Synthetic and PowerPoint tables with the default model keep the exact
    // stroke emission. DOCX selects WordPositiveAxisBands after #724; the #619
    // and #724 band regimes must not leak into unflagged tables.
    let border = CellBorder {
        top: Some(solid_side(1.0)),
        bottom: Some(solid_side(1.0)),
        left: Some(solid_side(1.0)),
        right: Some(solid_side(1.0)),
    };
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![bordered_text_cell("Word", border)],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(
            "stroke: (top: 1pt + rgb(0, 0, 0), bottom: 1pt + rgb(0, 0, 0), left: 1pt + rgb(0, 0, 0), right: 1pt + rgb(0, 0, 0))"
        ),
        "unflagged tables must keep the centred cell strokes: {result}"
    );
    assert!(
        !result.contains("#place("),
        "unflagged solid borders must not paint overlays: {result}"
    );
    assert!(
        !result.contains("#move("),
        "unflagged and PowerPoint tables must keep their content seat: {result}"
    );
}

/// Writer seats every DOCX table cell's content one tenth of a point into the
/// positive x side of the cell track, even when the table is borderless. The
/// seat belongs to the content box rather than to its paragraph alignment:
/// left-, centre-, and right-aligned lines all move by the same amount while
/// their baselines, column widths, and cell margins stay fixed (issue #1488).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_borderless_word_cells_share_the_writer_x_origin_seat() {
    fn table_source(model: TableBorderPaintModel) -> String {
        let cells = [
            ("LEFT", Some(Alignment::Left)),
            ("CENTER", Some(Alignment::Center)),
            ("RIGHT", Some(Alignment::Right)),
        ]
        .into_iter()
        .map(|(text, alignment)| TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    alignment,
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: text.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            ..TableCell::default()
        })
        .collect();
        let table = Table {
            rows: vec![TableRow {
                minimum_height: None,
                cells,
                height: Some(24.0),
            }],
            column_widths: vec![100.0, 100.0, 100.0],
            default_cell_padding: Some(Insets {
                top: 0.0,
                right: 5.4,
                bottom: 0.0,
                left: 5.4,
            }),
            border_paint_model: model,
            ..Table::default()
        };
        generate_typst(&make_doc(vec![make_flow_page(vec![Block::Table(table)])]))
            .unwrap()
            .source
    }

    let neutral_source = table_source(TableBorderPaintModel::CenteredStroke);
    let word_source = table_source(TableBorderPaintModel::WordPositiveAxisBands);
    let neutral_runs =
        crate::render::pdf::compiled_text_runs(&neutral_source, 0).unwrap_or_else(|error| {
            panic!("neutral table failed to compile: {error}\n{neutral_source}")
        });
    let word_runs = crate::render::pdf::compiled_text_runs(&word_source, 0)
        .unwrap_or_else(|error| panic!("Word table failed to compile: {error}\n{word_source}"));

    for text in ["LEFT", "CENTER", "RIGHT"] {
        let neutral = neutral_runs
            .iter()
            .find(|run| run.text == text)
            .unwrap_or_else(|| panic!("missing {text:?} in {neutral_runs:?}"));
        let word = word_runs
            .iter()
            .find(|run| run.text == text)
            .unwrap_or_else(|| panic!("missing {text:?} in {word_runs:?}"));
        assert!(
            (word.left_pt - neutral.left_pt - 0.1).abs() < 0.001,
            "{text} must move exactly 0.10pt right: neutral={}, Word={}\n{word_source}",
            neutral.left_pt,
            word.left_pt,
        );
        assert!(
            (word.baseline_pt - neutral.baseline_pt).abs() < 0.001,
            "{text} must keep its baseline: neutral={}, Word={}\n{word_source}",
            neutral.baseline_pt,
            word.baseline_pt,
        );
    }

    assert_eq!(
        word_source.matches("#move(dx: 0.1pt, dy: 0pt)[").count(),
        3,
        "each Word cell gets the same visual seat: {word_source}"
    );
    assert!(
        word_source.contains("columns: (100pt, 100pt, 100pt)")
            && word_source.contains("inset: (top: 0pt, right: 5.4pt, bottom: 0pt, left: 5.4pt)"),
        "the seat must not rewrite table width or cell margins: {word_source}"
    );
}

#[test]
fn test_word_bands_quantize_and_paint_on_the_positive_axis() {
    let border = CellBorder {
        top: Some(solid_side(0.5)),
        bottom: Some(solid_side(0.5)),
        left: Some(solid_side(0.5)),
        right: Some(solid_side(0.5)),
    };
    let table = Table {
        rows: vec![fixed_row(vec![bordered_text_cell("Word", border)])],
        column_widths: vec![100.0],
        border_paint_model: TableBorderPaintModel::WordPositiveAxisBands,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("stroke: ("),
        "Word borders must not use centred table-cell strokes: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5.25pt, rect(width: 100% + 10pt, height: 0.48pt, fill: rgb(0, 0, 0), stroke: none))"
        ),
        "the top 0.48pt band must start at the boundary and paint down: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 5.73pt, rect(width: 100% + 10pt, height: 0.48pt, fill: rgb(0, 0, 0), stroke: none))"
        ),
        "the bottom band must start at the boundary and paint down: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5.25pt, rect(width: 0.48pt, height: 20pt, fill: rgb(0, 0, 0), stroke: none))"
        ),
        "the left band must start at the boundary and paint right: {result}"
    );
    assert!(
        result.contains(
            "#place(top + right, dx: 5.48pt, dy: -5.25pt, rect(width: 0.48pt, height: 20pt, fill: rgb(0, 0, 0), stroke: none))"
        ),
        "the right band must start at the boundary and paint right: {result}"
    );
    assert!(
        result.contains("#move(dx: 0.34pt, dy: 0.48pt)[Word]"),
        "Writer's borderless 0.10pt x seat composes with half a painted left \
         border inward and one painted top border down without changing row \
         or column layout: {result}"
    );
}

#[test]
fn test_word_shared_boundary_starts_inside_the_following_cell() {
    let left_cell = bordered_text_cell(
        "L",
        CellBorder {
            top: None,
            bottom: None,
            left: None,
            right: Some(solid_side(1.25)),
        },
    );
    let right_cell = bordered_text_cell(
        "R",
        CellBorder {
            top: None,
            bottom: None,
            left: Some(solid_side(1.25)),
            right: None,
        },
    );
    let table = Table {
        rows: vec![fixed_row(vec![left_cell, right_cell])],
        column_widths: vec![100.0, 100.0],
        border_paint_model: TableBorderPaintModel::WordPositiveAxisBands,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert_eq!(
        result.matches("width: 1.2pt, height: 20pt").count(),
        1,
        "a shared Word boundary must paint once: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5pt, rect(width: 1.2pt, height: 20pt, fill: rgb(0, 0, 0), stroke: none))"
        ),
        "the following cell must own the 1.2pt band on its interior side: {result}"
    );
}

#[test]
fn test_word_auto_row_bottom_twin_anchors_at_the_cell_boundary() {
    let cell = bordered_text_cell(
        "Auto",
        CellBorder {
            top: None,
            bottom: None,
            left: Some(solid_side(0.5)),
            right: None,
        },
    );
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        border_paint_model: TableBorderPaintModel::WordPositiveAxisBands,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#place(bottom + left, dx: -5pt, dy: 5pt, rect(width: 0.48pt, height:"),
        "the lower twin's bottom edge must sit on the cell boundary: {result}"
    );
}

#[test]
fn test_word_repeating_header_boundary_starts_inside_the_body_row() {
    let upper = bordered_text_cell(
        "Header",
        CellBorder {
            top: None,
            bottom: Some(solid_side(0.5)),
            left: None,
            right: None,
        },
    );
    let lower = bordered_text_cell(
        "Body",
        CellBorder {
            top: Some(solid_side(0.5)),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let table = Table {
        rows: vec![fixed_row(vec![upper]), fixed_row(vec![lower])],
        column_widths: vec![100.0],
        header_row_count: 1,
        border_paint_model: TableBorderPaintModel::WordPositiveAxisBands,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5.25pt, rect(width: 100% + 10pt, height: 0.48pt, fill: rgb(0, 0, 0), stroke: none))#move(dx: 0.1pt, dy: 0.48pt)[Body]"
        ),
        "the Word body row must own the repeated-header boundary and seat its \
         content below that band: {result}"
    );
    assert!(
        !result.contains("#place(bottom + left"),
        "the Word header must not paint the shared rule upward: {result}"
    );
}

/// An auto-sized row's final height only the renderer knows, and a
/// Typst-relative length inside `#place` resolves against the page there
/// (issue #619 probe), so vertical bands in auto rows are painted as two
/// concrete twin bands anchored at the cell's top and bottom edges, sized
/// from the row's tallest single-line box.
#[test]
fn test_boundary_band_auto_row_verticals_paint_concrete_twin_bands() {
    let Some(line_box) = word_cell_line_box(
        &[Run {
            text: "Auto".to_string(),
            style: TextStyle {
                font_family: Some("Libertinus Serif".to_string()),
                font_size: Some(10.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
        &ParagraphStyle::default(),
        None,
        RowEastAsianMetrics {
            has_east_asian_text: false,
            takes_east_asian_metrics: false,
        },
        None,
        false,
        None,
        None,
        // `boundary_band_table` leaves the spreadsheet marker off, which is
        // what the codegen passes for this table.
        None,
    ) else {
        return; // no font book available (e.g. exotic CI sandbox)
    };
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Auto".to_string(),
                style: TextStyle {
                    font_family: Some("Libertinus Serif".to_string()),
                    font_size: Some(10.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: None,
            bottom: None,
            left: Some(solid_side(1.0)),
            right: None,
        }),
        ..TableCell::default()
    };
    let table = boundary_band_table(
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // The frame estimate is the single-line box plus the cell's insets (5pt
    // padding each side; no half border on top/bottom for a left-only
    // border), and the band adds the 1pt run extension.
    let frame_estimate_pt: f64 =
        (line_box.top_em + line_box.bottom_em) * line_box.font_size_pt + 5.0 + 5.0;
    let band_length: String = tables::format_geometry(frame_estimate_pt + 1.0);
    assert!(
        result.contains(&format!(
            "#place(top + left, dx: -4.5pt, dy: -5pt, line(length: {band_length}pt, angle: 90deg, stroke: 1pt + rgb(0, 0, 0)))"
        )),
        "the top twin must hang from the top boundary: {result}"
    );
    assert!(
        result.contains(&format!(
            "#place(bottom + left, dx: -4.5pt, dy: 6pt, line(length: {band_length}pt, angle: -90deg, stroke: 1pt + rgb(0, 0, 0)))"
        )),
        "the bottom twin must rise from 1pt past the bottom boundary: {result}"
    );
}

/// Without line metrics the twins fall back to the ambient text size,
/// following the data-bar `1.2em` precedent.
#[test]
fn test_boundary_band_auto_row_verticals_em_fallback_without_metrics() {
    // Default `TextStyle` declares no font family, so no line metrics exist.
    let cell = bordered_text_cell(
        "NoMetrics",
        CellBorder {
            top: None,
            bottom: None,
            left: Some(solid_side(1.0)),
            right: None,
        },
    );
    let table = boundary_band_table(
        vec![TableRow {
            minimum_height: None,
            cells: vec![cell],
            height: None,
        }],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(
            "#place(top + left, dx: -4.5pt, dy: -5pt, line(length: 1.2em + 11pt, angle: 90deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "the top twin must fall back to an em-sized band: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -4.5pt, dy: 6pt, line(length: 1.2em + 11pt, angle: -90deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "the bottom twin must fall back to an em-sized band: {result}"
    );
}

#[test]
fn test_boundary_band_double_declaration_survives_thin_neighbour() {
    // Excel's conflict rule ranks a double rule above every single band even
    // though each of its strokes is stored at the thin 1pt weight: A1
    // bottom=double against A2 top=thin must paint the double's two bands
    // (issue #619 review, remediation 1).
    let upper = bordered_text_cell(
        "Upper",
        CellBorder {
            top: None,
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Double,
                join: LineJoin::Round,
            }),
            left: None,
            right: None,
        },
    );
    let lower = bordered_text_cell(
        "Lower",
        CellBorder {
            top: Some(solid_side(1.0)),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let table = boundary_band_table(
        vec![fixed_row(vec![upper]), fixed_row(vec![lower])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert_eq!(
        result.matches("angle: 0deg").count(),
        2,
        "the double must paint both of its bands, erasing the thin: {result}"
    );
    assert_eq!(
        result.matches("#place(bottom + left").count(),
        2,
        "both bands must come from the upper cell's double bottom: {result}"
    );
    assert!(
        !result.contains("#place(top + left, dx: -5pt, dy: -5pt"),
        "the lower cell's thin top must yield to the double: {result}"
    );
}

#[test]
fn test_boundary_band_solid_thin_outranks_hair_at_equal_width() {
    // `hair` shares `thin`'s 1pt band width and differs only in its dotted
    // texture, so a raw width comparison ties; Excel keeps the solid rule
    // (issue #619 review, remediation 1).
    let upper = bordered_text_cell(
        "Upper",
        CellBorder {
            top: None,
            bottom: Some(solid_side(1.0)),
            left: None,
            right: None,
        },
    );
    let lower = bordered_text_cell(
        "Lower",
        CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Dotted,
                join: LineJoin::Round,
            }),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let table = boundary_band_table(
        vec![fixed_row(vec![upper]), fixed_row(vec![lower])],
        vec![100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert_eq!(
        result.matches("angle: 0deg").count(),
        1,
        "the conflicting edge must paint exactly one band: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 6pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "the solid thin declaration must paint the shared edge: {result}"
    );
    assert!(
        !result.contains("dash"),
        "the patterned hair declaration must yield to the solid: {result}"
    );
}

#[test]
fn test_boundary_band_header_body_tie_paints_from_repeating_header() {
    // A print-title header row repeats on every page while the first body
    // row renders once. On an equal-rank tie the band must therefore be
    // emitted from the header's bottom slot — inside `table.header(...)` —
    // so pages 2+ keep the rule under the repeated header (issue #619
    // review, remediation 2).
    let header_cell = bordered_text_cell(
        "Head",
        CellBorder {
            top: None,
            bottom: Some(solid_side(1.0)),
            left: None,
            right: None,
        },
    );
    let body_cell = bordered_text_cell(
        "Body",
        CellBorder {
            top: Some(solid_side(1.0)),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let mut table = boundary_band_table(
        vec![fixed_row(vec![header_cell]), fixed_row(vec![body_cell])],
        vec![100.0],
    );
    table.header_row_count = 1;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert_eq!(
        result.matches("angle: 0deg").count(),
        1,
        "the shared header/body edge must still paint exactly once: {result}"
    );
    let header_pos: usize = result
        .find("table.header(")
        .expect("header block must exist");
    let band_pos: usize = result.find("angle: 0deg").expect("band must exist");
    let body_pos: usize = result.find("Body]").expect("body cell must exist");
    assert!(
        header_pos < band_pos && band_pos < body_pos,
        "the tied band must be emitted within the table.header block, \
         before the first body cell: {result}"
    );
}

#[test]
fn test_boundary_band_header_body_heavier_body_band_repeats_with_header() {
    // When the first body row declares a strictly heavier rule than the
    // repeating header above it, the body's band wins — but the header side
    // must also carry it, or the repeated header instances on pages 2+ would
    // lose the rule (issue #619 review, remediation 2).
    let header_cell = bordered_text_cell(
        "Head",
        CellBorder {
            top: None,
            bottom: Some(solid_side(1.0)),
            left: None,
            right: None,
        },
    );
    let body_cell = bordered_text_cell(
        "Body",
        CellBorder {
            top: Some(solid_side(2.0)),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let mut table = boundary_band_table(
        vec![fixed_row(vec![header_cell]), fixed_row(vec![body_cell])],
        vec![100.0],
    );
    table.header_row_count = 1;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert_eq!(
        result.matches("angle: 0deg").count(),
        2,
        "the heavier body band must paint from both sides of the boundary: {result}"
    );
    assert_eq!(
        result
            .matches("angle: 0deg, stroke: 2pt + rgb(0, 0, 0)")
            .count(),
        2,
        "both emissions must carry the body's heavier (medium) style: {result}"
    );
    // Each cell's overlays precede its text within the cell bracket, so a
    // band before "Head]" is the header cell's and one after it is the body
    // cell's.
    let head_pos: usize = result.find("Head]").expect("header cell must exist");
    let first_band_pos: usize = result.find("angle: 0deg").expect("band must exist");
    let last_band_pos: usize = result.rfind("angle: 0deg").expect("band must exist");
    assert!(
        first_band_pos < head_pos && head_pos < last_band_pos,
        "one emission must sit in the header block and one in the body row: {result}"
    );
}

#[test]
fn test_boundary_band_auto_row_frame_estimate_computed_once_per_row() {
    // The frame estimate walks every cell in the row; computing it per cell
    // makes vertical-band preparation O(cells^2) in wide wrap-text rows.
    // It must be computed at most once per auto-sized row (issue #619
    // review, remediation 3).
    let vertical_border = CellBorder {
        top: None,
        bottom: None,
        left: Some(solid_side(1.0)),
        right: Some(solid_side(1.0)),
    };
    let cells: Vec<TableCell> = (0..6)
        .map(|i| bordered_text_cell(&format!("C{i}"), vertical_border.clone()))
        .collect();
    let table = boundary_band_table(
        vec![
            TableRow {
                minimum_height: None,
                cells: cells.clone(),
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells,
                height: None,
            },
        ],
        vec![50.0; 6],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);

    tables::AUTO_ROW_FRAME_ESTIMATE_CALLS.with(|calls| calls.set(0));
    let result = generate_typst(&doc).unwrap().source;
    let estimate_calls: usize = tables::AUTO_ROW_FRAME_ESTIMATE_CALLS.with(|calls| calls.get());

    assert!(
        result.contains("angle: 90deg"),
        "the rows must actually paint vertical bands: {result}"
    );
    assert!(
        estimate_calls <= 2,
        "the per-row frame estimate must be computed at most once per row \
         (2 rows), got {estimate_calls} calls"
    );
}

// ---------------------------------------------------------------------------
// Printed gridlines (issue #622)
//
// `<printOptions gridLines="1"/>` prints Excel's gridline on every cell
// boundary of the printed range. Measured on native Excel exports of the
// NumberFormatTests fixture (/Volumes/T7/scratch/issue-622/nft2-p1.rects.txt,
// nft2-p2.trace): every gridline is a fill band exactly 1.0pt thick, pure
// black, boundary-anchored [B, B+1] toward +x/+y — the same convention as the
// #619 thin border band. Any explicit border outranks the gridline on its
// boundary (a hair border replaces the black gridline at C337), and a cell
// fill suppresses all four adjacent gridline segments.
// ---------------------------------------------------------------------------

/// A borderless text cell, the shape most sheet cells have.
fn plain_text_cell(text: &str) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        ..TableCell::default()
    }
}

fn gridline_table(rows: Vec<TableRow>, column_widths: Vec<f64>) -> Table {
    Table {
        rows,
        column_widths,
        border_paint_model: TableBorderPaintModel::ExcelBoundaryBands,
        prints_gridlines: true,
        ..Table::default()
    }
}

#[test]
fn test_printed_gridlines_rule_every_boundary_at_measured_geometry() {
    let table = gridline_table(vec![fixed_row(vec![plain_text_cell("A1")])], vec![100.0]);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // Gridline = 1pt pure black band [B, B+1]: with the default 5pt padding
    // the top boundary sits at inset.top = 5pt, so the band's centre line is
    // at dy = -5 + 0.5 = -4.5pt, exactly the #619 thin-border geometry.
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -4.5pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "top gridline must fill [B, B+1] below the top boundary: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 5.5pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "bottom gridline must fill [B, B+1] below the bottom boundary: {result}"
    );
    assert!(
        result.contains(
            "#place(top + left, dx: -4.5pt, dy: -5pt, line(length: 21pt, angle: 90deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "left gridline must fill [B, B+1] right of the left boundary: {result}"
    );
    assert!(
        result.contains(
            "#place(top + right, dx: 5.5pt, dy: -5pt, line(length: 21pt, angle: 90deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "right gridline must fill [B, B+1] right of the right boundary: {result}"
    );
}

#[test]
fn test_printed_gridlines_paint_interior_boundaries_from_both_sides() {
    // GT closes the grid at every page break: the row above the break draws
    // the bottom rule. Which row breaks a page only the renderer knows, so
    // every cell paints its own bottom (and right-columnless top) band; the
    // two seeds of an interior boundary are boundary-anchored to the same
    // [B, B+1] strip and coincide invisibly.
    let table = gridline_table(
        vec![
            fixed_row(vec![plain_text_cell("A1"), plain_text_cell("B1")]),
            fixed_row(vec![plain_text_cell("A2"), plain_text_cell("B2")]),
        ],
        vec![100.0, 100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert_eq!(
        result.matches("angle: 0deg").count(),
        8,
        "each of the 4 cells must paint its top and bottom gridline: {result}"
    );
    assert_eq!(
        result.matches("angle: 90deg").count(),
        8,
        "each of the 4 cells must paint its left and right gridline: {result}"
    );
}

#[test]
fn test_explicit_border_outranks_gridline_on_its_boundary() {
    // A medium top border owns its boundary: no black 1pt gridline may paint
    // there, while the other three boundaries keep theirs.
    let bordered = bordered_text_cell(
        "Med",
        CellBorder {
            top: Some(solid_side(2.0)),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let table = gridline_table(vec![fixed_row(vec![bordered])], vec![100.0]);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // Medium band [B-1, B+1] centred on the boundary at inset.top = 6pt.
    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -6pt, line(length: 100% + 11pt, angle: 0deg, stroke: 2pt + rgb(0, 0, 0)))"
        ),
        "the explicit medium border must paint its boundary: {result}"
    );
    assert_eq!(
        result.matches("angle: 0deg").count(),
        2,
        "the top boundary must carry only the medium band, the bottom only \
         its gridline: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 5.5pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt + rgb(0, 0, 0)))"
        ),
        "the undeclared bottom boundary must keep its gridline: {result}"
    );
}

#[test]
fn test_hair_border_replaces_gridline_not_the_reverse() {
    // GT: C337's hair borders replace the black gridline at their boundary
    // even though a solid rule outranks a patterned one in the #619 conflict
    // rank — the gridline is below every explicit declaration, not a peer.
    let haired = bordered_text_cell(
        "Hair",
        CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Dotted,
                join: LineJoin::Round,
            }),
            bottom: None,
            left: None,
            right: None,
        },
    );
    let table = gridline_table(vec![fixed_row(vec![haired])], vec![100.0]);
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(
            "#place(top + left, dx: -5pt, dy: -5pt, line(length: 100% + 11pt, angle: 0deg, stroke: (paint: rgb(0, 0, 0), thickness: 1pt, dash: \"dotted\")))"
        ),
        "the hair border must paint its boundary: {result}"
    );
    assert_eq!(
        result.matches("angle: 0deg").count(),
        2,
        "no solid gridline may double the hair boundary: {result}"
    );
    assert!(
        !result.contains("dy: -5pt, line(length: 100% + 11pt, angle: 0deg, stroke: 1pt"),
        "the gridline must yield to the hair border: {result}"
    );
}

#[test]
fn test_cell_fill_suppresses_adjacent_gridlines() {
    // GT (Tests p1 vs the fill-free p2): a cell fill suppresses all four
    // adjacent gridline segments — Excel truncates the verticals at the
    // filled row and omits the horizontal at its bottom boundary.
    let filled = TableCell {
        background: Some(Color::new(237, 125, 49)),
        ..plain_text_cell("Filled")
    };
    let table = gridline_table(
        vec![fixed_row(vec![filled, plain_text_cell("Plain")])],
        vec![100.0, 100.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // Filled cell: no bands at all. Plain cell: its left boundary abuts the
    // fill and is suppressed too; top, bottom, and right survive.
    assert_eq!(
        result.matches("angle: 0deg").count(),
        2,
        "only the plain cell's top and bottom gridlines may paint: {result}"
    );
    assert_eq!(
        result.matches("angle: 90deg").count(),
        1,
        "only the plain cell's right gridline may paint: {result}"
    );
    assert!(
        result.contains("#place(top + right, dx: 5.5pt"),
        "the surviving vertical must be the plain cell's right band: {result}"
    );
}

#[test]
fn test_gridlines_repeat_with_the_print_title_header() {
    // A print-title header repeats on every page; its own top and bottom
    // gridline seeds must repeat with it so the grid stays closed under the
    // header on pages 2+.
    let mut table = gridline_table(
        vec![
            fixed_row(vec![plain_text_cell("Head")]),
            fixed_row(vec![plain_text_cell("Body")]),
        ],
        vec![100.0],
    );
    table.header_row_count = 1;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    let header_pos: usize = result
        .find("table.header(")
        .expect("header block must exist");
    let head_text_pos: usize = result.find("Head]").expect("header cell must exist");
    let header_cell: &str = &result[header_pos..head_text_pos];
    assert_eq!(
        header_cell.matches("angle: 0deg").count(),
        2,
        "the header cell must carry its top and bottom gridline bands inside \
         the repeating header block: {result}"
    );
}

#[test]
fn test_gridlines_absent_without_the_flag() {
    // The same sheet without `printOptions gridLines` prints no gridlines at
    // all: the native-export probe workbooks measured for the #621 column
    // model declare no printOptions element and their GT traces carry zero
    // gridline primitives, so the flag strictly gates printing (#622).
    let mut unflagged = gridline_table(vec![fixed_row(vec![plain_text_cell("A1")])], vec![100.0]);
    unflagged.prints_gridlines = false;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(unflagged)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#place("),
        "a borderless sheet without the flag must paint nothing: {result}"
    );

    // The gridline convention is measured only for Excel's boundary-band
    // regime; a centred-stroke table outside it must ignore the flag.
    let mut word_style = gridline_table(vec![fixed_row(vec![plain_text_cell("W")])], vec![100.0]);
    word_style.border_paint_model = TableBorderPaintModel::CenteredStroke;
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(word_style)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#place("),
        "gridlines must not leak outside the boundary-band regime: {result}"
    );
}

/// A print-heading sheet keeps both declarations at a horizontal boundary.
///
/// Codegen cannot see the page breaks Typst chooses, so a boundary painted by
/// only one of its two owners is closed on only one side of a break. At a tie
/// the resolver hands the boundary to the top owner — the row *below* — which
/// on an intermediate page is the first row of the next page, leaving the
/// previous page's bottom edge open across the row-number gutter (issue #722).
///
/// Word tables keep the single-owner resolution, which the second half asserts:
/// their border geometry is calibrated against it.
#[test]
fn test_print_heading_boundary_keeps_both_coincident_bands() {
    let rule = || BorderSide {
        width: 1.0,
        color: Color::black(),
        style: BorderLineStyle::Solid,
        join: LineJoin::Round,
    };
    let cell = |text: &str| TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        // Both rows declare the same rule at the boundary between them.
        border: Some(CellBorder {
            top: Some(rule()),
            bottom: Some(rule()),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = |prints_headings: bool| Table {
        rows: vec![
            TableRow {
                minimum_height: None,
                cells: vec![cell("1")],
                height: None,
            },
            TableRow {
                minimum_height: None,
                cells: vec![cell("2")],
                height: None,
            },
        ],
        column_widths: vec![60.0],
        prints_headings,
        ..Table::default()
    };

    let painted_bottom = |prints_headings: bool| -> bool {
        super::tables::resolve_boundary_painted_borders(&table(prints_headings), 1, &[])[0][0]
            .as_ref()
            .is_some_and(|border| border.bottom.is_some())
    };

    assert!(
        painted_bottom(true),
        "a print-heading sheet must keep the upper row's bottom band, so a page \
         break at that boundary still closes the page"
    );
    assert!(
        !painted_bottom(false),
        "an ordinary table must still resolve the tie to a single owner"
    );
}

// ---------------------------------------------------------------------------
// Excel background bleed (issue #1190)
//
// The shading under those bands follows the same boundary convention: Excel
// paints a cell's background over its box *plus* the 1pt strip on its bottom
// and right grid boundaries, so neighbouring shadings overlap by exactly the
// strip a border then covers. Typst's own cell `fill:` stops on the boundary,
// so completed background rectangles are extended in the original fill layer.
// ---------------------------------------------------------------------------

/// The `#D9D9D9` band `TableStyleLight1` prints, as measured on a native
/// Excel-for-Mac export of `tests/fixtures/xlsx/ExcelTables.xlsx`.
fn banded_cell(text: &str) -> TableCell {
    TableCell {
        background: Some(Color::new(0xd9, 0xd9, 0xd9)),
        ..plain_text_cell(text)
    }
}

/// A filled horizontal merge's visible Excel region reaches through its
/// positive-axis background band. Excel centres the line on that effective
/// region, while an unfilled merge and a left-aligned line keep the nominal
/// track origin. This is the one-factor rule behind the A1:B1 title in the
/// workbook attached to #982 (issue #1493).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_centered_merged_fill_uses_the_excel_background_band_for_its_text_seat() {
    const EXCEL_POSITIVE_AXIS_BACKGROUND_BAND_PT: f64 = 1.0;

    fn text_run(
        alignment: Alignment,
        background: Option<Color>,
    ) -> crate::render::pdf::PlacedTextRun {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    alignment: Some(alignment),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "MERGED TITLE".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            background,
            col_span: 2,
            spill_width: Some(160.0),
            vertical_align: Some(CellVerticalAlign::Center),
            ..TableCell::default()
        };
        let table = boundary_band_table(vec![fixed_row(vec![cell])], vec![80.0, 80.0]);
        let source = generate_typst(&make_doc(vec![make_flow_page(vec![Block::Table(table)])]))
            .expect("merged sheet table should generate")
            .source;
        crate::render::pdf::compiled_text_runs(&source, 0)
            .unwrap_or_else(|error| {
                panic!("merged sheet table failed to compile: {error}\n{source}")
            })
            .into_iter()
            .find(|run| run.text == "MERGED TITLE")
            .unwrap_or_else(|| panic!("missing merged title in:\n{source}"))
    }

    let rose = Some(Color::new(218, 182, 186));
    let centered_without_fill = text_run(Alignment::Center, None);
    let centered_with_fill = text_run(Alignment::Center, rose);
    assert!(
        (centered_with_fill.left_pt
            - centered_without_fill.left_pt
            - EXCEL_POSITIVE_AXIS_BACKGROUND_BAND_PT)
            .abs()
            < 0.001,
        "the centred filled merge must follow Excel's positive-axis background band: \
         unfilled={}, filled={}",
        centered_without_fill.left_pt,
        centered_with_fill.left_pt,
    );
    assert!(
        (centered_with_fill.baseline_pt - centered_without_fill.baseline_pt).abs() < 0.001,
        "the horizontal seat must not move the title baseline"
    );

    let left_without_fill = text_run(Alignment::Left, None);
    let left_with_fill = text_run(Alignment::Left, rose);
    assert!(
        (left_with_fill.left_pt - left_without_fill.left_pt).abs() < 0.001,
        "a left-aligned merge keeps its nominal origin: unfilled={}, filled={}",
        left_without_fill.left_pt,
        left_with_fill.left_pt,
    );
}

/// A wrapped centred merge never reaches the unwrapped case #1493 probed:
/// `compute_spill_width` (`xlsx_cells.rs`) returns `None` once `wrapText` is
/// set, before the `col_span > 1` branch that returns `Some(merged_width)`
/// for an unwrapped merge. Its line already lands on Excel's real position
/// from the ordinary whole-point sheet seat alone, so the background-band
/// extension must not also apply — doing so moved `04_payroll_ko.xlsx`'s
/// `합계` and `08_budget_ko.xlsx`'s `총계` one point right of Excel's own
/// export (issue #1626). Parametrised over two merge widths so the rule is
/// pinned generally, not to the one reported span.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_wrapped_centered_merged_fill_keeps_its_nominal_track_origin() {
    fn text_run(col_span: u32, background: Option<Color>) -> crate::render::pdf::PlacedTextRun {
        let cell = TableCell {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    alignment: Some(Alignment::Center),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "TOTAL".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            background,
            col_span,
            wraps_text: true,
            vertical_align: Some(CellVerticalAlign::Center),
            ..TableCell::default()
        };
        let column_widths = vec![80.0; col_span.max(1) as usize];
        let table = boundary_band_table(vec![fixed_row(vec![cell])], column_widths);
        let source = generate_typst(&make_doc(vec![make_flow_page(vec![Block::Table(table)])]))
            .expect("merged sheet table should generate")
            .source;
        crate::render::pdf::compiled_text_runs(&source, 0)
            .unwrap_or_else(|error| {
                panic!("merged sheet table failed to compile: {error}\n{source}")
            })
            .into_iter()
            .find(|run| run.text == "TOTAL")
            .unwrap_or_else(|| panic!("missing merged title in:\n{source}"))
    }

    let rose = Some(Color::new(218, 182, 186));
    for col_span in [2, 3] {
        let without_fill = text_run(col_span, None);
        let with_fill = text_run(col_span, rose);
        assert!(
            (with_fill.left_pt - without_fill.left_pt).abs() < 0.001,
            "a wrapped centred merge (col_span={col_span}) takes no band seat: \
             unfilled={}, filled={}",
            without_fill.left_pt,
            with_fill.left_pt,
        );
    }
}

/// Native #982 paints each later cell's fill over its predecessor's positive
/// extension. Probe final rectangle colors inside both neighboring cells;
/// inspecting each strip's dimensions alone missed the reversed order (#1599).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn later_cell_fills_own_shared_boundaries_without_losing_outer_extensions() {
    use crate::render::pdf::compiled_paint_sequence;
    use typst::visualize::Color as PaintColor;

    let pale = Color::new(248, 239, 240);
    let rose = Color::new(218, 182, 186);
    let pale_paint = PaintColor::from_u8(248, 239, 240, 255);
    let rose_paint = PaintColor::from_u8(218, 182, 186, 255);
    let mut failures = Vec::new();
    for scale in [1.0, 0.82] {
        for (horizontal, merged) in [(false, false), (true, false), (false, true), (true, true)] {
            let cell = |color| TableCell {
                background: Some(color),
                content: Vec::new(),
                padding: Some(Insets {
                    top: 0.0,
                    bottom: 0.0,
                    left: 0.0,
                    right: 0.0,
                }),
                ..TableCell::default()
            };
            let row = |cells| TableRow {
                cells,
                height: Some(30.0 * scale),
                minimum_height: None,
            };
            let merged_cell = |color| TableCell {
                col_span: 2,
                ..cell(color)
            };
            let (rows, widths) = if horizontal && merged {
                (
                    vec![row(vec![cell(pale), merged_cell(rose)])],
                    vec![40.0 * scale, 20.0 * scale, 20.0 * scale],
                )
            } else if merged {
                (
                    vec![
                        row(vec![merged_cell(pale)]),
                        row(vec![cell(rose), cell(rose)]),
                    ],
                    vec![20.0 * scale; 2],
                )
            } else if horizontal {
                (
                    vec![row(vec![cell(pale), cell(rose)])],
                    vec![40.0 * scale; 2],
                )
            } else {
                (
                    vec![row(vec![cell(pale)]), row(vec![cell(rose)])],
                    vec![40.0 * scale],
                )
            };
            let mut table = boundary_band_table(rows, widths);
            table.print_scale = Some(scale);
            table.seats_bottom_aligned_text_on_descender = true;
            let doc = make_doc(vec![Page::Sheet(SheetPage {
                name: "Fill ownership".into(),
                size: PageSize {
                    width: 300.0,
                    height: 300.0,
                },
                margins: Margins {
                    top: 30.0,
                    bottom: 30.0,
                    left: 30.0,
                    right: 30.0,
                },
                table,
                header: None,
                footer: None,
                charts: Vec::new(),
                images: Vec::new(),
                text_boxes: Vec::new(),
                shapes: Vec::new(),
            })]);
            let output = generate_typst(&doc).unwrap();
            let paints = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
            let first_fill = paints
                .iter()
                .find(|paint| {
                    paint.rectangle_fill.as_ref() == Some(&pale_paint)
                        && paint.bounds.2 - paint.bounds.0 > 20.0
                        && paint.bounds.3 - paint.bounds.1 > 20.0
                })
                .expect("the first cell has a filled interior")
                .bounds;
            // The page clip trims one sheet point off the first cell's top
            // and left (#1605); the grid origin sits that far before it.
            let origin = (first_fill.0 - scale, first_fill.1 - scale);
            let visible_color = |x: f64, y: f64| {
                paints
                    .iter()
                    .rev()
                    .find(|paint| {
                        paint.rectangle_fill.is_some()
                            && paint.bounds.0 < x
                            && x < paint.bounds.2
                            && paint.bounds.1 < y
                            && y < paint.bounds.3
                    })
                    .and_then(|paint| paint.rectangle_fill.clone())
            };
            let x = origin.0;
            let y = origin.1;
            let width = 40.0 * scale;
            let height = 30.0 * scale;
            let probes = if horizontal {
                vec![
                    (
                        "before column boundary",
                        x + width - 0.5 * scale,
                        y + height / 2.0,
                        pale_paint.clone(),
                    ),
                    (
                        "after column boundary",
                        x + width + 0.5 * scale,
                        y + height / 2.0,
                        rose_paint.clone(),
                    ),
                    (
                        "outer right extension",
                        x + 2.0 * width + 0.5 * scale,
                        y + height / 2.0,
                        rose_paint.clone(),
                    ),
                    (
                        "outer bottom extension",
                        x + width + 10.0 * scale,
                        y + height + 0.5 * scale,
                        rose_paint.clone(),
                    ),
                ]
            } else {
                vec![
                    (
                        "before row boundary",
                        x + width / 2.0,
                        y + height - 0.5 * scale,
                        pale_paint.clone(),
                    ),
                    (
                        "after row boundary",
                        x + width / 2.0,
                        y + height + 0.5 * scale,
                        rose_paint.clone(),
                    ),
                    (
                        "outer bottom extension",
                        x + width / 2.0,
                        y + 2.0 * height + 0.5 * scale,
                        rose_paint.clone(),
                    ),
                    (
                        "outer right extension",
                        x + width + 0.5 * scale,
                        y + height + 10.0 * scale,
                        rose_paint.clone(),
                    ),
                ]
            };
            for (label, px, py, expected) in probes {
                let actual = visible_color(px, py);
                if actual.as_ref() != Some(&expected) {
                    failures.push(format!("scale={scale}, horizontal={horizontal}, merged={merged}, {label}: {actual:?} != {expected:?}"));
                }
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// Repainting a later cell's whole background can fix the fill seam while
/// silently erasing a winning border from an earlier merged cell (#1599).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_later_fill_preserves_the_merged_cells_winning_bottom_border() {
    use crate::render::pdf::compiled_paint_sequence;
    use typst::visualize::Color as PaintColor;

    let dark = Color::new(30, 40, 50);
    let dark_paint = PaintColor::from_u8(30, 40, 50, 255);
    let rose = Color::new(218, 182, 186);
    let cell = |color| TableCell {
        background: Some(color),
        content: Vec::new(),
        padding: Some(Insets {
            top: 0.0,
            bottom: 0.0,
            left: 0.0,
            right: 0.0,
        }),
        ..TableCell::default()
    };
    let upper = TableCell {
        col_span: 2,
        border: Some(CellBorder {
            bottom: Some(BorderSide {
                color: dark,
                ..solid_side(3.0)
            }),
            ..CellBorder::default()
        }),
        ..cell(Color::new(248, 239, 240))
    };
    let table = boundary_band_table(
        vec![
            fixed_row(vec![upper]),
            fixed_row(vec![cell(rose), cell(rose)]),
        ],
        vec![40.0, 40.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let output = generate_typst(&doc).unwrap();
    let paints = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
    let border = paints
        .iter()
        .find(|paint| {
            paint
                .stroke
                .as_ref()
                .is_some_and(|stroke| stroke.color.as_ref() == Some(&dark_paint))
        })
        .expect("the declared 3pt border paints");
    let point_y = (border.bounds.1 + border.bounds.3) / 2.0;
    for fraction in [0.25, 0.75] {
        let point_x = border.bounds.0 + fraction * (border.bounds.2 - border.bounds.0);
        let color = paints.iter().rev().find_map(|paint| {
            let (x0, y0, x1, y1) = paint.bounds;
            if let Some(stroke) = &paint.stroke {
                let half = stroke.thickness_pt / 2.0;
                if point_x > x0 - half
                    && point_x < x1 + half
                    && point_y > y0 - half
                    && point_y < y1 + half
                {
                    return stroke.color.clone();
                }
            }
            if point_x > x0 && point_x < x1 && point_y > y0 && point_y < y1 {
                return paint.rectangle_fill.clone();
            }
            None
        });
        assert_eq!(
            color,
            Some(dark_paint.clone()),
            "both lower cells preserve the resolved band"
        );
    }
}

/// The cell reached later in row order owns the corner overlap. An unfilled
/// lower neighbor leaves the upper-right fill visible instead (#1495, #1599).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn row_order_determines_the_visible_fill_at_shared_corners() {
    use crate::render::pdf::compiled_paint_sequence;
    use typst::visualize::Color as PaintColor;
    let pale = Color::new(248, 239, 240);
    let rose = Color::new(218, 182, 186);
    let green = Color::new(40, 120, 60);
    for lower_filled in [false, true] {
        let cell = |background| TableCell {
            background,
            content: Vec::new(),
            ..TableCell::default()
        };
        let table = boundary_band_table(
            vec![
                fixed_row(vec![cell(Some(pale)), cell(Some(rose))]),
                fixed_row(vec![cell(lower_filled.then_some(green)), cell(None)]),
            ],
            vec![40.0, 40.0],
        );
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let output = generate_typst(&doc).unwrap();
        let paints = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
        let first = paints
            .iter()
            .find(|p| p.rectangle_fill == Some(PaintColor::from_u8(248, 239, 240, 255)))
            .unwrap();
        // The first fill starts one sheet point inside the grid origin (#1605).
        let (x, y) = (first.bounds.0 - 1.0 + 40.5, first.bounds.1 - 1.0 + 20.5);
        let visible = paints
            .iter()
            .rev()
            .find(|p| {
                p.rectangle_fill.is_some()
                    && p.bounds.0 < x
                    && x < p.bounds.2
                    && p.bounds.1 < y
                    && y < p.bounds.3
            })
            .and_then(|p| p.rectangle_fill.clone());
        let expected = if lower_filled {
            PaintColor::from_u8(40, 120, 60, 255)
        } else {
            PaintColor::from_u8(218, 182, 186, 255)
        };
        assert_eq!(visible, Some(expected), "lower filled={lower_filled}");
    }
}

/// A long automatic row needs continuous outer fill coverage; painting two
/// estimated strips leaves a middle gap. Alpha must also apply only once at
/// the inner seam, rather than darkening an overlap (#1190, #1397, #1599).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn automatic_row_fill_has_continuous_single_alpha_coverage() {
    use crate::render::pdf::compiled_paint_sequence;
    let cell = TableCell {
        background: Some(Color::new(255, 0, 0)),
        background_alpha: Some(0.5),
        ..plain_text_cell(
            "A long automatic row wraps these words over many lines to exceed twice the estimated single line height.",
        )
    };
    let table = boundary_band_table(
        vec![TableRow {
            cells: vec![cell],
            height: None,
            minimum_height: None,
        }],
        vec![80.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let output = generate_typst(&doc).unwrap();
    let paints = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
    let main = paints
        .iter()
        .find(|p| {
            p.rectangle_fill.is_some()
                && p.bounds.2 - p.bounds.0 > 70.0
                && p.bounds.3 - p.bounds.1 > 40.0
        })
        .expect("the automatic row wraps to several lines");
    let y = (main.bounds.1 + main.bounds.3) / 2.0;
    // The page clip trims one sheet point off the first cell's left (#1605),
    // so the grid origin sits one point before the painted fill.
    let grid_left: f64 = main.bounds.0 - 1.0;
    for dx in [79.9, 80.5] {
        let x = grid_left + dx;
        let mut rgb = [1.0_f32; 3];
        for paint in &paints {
            if paint.bounds.0 < x
                && x < paint.bounds.2
                && paint.bounds.1 < y
                && y < paint.bounds.3
                && let Some(fill) = &paint.rectangle_fill
            {
                let color = fill.to_rgb();
                for (out, channel) in rgb.iter_mut().zip([color.red, color.green, color.blue]) {
                    *out = channel * color.alpha + *out * (1.0 - color.alpha);
                }
            }
        }
        let expected = [1.0, 127.0 / 255.0, 127.0 / 255.0];
        for (actual, expected) in rgb.into_iter().zip(expected) {
            assert!((actual - expected).abs() < 0.00001, "dx={dx}: {rgb:?}");
        }
    }
}

/// An Excel table beside ordinary Word/PowerPoint tables must not change
/// their fill extents, even when all three tables use the same paint.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_word_and_centred_stroke_tables_do_not_bleed_their_fills() {
    use crate::render::pdf::compiled_paint_sequence;
    use typst::visualize::Color as PaintColor;
    let models = [
        TableBorderPaintModel::ExcelBoundaryBands,
        TableBorderPaintModel::WordPositiveAxisBands,
        TableBorderPaintModel::CenteredStroke,
    ];
    let blocks = models
        .into_iter()
        .map(|model| {
            Block::Table(Table {
                rows: vec![fixed_row(vec![banded_cell("Anton")])],
                column_widths: vec![69.0],
                border_paint_model: model,
                ..Table::default()
            })
        })
        .collect();
    let doc = make_doc(vec![make_flow_page(blocks)]);
    let output = generate_typst(&doc).unwrap();
    let paints = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
    let fills: Vec<_> = paints
        .iter()
        .filter(|p| p.rectangle_fill == Some(PaintColor::from_u8(217, 217, 217, 255)))
        .collect();
    assert_eq!(fills.len(), 3);
    // The Word and centred-stroke fills cover exactly their 69x20pt cell.
    for fill in &fills[1..] {
        assert!((fill.bounds.2 - fill.bounds.0 - 69.0).abs() < 0.001);
        assert!((fill.bounds.3 - fill.bounds.1 - 20.0).abs() < 0.001);
        assert!((fill.bounds.0 - fills[1].bounds.0).abs() < 0.001);
    }
    // The Excel fill bleeds one point past its right and bottom boundaries
    // and loses one point to the page clip on its left and top (#1605), so
    // its box keeps the nominal size but starts one point inside the grid
    // origin the other two tables share.
    let excel = fills[0];
    assert!((excel.bounds.0 - fills[1].bounds.0 - 1.0).abs() < 0.001);
    assert!((excel.bounds.2 - fills[1].bounds.0 - 70.0).abs() < 0.001);
    assert!((excel.bounds.3 - excel.bounds.1 - 20.0).abs() < 0.001);
}

/// The native Excel-for-Mac export places the first TableStyleLight1 body
/// cell at (464, 73), extending its 69pt by 17pt track to (534, 91).
/// Check the compiled fill itself so this real-fixture regression for #1190
/// survives changes in how background coverage is generated (#1599).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn structure_light1_table_band_bleeds_past_its_bottom_and_right_boundaries() {
    use crate::parser::Parser;
    use crate::parser::xlsx::XlsxParser;
    use crate::render::pdf::compiled_paint_sequence;
    use typst::visualize::Color as PaintColor;

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx/ExcelTables.xlsx");
    let data = std::fs::read(path).expect("fixture should be available");
    let (document, _) = XlsxParser
        .parse(&data, &crate::ConvertOptions::default())
        .expect("fixture should parse");
    let output = generate_typst(&document).expect("fixture should generate Typst");
    let paints =
        compiled_paint_sequence(&output.source, &output.images, 0).expect("fixture should compile");
    let fills: Vec<_> = paints
        .iter()
        .filter(|paint| paint.rectangle_fill == Some(PaintColor::from_u8(217, 217, 217, 255)))
        .map(|paint| paint.bounds)
        .collect();
    assert!(
        fills.iter().any(|&(left, top, right, bottom)| {
            (left - 464.0).abs() < 0.001
                && (top - 73.0).abs() < 0.001
                && (right - 534.0).abs() < 0.001
                && (bottom - 91.0).abs() < 0.001
        }),
        "the first body cell must fill through the bottom and right boundaries: {fills:?}"
    );
}

// ---------------------------------------------------------------------------
// Excel page fill clip (issue #1605)
//
// Native Excel clips every page's cell fills to that page's grid region inset
// by one sheet point on the top and left edges, while the bottom and right
// edges keep the positive-axis bleed. The public #982 workbook traces the
// unscaled A1:B8 sheet's clip as `[51, 55, 474, 418]` around raw fills that
// start at x=50/y=54; the fitted 0.82 explicit-area sheet clips at
// `[55.76, 54.12]` for a grid origin of `[54.94, 53.30]`; and the unscaled
// explicit-area sheet's continuation page clips at x=473 for a raw fill that
// starts at x=472. Later rows and columns keep their raw origin, so only the
// page's first row loses its top strip and only its first column loses its
// left strip.
// ---------------------------------------------------------------------------

/// A 2x2 filled sheet table on one page and the same table on each of two
/// pages, as the XLSX parser emits a continuation page, each compiled at the
/// unscaled and the fitted print scale.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn page_fills_lose_one_sheet_point_on_the_top_and_left_grid_edges() {
    use crate::render::pdf::compiled_paint_sequence;
    use typst::visualize::Color as PaintColor;

    let pale = Color::new(248, 239, 240);
    let rose = Color::new(218, 182, 186);
    let pale_paint = PaintColor::from_u8(248, 239, 240, 255);
    let rose_paint = PaintColor::from_u8(218, 182, 186, 255);
    let mut failures: Vec<String> = Vec::new();
    for scale in [1.0, 0.82] {
        for page_count in [1usize, 2] {
            let cell = |color: Color| TableCell {
                background: Some(color),
                content: Vec::new(),
                padding: Some(Insets {
                    top: 0.0,
                    bottom: 0.0,
                    left: 0.0,
                    right: 0.0,
                }),
                ..TableCell::default()
            };
            let width: f64 = 40.0 * scale;
            let height: f64 = 30.0 * scale;
            let sheet_page = || {
                let rows: Vec<TableRow> = (0..2)
                    .map(|index| TableRow {
                        cells: vec![cell(if index == 0 { pale } else { rose }), cell(rose)],
                        height: Some(height),
                        minimum_height: None,
                    })
                    .collect();
                let mut table = boundary_band_table(rows, vec![width; 2]);
                table.print_scale = Some(scale);
                table.seats_bottom_aligned_text_on_descender = true;
                Page::Sheet(SheetPage {
                    name: "Fill clip".into(),
                    size: PageSize {
                        width: 300.0,
                        height: 300.0,
                    },
                    margins: Margins {
                        top: 30.0,
                        bottom: 30.0,
                        left: 30.0,
                        right: 30.0,
                    },
                    table,
                    header: None,
                    footer: None,
                    charts: Vec::new(),
                    images: Vec::new(),
                    text_boxes: Vec::new(),
                    shapes: Vec::new(),
                })
            };
            let doc = make_doc((0..page_count).map(|_| sheet_page()).collect());
            let output = generate_typst(&doc).unwrap();
            for page_index in 0..page_count {
                let paints = compiled_paint_sequence(&output.source, &output.images, page_index)
                    .unwrap_or_else(|error| panic!("page {page_index} failed to compile: {error}"));
                let fills: Vec<(f64, f64, f64, f64)> = paints
                    .iter()
                    .filter(|paint| {
                        let fill = paint.rectangle_fill.as_ref();
                        fill == Some(&pale_paint) || fill == Some(&rose_paint)
                    })
                    .map(|paint| paint.bounds)
                    .collect();
                let label = format!("scale={scale}, pages={page_count}, page={page_index}");
                if fills.len() < 4 {
                    failures.push(format!(
                        "{label}: expected at least four fills, got {fills:?}"
                    ));
                    continue;
                }
                // The second column's fills keep the raw grid origin: their
                // left edge is the shared boundary, one column past the grid.
                let second_column_left: f64 = fills
                    .iter()
                    .map(|fill| fill.0)
                    .fold(f64::NEG_INFINITY, f64::max);
                let second_row_top: f64 = fills
                    .iter()
                    .filter(|fill| fill.0 < second_column_left - 1.0)
                    .map(|fill| fill.1)
                    .filter(|top| {
                        *top > fills
                            .iter()
                            .map(|fill| fill.1)
                            .fold(f64::INFINITY, f64::min)
                            + 1.0
                    })
                    .fold(f64::INFINITY, f64::min);
                let grid_left: f64 = second_column_left - width;
                let grid_top: f64 = second_row_top - height;
                let first_left: f64 = fills
                    .iter()
                    .map(|fill| fill.0)
                    .fold(f64::INFINITY, f64::min);
                let first_top: f64 = fills
                    .iter()
                    .map(|fill| fill.1)
                    .fold(f64::INFINITY, f64::min);
                if (first_left - (grid_left + scale)).abs() > 0.01 {
                    failures.push(format!(
                        "{label}: first column fills start at x={first_left:.3}, expected the grid's \
                         {grid_left:.3} plus the {scale} sheet point clip"
                    ));
                }
                if (first_top - (grid_top + scale)).abs() > 0.01 {
                    failures.push(format!(
                        "{label}: first row fills start at y={first_top:.3}, expected the grid's \
                         {grid_top:.3} plus the {scale} sheet point clip"
                    ));
                }
                // The positive-axis bleed past the last column and row stays.
                let last_right: f64 = fills
                    .iter()
                    .map(|fill| fill.2)
                    .fold(f64::NEG_INFINITY, f64::max);
                if (last_right - (grid_left + 2.0 * width + scale)).abs() > 0.01 {
                    failures.push(format!(
                        "{label}: the right bleed ends at x={last_right:.3}, expected {:.3}",
                        grid_left + 2.0 * width + scale
                    ));
                }
                let last_row_bottom: f64 = fills
                    .iter()
                    .map(|fill| fill.3)
                    .fold(f64::NEG_INFINITY, f64::max);
                let last_row_top: f64 = fills
                    .iter()
                    .map(|fill| fill.1)
                    .fold(f64::NEG_INFINITY, f64::max);
                if (last_row_bottom - (last_row_top + height + scale)).abs() > 0.01 {
                    failures.push(format!(
                        "{label}: the bottom bleed ends at y={last_row_bottom:.3}, expected {:.3}",
                        last_row_top + height + scale
                    ));
                }
                // Interior boundaries are untouched: the second column and
                // row fills start exactly one track past the first.
                let interior_lefts: Vec<f64> = fills
                    .iter()
                    .map(|fill| fill.0)
                    .filter(|left| (*left - first_left).abs() > 0.01)
                    .collect();
                if interior_lefts
                    .iter()
                    .any(|left| (*left - second_column_left).abs() > 0.01)
                {
                    failures.push(format!(
                        "{label}: interior column origins moved: {interior_lefts:?}"
                    ));
                }
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}
