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
// so codegen must add that strip as an overlay.
// ---------------------------------------------------------------------------

/// The `#D9D9D9` band `TableStyleLight1` prints, as measured on a native
/// Excel-for-Mac export of `tests/fixtures/xlsx/ExcelTables.xlsx`.
fn banded_cell(text: &str) -> TableCell {
    TableCell {
        background: Some(Color::new(0xd9, 0xd9, 0xd9)),
        ..plain_text_cell(text)
    }
}

#[test]
fn test_boundary_band_background_bleed_overlaps_its_cell_at_the_corner_junction() {
    let table = boundary_band_table(
        vec![fixed_row(vec![banded_cell("Anton"), plain_text_cell("44")])],
        vec![69.0, 69.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // The track itself still comes from Typst's own cell fill.
    assert!(
        result.contains("table.cell(fill: rgb(217, 217, 217))"),
        "the cell must keep filling its own box: {result}"
    );
    // Bottom strip: its outer edge still lands 1pt past the row boundary, but
    // its inner edge overlaps the cell fill by 0.25pt. Three same-colour paths
    // that only meet at the boundary corner leave a one-pixel pinhole when the
    // PDF is rasterised at 150 DPI (issue #1397).
    assert!(
        result.contains(
            "#place(bottom + left, dx: -5pt, dy: 6pt, rect(width: 100% + 11pt, height: 1.25pt, fill: rgb(217, 217, 217), stroke: none))"
        ),
        "the background must overlap the cell before bleeding 1pt past the bottom boundary: {result}"
    );
    // The right strip follows the same inward overlap while preserving its
    // measured outer edge and the 20pt row frame plus corner block.
    assert!(
        result.contains(
            "#place(top + right, dx: 6pt, dy: -5pt, rect(width: 1.25pt, height: 21pt, fill: rgb(217, 217, 217), stroke: none))"
        ),
        "the background must overlap the cell before bleeding 1pt past the right boundary: {result}"
    );
    // An unshaded neighbour paints nothing.
    assert_eq!(
        result.matches("rect(width: 1.25pt").count(),
        1,
        "only the shaded cell may bleed: {result}"
    );
}

/// Excel applies a fit-to-page transform after painting its sheet-space 1pt
/// positive-axis background band. On the issue #1538 workbook's 0.82-scale
/// page, the outer bleed is therefore 0.82pt and the 0.25pt seam overlap is
/// 0.205pt, not the unscaled 1pt and 0.25pt emitted before this regression.
#[test]
fn a_fit_scaled_sheet_scales_its_excel_background_bleed() {
    let mut table = boundary_band_table(vec![fixed_row(vec![banded_cell("Gift")])], vec![69.0]);
    table.print_scale = Some(0.82);
    table.seats_bottom_aligned_text_on_descender = true;
    let doc = make_doc(vec![Page::Sheet(SheetPage {
        name: "Gift budget and tracker".to_string(),
        size: PageSize {
            width: 1_190.55,
            height: 841.89,
        },
        margins: Margins {
            top: 54.0,
            bottom: 54.0,
            left: 50.0,
            right: 50.0,
        },
        table,
        header: None,
        footer: None,
        charts: Vec::new(),
        images: Vec::new(),
        text_boxes: Vec::new(),
    })]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("rect(width: 1.025pt, height: 20.82pt"),
        "the right bleed must scale its band, overlap, and row-end extension: {result}"
    );
    assert!(
        result.contains("height: 1.025pt"),
        "the bottom bleed must scale its band and overlap: {result}"
    );
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

/// `Gift Budget and Tracker1.xlsx` (attached to #982) puts a rose title fill
/// above pale body cells. The title's bottom bleed forms the visible rule, so
/// the later cells' right-edge bleed must start below that positive-axis band;
/// otherwise its later paint order cuts the two pale notches in issue #1475.
#[test]
fn a_later_cell_fill_does_not_cut_through_the_rule_above_it() {
    let rose = Color::new(218, 182, 186);
    let pale_left = Color::new(248, 239, 240);
    let pale_right = Color::new(247, 238, 239);
    let header = TableCell {
        background: Some(rose),
        col_span: 2,
        ..plain_text_cell("Title")
    };
    let body_left = TableCell {
        background: Some(pale_left),
        ..plain_text_cell("Left")
    };
    let body_right = TableCell {
        background: Some(pale_right),
        ..plain_text_cell("Right")
    };
    let table = boundary_band_table(
        vec![
            fixed_row(vec![header]),
            fixed_row(vec![body_left, body_right]),
        ],
        vec![69.0, 20.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(
            "#place(top + right, dx: 6pt, dy: -4pt, rect(width: 1.25pt, height: 20pt, fill: rgb(248, 239, 240), stroke: none))"
        ),
        "the internal junction beneath the merge must preserve the title fill's 1pt bottom band: {result}"
    );
    assert!(
        result.contains(
            "#place(top + right, dx: 6pt, dy: -4pt, rect(width: 1.25pt, height: 20pt, fill: rgb(247, 238, 239), stroke: none))"
        ),
        "the exterior junction beneath the merge must preserve the title fill's 1pt bottom band: {result}"
    );
    assert!(
        !result.contains(
            "#place(top + right, dx: 6pt, dy: -5pt, rect(width: 1.25pt, height: 21pt, fill: rgb(248, 239, 240), stroke: none))"
        ),
        "the later fill must not begin on top of the preceding rule: {result}"
    );
}

/// The second sheet in `Gift Budget and Tracker1.xlsx` has pale E cells beside
/// a rose F:Q merge. Native Excel steps the shared corner: the lower-left
/// cell's right bleed begins below the upper row, while the upper-right cell's
/// bottom bleed begins to the right of the pale corner block. Letting either
/// strip cross the other cuts a differently coloured 1pt notch (#1495).
#[test]
fn differently_colored_adjacent_fills_step_their_shared_corner() {
    let pale = Color::new(248, 239, 240);
    let rose = Color::new(218, 182, 186);
    let upper_left = TableCell {
        background: Some(pale),
        ..plain_text_cell("E2")
    };
    let upper_right = TableCell {
        background: Some(rose),
        col_span: 2,
        ..plain_text_cell("F2:Q2")
    };
    let lower_left = TableCell {
        background: Some(pale),
        ..plain_text_cell("E3")
    };
    let table = boundary_band_table(
        vec![
            fixed_row(vec![upper_left, upper_right]),
            fixed_row(vec![
                lower_left,
                plain_text_cell("F3"),
                plain_text_cell("G3"),
            ]),
        ],
        vec![20.0, 69.0, 69.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(
            "#place(top + right, dx: 6pt, dy: -4pt, rect(width: 1.25pt, height: 20pt, fill: rgb(248, 239, 240), stroke: none))"
        ),
        "the lower-left bleed must start below the upper-right cell's bottom band: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + left, dx: -4pt, dy: 6pt, rect(width: 100% + 10pt, height: 1.25pt, fill: rgb(218, 182, 186), stroke: none))"
        ),
        "the upper-right bottom bleed must begin beyond the pale corner block: {result}"
    );
}

/// With no filled cell starting on the upper-right side of a junction, the
/// upper cell's bottom band remains the crossing owner. This is the G-column
/// pattern on the fixture's second sheet; failing to trim produces four
/// one-point notches between its differently coloured conditional fills.
#[test]
fn an_upper_fill_keeps_the_crossing_when_no_upper_right_fill_replaces_it() {
    let dark = Color::new(61, 43, 45);
    let purple = Color::new(93, 55, 84);
    let upper_left = TableCell {
        background: Some(dark),
        ..plain_text_cell("G7")
    };
    let lower_left = TableCell {
        background: Some(purple),
        ..plain_text_cell("G8")
    };
    let table = boundary_band_table(
        vec![
            fixed_row(vec![upper_left, plain_text_cell("H7")]),
            fixed_row(vec![lower_left, plain_text_cell("H8")]),
        ],
        vec![69.0, 69.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains(
            "#place(top + right, dx: 6pt, dy: -4pt, rect(width: 1.25pt, height: 20pt, fill: rgb(93, 55, 84), stroke: none))"
        ),
        "the lower vertical bleed must start below the upper cell's bottom band: {result}"
    );
}

/// The bleed's vertical run obeys the same extent rule the vertical border
/// bands do: a relative height inside `#place` resolves against the page in an
/// auto-sized row, so the strip is painted as concrete twins instead.
#[test]
fn test_boundary_band_background_bleed_paints_twins_in_an_auto_row() {
    let table = boundary_band_table(
        vec![TableRow {
            minimum_height: None,
            cells: vec![banded_cell("Anton")],
            height: None,
        }],
        vec![69.0],
    );
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    // `plain_text_cell` declares no font family, so no line metrics exist and
    // the extent falls back to the ambient text size.
    assert!(
        result.contains(
            "#place(top + right, dx: 6pt, dy: -5pt, rect(width: 1.25pt, height: 1.2em + 11pt, fill: rgb(217, 217, 217), stroke: none))"
        ),
        "the top twin must hang from the row's top boundary: {result}"
    );
    assert!(
        result.contains(
            "#place(bottom + right, dx: 6pt, dy: 6pt, rect(width: 1.25pt, height: 1.2em + 11pt, fill: rgb(217, 217, 217), stroke: none))"
        ),
        "the bottom twin must rise from 1pt past the bottom boundary: {result}"
    );
}

/// The bleed belongs to Excel's convention alone: Word's own band model
/// (issue #724) was measured on borders only, and PowerPoint and Word tables
/// keep Typst's centred strokes.
#[test]
fn test_word_and_centred_stroke_tables_do_not_bleed_their_fills() {
    for model in [
        TableBorderPaintModel::WordPositiveAxisBands,
        TableBorderPaintModel::CenteredStroke,
    ] {
        let table = Table {
            rows: vec![fixed_row(vec![banded_cell("Anton")])],
            column_widths: vec![69.0],
            border_paint_model: model,
            ..Table::default()
        };
        let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
        let result = generate_typst(&doc).unwrap().source;
        assert!(
            !result.contains("rect(width: 1pt"),
            "{model:?} must keep filling the cell box exactly: {result}"
        );
    }
}
