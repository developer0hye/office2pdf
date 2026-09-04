use super::*;

// ── Border direction tracking ───────────────────────────────────────

#[derive(Clone, Copy, PartialEq, Eq)]
enum BorderDir {
    None,
    Left,
    Right,
    Top,
    Bottom,
}

// ── Table parser state machine ──────────────────────────────────────

/// Bundles all mutable state needed to parse a `<a:tbl>` element into
/// a [`Table`] IR node. Each XML event is dispatched to a method on
/// this struct, keeping the top-level loop minimal.
struct PptxTableParser<'a> {
    // External context (immutable references)
    theme: &'a ThemeData,
    color_map: &'a ColorMapData,
    table_styles: &'a table_styles::TableStyleMap,

    // ── Table-level state ───────────────────────────────────────────
    column_widths: Vec<f64>,
    rows: Vec<TableRow>,
    table_props: table_styles::PptxTableProps,
    is_in_tbl_pr: bool,
    is_in_table_style_id: bool,

    // ── Row-level state ─────────────────────────────────────────────
    is_in_row: bool,
    row_height_emu: i64,
    cells: Vec<TableCell>,

    // ── Cell-level state ────────────────────────────────────────────
    is_in_cell: bool,
    cell_col_span: u32,
    cell_row_span: u32,
    is_horizontal_merge: bool,
    is_vertical_merge: bool,
    cell_text_entries: Vec<PptxParagraphEntry>,
    cell_background: Option<Color>,
    cell_background_alpha: Option<f64>,
    cell_fill_suppressed: bool,
    row_no_fill_cells: Vec<bool>,
    cell_suppressed_borders: table_styles::SuppressedCellBorders,
    row_suppressed_cell_borders: Vec<table_styles::SuppressedCellBorders>,
    cell_vertical_align: Option<CellVerticalAlign>,
    cell_padding: Option<Insets>,

    // ── Text body state ─────────────────────────────────────────────
    is_in_text_body: bool,
    text_body_style_defaults: PptxTextBodyStyleDefaults,

    // ── Paragraph-level state ───────────────────────────────────────
    is_in_paragraph: bool,
    paragraph_style: ParagraphStyle,
    paragraph_level: u32,
    /// `p:defaultTextStyle` size, applied to any run declaring none.
    default_text_size_pt: Option<f64>,
    paragraph_default_run_style: TextStyle,
    paragraph_end_run_style: TextStyle,
    paragraph_bullet_definition: PptxBulletDefinition,
    is_in_line_spacing: bool,
    runs: Vec<Run>,

    // ── Run-level state ─────────────────────────────────────────────
    is_in_run: bool,
    run_style: TextStyle,
    run_text: String,
    run_has_explicit_underline: bool,
    run_marker_style_before_hyperlink: Option<TextStyle>,
    first_run_marker_style_override: Option<TextStyle>,
    is_in_text: bool,
    is_in_run_properties: bool,
    is_in_end_paragraph_run_properties: bool,

    // ── Fill context ────────────────────────────────────────────────
    solid_fill_context: SolidFillCtx,

    // ── Cell property / border state ────────────────────────────────
    is_in_table_cell_properties: bool,
    border_left: Option<BorderSide>,
    border_right: Option<BorderSide>,
    border_top: Option<BorderSide>,
    border_bottom: Option<BorderSide>,
    is_in_border_line: bool,
    border_line_width_emu: i64,
    border_line_color: Option<Color>,
    border_line_dash_style: BorderLineStyle,
    current_border_dir: BorderDir,
}

impl<'a> PptxTableParser<'a> {
    fn new(
        theme: &'a ThemeData,
        color_map: &'a ColorMapData,
        table_styles: &'a table_styles::TableStyleMap,
        default_text_size_pt: Option<f64>,
    ) -> Self {
        Self {
            default_text_size_pt,
            theme,
            color_map,
            table_styles,

            column_widths: Vec::new(),
            rows: Vec::new(),
            table_props: table_styles::PptxTableProps::default(),
            is_in_tbl_pr: false,
            is_in_table_style_id: false,

            is_in_row: false,
            row_height_emu: 0,
            cells: Vec::new(),

            is_in_cell: false,
            cell_col_span: 1,
            cell_row_span: 1,
            is_horizontal_merge: false,
            is_vertical_merge: false,
            cell_text_entries: Vec::new(),
            cell_background: None,
            cell_background_alpha: None,
            cell_fill_suppressed: false,
            row_no_fill_cells: Vec::new(),
            cell_suppressed_borders: table_styles::SuppressedCellBorders::default(),
            row_suppressed_cell_borders: Vec::new(),
            cell_vertical_align: None,
            cell_padding: None,

            is_in_text_body: false,
            text_body_style_defaults: PptxTextBodyStyleDefaults::default(),

            is_in_paragraph: false,
            paragraph_style: ParagraphStyle::default(),
            paragraph_level: 0,
            paragraph_default_run_style: TextStyle::default(),
            paragraph_end_run_style: TextStyle::default(),
            paragraph_bullet_definition: PptxBulletDefinition::default(),
            is_in_line_spacing: false,
            runs: Vec::new(),

            is_in_run: false,
            run_style: TextStyle::default(),
            run_text: String::new(),
            run_has_explicit_underline: false,
            run_marker_style_before_hyperlink: None,
            first_run_marker_style_override: None,
            is_in_text: false,
            is_in_run_properties: false,
            is_in_end_paragraph_run_properties: false,

            solid_fill_context: SolidFillCtx::None,

            is_in_table_cell_properties: false,
            border_left: None,
            border_right: None,
            border_top: None,
            border_bottom: None,
            is_in_border_line: false,
            border_line_width_emu: 0,
            border_line_color: None,
            border_line_dash_style: BorderLineStyle::Solid,
            current_border_dir: BorderDir::None,
        }
    }

    // ── Start element dispatch ──────────────────────────────────────

    fn handle_start(
        &mut self,
        reader: &mut Reader<&[u8]>,
        e: &BytesStart,
    ) -> Result<(), ConvertError> {
        let local = e.local_name();
        match local.as_ref() {
            b"tblPr" => {
                self.is_in_tbl_pr = true;
                self.parse_tbl_pr_attrs(e);
            }
            b"tableStyleId" if self.is_in_tbl_pr => {
                self.is_in_table_style_id = true;
            }
            b"gridCol" => {
                if let Some(width) = get_attr_i64(e, b"w") {
                    self.column_widths.push(emu_to_pt(width));
                }
            }
            b"tr" => {
                self.is_in_row = true;
                self.row_height_emu = get_attr_i64(e, b"h").unwrap_or(0);
                self.cells.clear();
            }
            b"tc" if self.is_in_row => {
                self.enter_cell(e);
            }
            b"txBody" if self.is_in_cell => {
                self.is_in_text_body = true;
                self.text_body_style_defaults = PptxTextBodyStyleDefaults::default();
            }
            b"lstStyle" if self.is_in_text_body => {
                let local_defaults = parse_pptx_list_style(reader, self.theme, self.color_map);
                self.text_body_style_defaults.merge_from(&local_defaults);
            }
            b"p" if self.is_in_text_body => {
                self.enter_paragraph();
            }
            b"pPr" if self.is_in_paragraph && !self.is_in_run => {
                self.handle_paragraph_properties(e);
            }
            b"lnSpc" if self.is_in_paragraph && !self.is_in_run => {
                self.is_in_line_spacing = true;
            }
            b"spcPct" if self.is_in_line_spacing => {
                extract_pptx_line_spacing_pct(e, &mut self.paragraph_style);
            }
            b"spcPts" if self.is_in_line_spacing => {
                extract_pptx_line_spacing_pts(e, &mut self.paragraph_style);
            }
            name if self.is_in_paragraph && !self.is_in_run => {
                if !self.dispatch_bullet_element(name, e) {
                    self.handle_start_non_bullet(reader, name, e)?;
                }
            }
            _ if self.is_in_paragraph => {
                self.handle_start_run_and_fill(reader, local.as_ref(), e)?;
            }
            b"tcPr" if self.is_in_cell => {
                self.is_in_table_cell_properties = true;
                extract_pptx_table_cell_props(
                    e,
                    &mut self.cell_vertical_align,
                    &mut self.cell_padding,
                );
            }
            b"lnL" if self.is_in_table_cell_properties => {
                self.enter_border_line(BorderDir::Left, e);
            }
            b"lnR" if self.is_in_table_cell_properties => {
                self.enter_border_line(BorderDir::Right, e);
            }
            b"lnT" if self.is_in_table_cell_properties => {
                self.enter_border_line(BorderDir::Top, e);
            }
            b"lnB" if self.is_in_table_cell_properties => {
                self.enter_border_line(BorderDir::Bottom, e);
            }
            b"prstDash" if self.is_in_border_line => {
                self.border_line_dash_style = get_attr_str(e, b"val")
                    .as_deref()
                    .map(pptx_dash_to_border_style)
                    .unwrap_or(BorderLineStyle::Solid);
            }
            b"solidFill" if self.is_in_table_cell_properties && !self.is_in_border_line => {
                self.solid_fill_context = SolidFillCtx::ShapeFill;
            }
            b"solidFill" if self.is_in_border_line => {
                self.solid_fill_context = SolidFillCtx::LineFill;
            }
            b"noFill" if self.is_in_border_line => {
                self.suppress_current_border();
            }
            b"noFill" if self.is_in_table_cell_properties && !self.is_in_border_line => {
                self.cell_background = None;
                self.cell_background_alpha = None;
                self.cell_fill_suppressed = true;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr"
                if self.solid_fill_context != SolidFillCtx::None =>
            {
                self.apply_color_start(reader, e);
            }
            _ => {}
        }
        Ok(())
    }

    // ── Empty element dispatch ──────────────────────────────────────

    fn handle_empty(&mut self, e: &BytesStart) {
        let local = e.local_name();
        match local.as_ref() {
            b"tblPr" => {
                self.parse_tbl_pr_attrs(e);
            }
            b"gridCol" => {
                if let Some(width) = get_attr_i64(e, b"w") {
                    self.column_widths.push(emu_to_pt(width));
                }
            }
            b"srgbClr" | b"schemeClr" | b"sysClr"
                if self.solid_fill_context != SolidFillCtx::None =>
            {
                self.apply_color_empty(e);
            }
            b"prstDash" if self.is_in_border_line => {
                self.border_line_dash_style = get_attr_str(e, b"val")
                    .as_deref()
                    .map(pptx_dash_to_border_style)
                    .unwrap_or(BorderLineStyle::Solid);
            }
            b"rPr" if self.is_in_run => {
                self.run_has_explicit_underline = get_attr_str(e, b"u").is_some();
                extract_rpr_attributes(e, &mut self.run_style);
            }
            b"endParaRPr" if self.is_in_paragraph && !self.is_in_run => {
                self.paragraph_end_run_style = self.paragraph_default_run_style.clone();
                extract_rpr_attributes(e, &mut self.paragraph_end_run_style);
            }
            b"tcPr" if self.is_in_cell => {
                extract_pptx_table_cell_props(
                    e,
                    &mut self.cell_vertical_align,
                    &mut self.cell_padding,
                );
            }
            b"noFill" if self.is_in_table_cell_properties && !self.is_in_border_line => {
                self.cell_background = None;
                self.cell_background_alpha = None;
                self.cell_fill_suppressed = true;
            }
            b"noFill" if self.is_in_border_line => {
                self.suppress_current_border();
            }
            b"pPr" if self.is_in_paragraph && !self.is_in_run => {
                self.handle_paragraph_properties(e);
            }
            b"lnSpc" if self.is_in_paragraph && !self.is_in_run => {
                self.is_in_line_spacing = true;
            }
            b"spcPct" if self.is_in_line_spacing => {
                extract_pptx_line_spacing_pct(e, &mut self.paragraph_style);
            }
            b"spcPts" if self.is_in_line_spacing => {
                extract_pptx_line_spacing_pts(e, &mut self.paragraph_style);
            }
            name if self.is_in_paragraph && !self.is_in_run => {
                if !self.dispatch_bullet_element(name, e) {
                    self.handle_empty_non_bullet(name, e);
                }
            }
            b"latin" | b"ea" | b"cs" if self.is_in_run_properties => {
                apply_typeface_to_style(e, &mut self.run_style, self.theme);
            }
            b"latin" | b"ea" | b"cs" if self.is_in_end_paragraph_run_properties => {
                apply_typeface_to_style(e, &mut self.paragraph_end_run_style, self.theme);
            }
            b"hlinkClick" if self.is_in_run_properties => {
                if self.run_marker_style_before_hyperlink.is_none() {
                    self.run_marker_style_before_hyperlink = Some(self.run_style.clone());
                }
                apply_pptx_hyperlink_style(
                    &mut self.run_style,
                    self.run_has_explicit_underline,
                    self.theme,
                    self.color_map,
                );
            }
            _ => {}
        }
    }

    // ── Text / GeneralRef events ────────────────────────────────────

    fn handle_text(&mut self, text: &quick_xml::events::BytesText<'_>) {
        if self.is_in_table_style_id {
            if let Some(decoded) = decode_pptx_text_event(text) {
                self.table_props.style_id = Some(decoded);
            }
        } else if self.is_in_text
            && let Some(decoded) = decode_pptx_text_event(text)
        {
            self.run_text.push_str(&decoded);
        }
    }

    fn handle_general_ref(&mut self, reference: &quick_xml::events::BytesRef<'_>) {
        if self.is_in_text
            && let Some(decoded) = crate::parser::xml_util::decode_general_ref(reference)
        {
            self.run_text.push_str(&decoded);
        }
    }

    // ── End element dispatch ────────────────────────────────────────

    /// Returns `true` when the `</a:tbl>` closing tag is reached.
    fn handle_end(&mut self, e: &quick_xml::events::BytesEnd<'_>) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            b"tbl" => return true,
            b"tblPr" if self.is_in_tbl_pr => {
                self.is_in_tbl_pr = false;
            }
            b"tableStyleId" if self.is_in_table_style_id => {
                self.is_in_table_style_id = false;
            }
            b"tr" if self.is_in_row => {
                self.finish_row();
            }
            b"tc" if self.is_in_cell => {
                self.finish_cell();
            }
            b"txBody" if self.is_in_text_body => {
                self.is_in_text_body = false;
            }
            b"p" if self.is_in_paragraph => {
                self.finish_paragraph();
            }
            b"r" if self.is_in_run => {
                self.finish_run();
            }
            b"rPr" if self.is_in_run_properties => {
                self.is_in_run_properties = false;
            }
            b"endParaRPr" if self.is_in_end_paragraph_run_properties => {
                self.is_in_end_paragraph_run_properties = false;
            }
            b"lnSpc" if self.is_in_line_spacing => {
                self.is_in_line_spacing = false;
            }
            b"solidFill" if self.solid_fill_context != SolidFillCtx::None => {
                self.solid_fill_context = SolidFillCtx::None;
            }
            b"t" if self.is_in_text => {
                self.is_in_text = false;
            }
            b"tcPr" if self.is_in_table_cell_properties => {
                self.is_in_table_cell_properties = false;
            }
            b"lnL" | b"lnR" | b"lnT" | b"lnB" if self.is_in_border_line => {
                self.finish_border_line();
            }
            _ => {}
        }
        false
    }

    // ── Consume accumulated state into the final Table ──────────────

    fn finish(self) -> Table {
        let header_row_count: usize = if self.table_props.first_row { 1 } else { 0 };
        let mut table = Table {
            rows: self.rows,
            column_widths: self.column_widths,
            header_row_count,
            non_repeating_header_row_count: 0,
            alignment: None,
            default_cell_padding: Some(default_pptx_table_cell_padding()),
            use_content_driven_row_heights: true,
            default_vertical_align: None,
            // PowerPoint GT has not verified descender seating for bottom
            // cells (#618).
            seats_bottom_aligned_text_on_descender: false,
            bottom_aligned_descent_floor_pt: 0.0,
            border_paint_model: TableBorderPaintModel::CenteredStroke,
            prints_gridlines: false,
            prints_headings: false,
            centers_between_print_margins: false,
            print_scale: None,
        };
        table_styles::apply_table_style(&mut table, &self.table_props, self.table_styles);
        apply_table_theme_font_fallback(&mut table, self.theme);
        table
    }

    /// Extract tblPr attributes (firstRow, bandRow, etc.)
    fn parse_tbl_pr_attrs(&mut self, e: &BytesStart) {
        self.table_props.first_row = get_attr_str(e, b"firstRow").as_deref() == Some("1");
        self.table_props.last_row = get_attr_str(e, b"lastRow").as_deref() == Some("1");
        self.table_props.first_col = get_attr_str(e, b"firstCol").as_deref() == Some("1");
        self.table_props.last_col = get_attr_str(e, b"lastCol").as_deref() == Some("1");
        self.table_props.band_row = get_attr_str(e, b"bandRow").as_deref() == Some("1");
        self.table_props.band_col = get_attr_str(e, b"bandCol").as_deref() == Some("1");
    }

    // ── Private helpers: cell lifecycle ──────────────────────────────

    fn enter_cell(&mut self, e: &BytesStart) {
        self.is_in_cell = true;
        self.cell_col_span = get_attr_i64(e, b"gridSpan").map(|v| v as u32).unwrap_or(1);
        self.cell_row_span = get_attr_i64(e, b"rowSpan").map(|v| v as u32).unwrap_or(1);
        self.is_horizontal_merge = get_attr_str(e, b"hMerge").is_some();
        self.is_vertical_merge = get_attr_str(e, b"vMerge").is_some();
        self.cell_text_entries.clear();
        self.cell_background = None;
        self.cell_background_alpha = None;
        self.cell_fill_suppressed = false;
        self.cell_suppressed_borders = table_styles::SuppressedCellBorders::default();
        self.cell_vertical_align = None;
        self.cell_padding = None;
        self.is_in_table_cell_properties = false;
        self.border_left = None;
        self.border_right = None;
        self.border_top = None;
        self.border_bottom = None;
    }

    fn finish_cell(&mut self) {
        let has_border: bool = self.border_left.is_some()
            || self.border_right.is_some()
            || self.border_top.is_some()
            || self.border_bottom.is_some();

        let (col_span, row_span): (u32, u32) = if self.is_horizontal_merge {
            (0, 1)
        } else if self.is_vertical_merge {
            (1, 0)
        } else {
            (self.cell_col_span, self.cell_row_span)
        };

        self.cells.push(TableCell {
            content: group_pptx_text_blocks(std::mem::take(&mut self.cell_text_entries)),
            col_span,
            row_span,
            border: if has_border {
                Some(CellBorder {
                    left: self.border_left.take(),
                    right: self.border_right.take(),
                    top: self.border_top.take(),
                    bottom: self.border_bottom.take(),
                })
            } else {
                None
            },
            background: self.cell_background.take(),
            background_alpha: self.cell_background_alpha.take(),
            data_bar: None,
            sparkline: None,
            icon_text: None,
            icon_color: None,
            icon_shading: None,
            spill_width: None,
            vertical_align: self.cell_vertical_align.take(),
            padding: self.cell_padding.take(),
        });
        self.row_no_fill_cells.push(self.cell_fill_suppressed);
        self.row_suppressed_cell_borders
            .push(self.cell_suppressed_borders);
        self.is_in_cell = false;
        self.is_in_table_cell_properties = false;
    }

    // ── Private helpers: row lifecycle ───────────────────────────────

    fn finish_row(&mut self) {
        let minimum_height: Option<f64> = if self.row_height_emu > 0 {
            Some(emu_to_pt(self.row_height_emu))
        } else {
            None
        };
        self.rows.push(TableRow {
            // DrawingML `a:tr/@h` seeds the row's height, but PowerPoint lets
            // taller cell content grow past it. Keep it as the same floor the
            // renderer already uses for Word's at-least rows instead of a
            // fixed track that clips the cell (#1253).
            minimum_height,
            cells: std::mem::take(&mut self.cells),
            height: None,
        });
        self.table_props
            .no_fill_cells
            .push(std::mem::take(&mut self.row_no_fill_cells));
        self.table_props
            .suppressed_cell_borders
            .push(std::mem::take(&mut self.row_suppressed_cell_borders));
        self.is_in_row = false;
    }

    // ── Private helpers: paragraph lifecycle ─────────────────────────

    fn enter_paragraph(&mut self) {
        self.is_in_paragraph = true;
        self.paragraph_level = 0;
        self.paragraph_style = self
            .text_body_style_defaults
            .paragraph_style_for_level(self.paragraph_level);
        self.paragraph_default_run_style = self
            .text_body_style_defaults
            .run_style_for_level(self.paragraph_level);
        self.paragraph_end_run_style = self.paragraph_default_run_style.clone();
        self.paragraph_bullet_definition = self
            .text_body_style_defaults
            .bullet_for_level(self.paragraph_level);
        self.is_in_line_spacing = false;
        self.runs.clear();
        self.first_run_marker_style_override = None;
    }

    fn handle_paragraph_properties(&mut self, e: &BytesStart) {
        self.paragraph_level = extract_paragraph_level(e);
        self.paragraph_style = self
            .text_body_style_defaults
            .paragraph_style_for_level(self.paragraph_level);
        self.paragraph_default_run_style = self
            .text_body_style_defaults
            .run_style_for_level(self.paragraph_level);
        self.paragraph_end_run_style = self.paragraph_default_run_style.clone();
        self.paragraph_bullet_definition = self
            .text_body_style_defaults
            .bullet_for_level(self.paragraph_level);
        extract_paragraph_props(e, &mut self.paragraph_style);
    }

    fn finish_paragraph(&mut self) {
        let resolved_list_marker: Option<PptxListMarker> = resolve_pptx_list_marker(
            &self.paragraph_bullet_definition,
            self.paragraph_level,
            &self.runs,
            self.first_run_marker_style_override.as_ref(),
            &self.paragraph_end_run_style,
            &self.paragraph_default_run_style,
        );
        let mut paragraph_runs: Vec<Run> = std::mem::take(&mut self.runs);
        insert_hangul_kinsoku_break_markers(&mut paragraph_runs);
        let mut paragraph_style: ParagraphStyle = self.paragraph_style.clone();
        paragraph_style.paragraph_mark_font_family =
            pptx_paragraph_mark_font_family(&self.paragraph_end_run_style, self.theme);
        self.cell_text_entries.push(PptxParagraphEntry {
            paragraph: Paragraph {
                style: paragraph_style,
                runs: paragraph_runs,
            },
            list_marker: resolved_list_marker,
            paragraph_mark_font_size_pt: self.paragraph_end_run_style.font_size,
        });
        self.is_in_paragraph = false;
    }

    // ── Private helpers: run lifecycle ───────────────────────────────

    fn finish_run(&mut self) {
        if !self.run_text.is_empty() {
            if self.runs.is_empty() {
                self.first_run_marker_style_override =
                    self.run_marker_style_before_hyperlink.clone();
            }
            // A cell run that declares no `sz` anywhere in its own chain
            // takes the presentation's `p:defaultTextStyle` size. Without it
            // the run reached codegen with no size at all and Typst's own
            // 11pt default stood in (issue #675).
            let mut style: TextStyle = self.run_style.clone();
            if style.font_size.is_none() {
                style.font_size = self.default_text_size_pt;
            }
            push_pptx_run(
                &mut self.runs,
                Run {
                    text: std::mem::take(&mut self.run_text),
                    style,
                    href: None,
                    footnote: None,
                },
            );
        }
        self.is_in_run = false;
    }

    // ── Private helpers: bullet element dispatch ────────────────────

    /// Handles bullet-related elements that appear identically in both
    /// `Start` and `Empty` contexts. Returns `true` if the element was consumed.
    fn dispatch_bullet_element(&mut self, local_name: &[u8], e: &BytesStart) -> bool {
        match local_name {
            b"buAutoNum" => {
                self.paragraph_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                    parse_pptx_auto_numbering(e, self.paragraph_level),
                ));
            }
            b"buChar" => {
                self.paragraph_bullet_definition.kind =
                    parse_pptx_bullet_marker(e, self.paragraph_level);
            }
            b"buNone" => {
                self.paragraph_bullet_definition.kind = Some(PptxBulletKind::None);
            }
            b"buFontTx" => {
                self.paragraph_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
            }
            b"buFont" => {
                if let Some(typeface) = get_attr_str(e, b"typeface") {
                    self.paragraph_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                        resolve_theme_font(&typeface, self.theme),
                    ));
                }
            }
            b"buClrTx" => {
                self.paragraph_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
            }
            b"buClr" => {
                self.solid_fill_context = SolidFillCtx::BulletFill;
            }
            b"buSzTx" => {
                self.paragraph_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
            }
            b"buSzPct" => {
                if let Some(val) = get_attr_i64(e, b"val") {
                    self.paragraph_bullet_definition.size =
                        Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                }
            }
            b"buSzPts" => {
                if let Some(val) = get_attr_i64(e, b"val") {
                    self.paragraph_bullet_definition.size =
                        Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                }
            }
            b"br" => {
                push_pptx_soft_line_break(&mut self.runs, &self.paragraph_default_run_style);
            }
            _ => return false,
        }
        true
    }

    // ── Private helpers: non-bullet Start elements inside paragraph ─

    fn handle_start_non_bullet(
        &mut self,
        reader: &mut Reader<&[u8]>,
        local_name: &[u8],
        e: &BytesStart,
    ) -> Result<(), ConvertError> {
        match local_name {
            b"r" => {
                self.is_in_run = true;
                self.run_style = self.paragraph_default_run_style.clone();
                self.run_text.clear();
                self.run_has_explicit_underline = false;
                self.run_marker_style_before_hyperlink = None;
            }
            b"rPr" if self.is_in_run => {
                self.is_in_run_properties = true;
                self.run_has_explicit_underline = get_attr_str(e, b"u").is_some();
                extract_rpr_attributes(e, &mut self.run_style);
            }
            b"endParaRPr" => {
                self.is_in_end_paragraph_run_properties = true;
                self.paragraph_end_run_style = self.paragraph_default_run_style.clone();
                extract_rpr_attributes(e, &mut self.paragraph_end_run_style);
            }
            b"solidFill" if self.is_in_run_properties => {
                self.solid_fill_context = SolidFillCtx::RunFill;
            }
            b"solidFill" if self.is_in_end_paragraph_run_properties => {
                self.solid_fill_context = SolidFillCtx::EndParaFill;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr"
                if self.solid_fill_context != SolidFillCtx::None =>
            {
                self.apply_color_start(reader, e);
            }
            b"t" if self.is_in_run => {
                self.is_in_text = true;
            }
            b"hlinkClick" if self.is_in_run_properties => {
                if self.run_marker_style_before_hyperlink.is_none() {
                    self.run_marker_style_before_hyperlink = Some(self.run_style.clone());
                }
                apply_pptx_hyperlink_style(
                    &mut self.run_style,
                    self.run_has_explicit_underline,
                    self.theme,
                    self.color_map,
                );
            }
            _ => {}
        }
        Ok(())
    }

    // ── Private helpers: run/fill Start elements (when in_run) ──────

    fn handle_start_run_and_fill(
        &mut self,
        reader: &mut Reader<&[u8]>,
        local_name: &[u8],
        e: &BytesStart,
    ) -> Result<(), ConvertError> {
        match local_name {
            b"r" => {
                self.is_in_run = true;
                self.run_style = self.paragraph_default_run_style.clone();
                self.run_text.clear();
                self.run_has_explicit_underline = false;
                self.run_marker_style_before_hyperlink = None;
            }
            b"rPr" if self.is_in_run => {
                self.is_in_run_properties = true;
                self.run_has_explicit_underline = get_attr_str(e, b"u").is_some();
                extract_rpr_attributes(e, &mut self.run_style);
            }
            b"endParaRPr" if !self.is_in_run => {
                self.is_in_end_paragraph_run_properties = true;
                self.paragraph_end_run_style = self.paragraph_default_run_style.clone();
                extract_rpr_attributes(e, &mut self.paragraph_end_run_style);
            }
            b"solidFill" if self.is_in_run_properties => {
                self.solid_fill_context = SolidFillCtx::RunFill;
            }
            b"solidFill" if self.is_in_end_paragraph_run_properties => {
                self.solid_fill_context = SolidFillCtx::EndParaFill;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr"
                if self.solid_fill_context != SolidFillCtx::None =>
            {
                self.apply_color_start(reader, e);
            }
            b"t" if self.is_in_run => {
                self.is_in_text = true;
            }
            b"hlinkClick" if self.is_in_run_properties => {
                if self.run_marker_style_before_hyperlink.is_none() {
                    self.run_marker_style_before_hyperlink = Some(self.run_style.clone());
                }
                apply_pptx_hyperlink_style(
                    &mut self.run_style,
                    self.run_has_explicit_underline,
                    self.theme,
                    self.color_map,
                );
            }
            _ => {}
        }
        Ok(())
    }

    // ── Private helpers: non-bullet Empty elements inside paragraph ─

    fn handle_empty_non_bullet(&mut self, local_name: &[u8], e: &BytesStart) {
        match local_name {
            b"latin" | b"ea" | b"cs" if self.is_in_run_properties => {
                apply_typeface_to_style(e, &mut self.run_style, self.theme);
            }
            b"latin" | b"ea" | b"cs" if self.is_in_end_paragraph_run_properties => {
                apply_typeface_to_style(e, &mut self.paragraph_end_run_style, self.theme);
            }
            _ => {}
        }
    }

    // ── Private helpers: border line lifecycle ───────────────────────

    fn enter_border_line(&mut self, direction: BorderDir, e: &BytesStart) {
        self.is_in_border_line = true;
        self.current_border_dir = direction;
        self.border_line_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
        self.border_line_color = None;
        self.border_line_dash_style = BorderLineStyle::Solid;
    }

    fn finish_border_line(&mut self) {
        if let Some(color) = self.border_line_color.take() {
            let side = BorderSide {
                width: emu_to_pt(self.border_line_width_emu),
                color,
                style: self.border_line_dash_style,
                join: LineJoin::Round,
            };
            match self.current_border_dir {
                BorderDir::Left => self.border_left = Some(side),
                BorderDir::Right => self.border_right = Some(side),
                BorderDir::Top => self.border_top = Some(side),
                BorderDir::Bottom => self.border_bottom = Some(side),
                BorderDir::None => {}
            }
        }
        self.is_in_border_line = false;
        self.current_border_dir = BorderDir::None;
    }

    fn suppress_current_border(&mut self) {
        self.border_line_color = None;
        match self.current_border_dir {
            BorderDir::Left => self.cell_suppressed_borders.left = true,
            BorderDir::Right => self.cell_suppressed_borders.right = true,
            BorderDir::Top => self.cell_suppressed_borders.top = true,
            BorderDir::Bottom => self.cell_suppressed_borders.bottom = true,
            BorderDir::None => {}
        }
    }

    // ── Private helpers: color application ───────────────────────────

    fn apply_color_start(&mut self, reader: &mut Reader<&[u8]>, e: &BytesStart) {
        let parsed: ParsedColor = parse_color_from_start(reader, e, self.theme, self.color_map);
        self.apply_resolved_color(&parsed);
    }

    fn apply_color_empty(&mut self, e: &BytesStart) {
        let parsed: ParsedColor = parse_color_from_empty(e, self.theme, self.color_map);
        self.apply_resolved_color(&parsed);
    }

    fn apply_resolved_color(&mut self, parsed: &ParsedColor) {
        let color: Option<Color> = parsed.color;
        match self.solid_fill_context {
            SolidFillCtx::ShapeFill => {
                self.cell_background = color;
                self.cell_background_alpha = color.and(parsed.alpha);
            }
            SolidFillCtx::LineFill => self.border_line_color = color,
            // A cell's run states its opacity the same way a slide shape's run
            // does, so it composites the same way (issue #1121).
            SolidFillCtx::RunFill => {
                self.run_style.color = color;
                self.run_style.color_alpha = parsed.alpha;
            }
            SolidFillCtx::EndParaFill => {
                self.paragraph_end_run_style.color = color;
                self.paragraph_end_run_style.color_alpha = parsed.alpha;
            }
            SolidFillCtx::BulletFill => {
                self.paragraph_bullet_definition.color = color.map(PptxBulletColorSource::Explicit);
            }
            SolidFillCtx::PicLineFill | SolidFillCtx::None => {}
        }
    }
}

/// Fill the last gap in table-font inheritance after direct run formatting
/// and table-style regions have had their chance to supply a family.
fn apply_table_theme_font_fallback(table: &mut Table, theme: &ThemeData) {
    let Some(minor_font) = theme.minor_font.as_ref() else {
        return;
    };
    for row in &mut table.rows {
        for cell in &mut row.cells {
            for block in &mut cell.content {
                if let Block::Paragraph(paragraph) = block {
                    for run in &mut paragraph.runs {
                        if run.style.font_family.is_none() {
                            run.style.font_family = Some(minor_font.clone());
                        }
                    }
                }
            }
        }
    }
}

// ── Public entry point ──────────────────────────────────────────────

/// Parse a `<a:tbl>` element from the reader into a Table IR.
///
/// The reader should be positioned right after the `<a:tbl>` Start event.
/// Reads until the matching `</a:tbl>` End event.
pub(super) fn parse_pptx_table(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
    table_styles: &table_styles::TableStyleMap,
    default_text_size_pt: Option<f64>,
) -> Result<Table, ConvertError> {
    let mut state = PptxTableParser::new(theme, color_map, table_styles, default_text_size_pt);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                state.handle_start(reader, e)?;
            }
            Ok(Event::Empty(ref e)) => {
                state.handle_empty(e);
            }
            Ok(Event::Text(ref t)) => {
                state.handle_text(t);
            }
            Ok(Event::GeneralRef(ref reference)) => {
                state.handle_general_ref(reference);
            }
            Ok(Event::End(ref e)) => {
                if state.handle_end(e) {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(crate::parser::parse_err(format!(
                    "XML error in table: {error}"
                )));
            }
            _ => {}
        }
    }

    Ok(state.finish())
}

pub(super) fn scale_pptx_table_geometry_to_frame(
    table: &mut Table,
    frame_width_pt: f64,
    frame_height_pt: f64,
) {
    let intrinsic_width_pt: f64 = table.column_widths.iter().sum();
    if intrinsic_width_pt > 0.0 && frame_width_pt > 0.0 {
        let x_scale: f64 = frame_width_pt / intrinsic_width_pt;
        for width in &mut table.column_widths {
            *width *= x_scale;
        }
    }

    let intrinsic_height_pt: f64 = table.rows.iter().filter_map(|row| row.minimum_height).sum();
    if intrinsic_height_pt > 0.0 && frame_height_pt > intrinsic_height_pt {
        // A frame taller than the declared rows stretches them proportionally,
        // matching PowerPoint. A frame SHORTER than the rows is stale generator
        // output: PowerPoint grows the table instead of compressing rows, so
        // tr h acts as a minimum and must not be scaled down.
        let y_scale: f64 = frame_height_pt / intrinsic_height_pt;
        for row in &mut table.rows {
            if let Some(height) = row.minimum_height.as_mut() {
                *height *= y_scale;
            }
        }
    }
}
