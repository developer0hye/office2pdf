use super::*;

pub(super) fn merge_paragraph_style(target: &mut ParagraphStyle, source: &ParagraphStyle) {
    if source.alignment.is_some() {
        target.alignment = source.alignment;
    }
    if source.indent_left.is_some() {
        target.indent_left = source.indent_left;
    }
    if source.indent_right.is_some() {
        target.indent_right = source.indent_right;
    }
    if source.indent_first_line.is_some() {
        target.indent_first_line = source.indent_first_line;
    }
    if source.line_spacing.is_some() {
        target.line_spacing = source.line_spacing;
    }
    if source.space_before.is_some() {
        target.space_before = source.space_before;
    }
    if source.space_after.is_some() {
        target.space_after = source.space_after;
    }
    if source.heading_level.is_some() {
        target.heading_level = source.heading_level;
    }
    if source.direction.is_some() {
        target.direction = source.direction;
    }
    if source.tab_stops.is_some() {
        target.tab_stops = source.tab_stops.clone();
    }
}

pub(super) fn merge_text_style(target: &mut TextStyle, source: &TextStyle) {
    if source.font_family.is_some() {
        target.font_family = source.font_family.clone();
    }
    if source.font_size.is_some() {
        target.font_size = source.font_size;
    }
    if source.bold.is_some() {
        target.bold = source.bold;
    }
    if source.italic.is_some() {
        target.italic = source.italic;
    }
    if source.underline.is_some() {
        target.underline = source.underline;
    }
    if source.strikethrough.is_some() {
        target.strikethrough = source.strikethrough;
    }
    if source.color.is_some() {
        target.color = source.color;
    }
    if source.highlight.is_some() {
        target.highlight = source.highlight;
    }
    if source.vertical_align.is_some() {
        target.vertical_align = source.vertical_align;
    }
    if source.all_caps.is_some() {
        target.all_caps = source.all_caps;
    }
    if source.small_caps.is_some() {
        target.small_caps = source.small_caps;
    }
    if source.letter_spacing.is_some() {
        target.letter_spacing = source.letter_spacing;
    }
}

pub(super) fn merge_pptx_bullet_definition(
    target: &mut PptxBulletDefinition,
    source: &PptxBulletDefinition,
) {
    if source.kind.is_some() {
        target.kind = source.kind.clone();
    }
    if source.font.is_some() {
        target.font = source.font.clone();
    }
    if source.color.is_some() {
        target.color = source.color.clone();
    }
    if source.size.is_some() {
        target.size = source.size.clone();
    }
}

pub(super) fn parse_pptx_list_style_level(name: &[u8]) -> Option<u32> {
    if name.len() != 7 || !name.starts_with(b"lvl") || !name.ends_with(b"pPr") {
        return None;
    }
    let digit = name[3];
    if !(b'1'..=b'9').contains(&digit) {
        return None;
    }
    Some(u32::from(digit - b'1'))
}

pub(super) fn apply_typeface_to_style(
    element: &quick_xml::events::BytesStart,
    style: &mut TextStyle,
    theme: &ThemeData,
) {
    let Some(typeface) = get_attr_str(element, b"typeface") else {
        return;
    };
    if typeface.trim().is_empty() || style.font_family.is_some() {
        return;
    }
    style.font_family = Some(resolve_theme_font(&typeface, theme));
}

pub(super) fn parse_pptx_list_style(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> PptxTextBodyStyleDefaults {
    #[derive(Clone, Copy)]
    enum ParagraphTarget {
        Default,
        Level(u32),
    }

    let mut defaults = PptxTextBodyStyleDefaults::default();
    let mut active_paragraph_target: Option<ParagraphTarget> = None;
    let mut active_run_target: Option<ParagraphTarget> = None;
    let mut in_ln_spc = false;
    let mut in_run_fill = false;
    let mut in_bullet_fill = false;

    fn paragraph_style_mut(
        defaults: &mut PptxTextBodyStyleDefaults,
        target: ParagraphTarget,
    ) -> &mut ParagraphStyle {
        match target {
            ParagraphTarget::Default => &mut defaults.default_paragraph,
            ParagraphTarget::Level(level) => {
                &mut defaults.levels.entry(level).or_default().paragraph
            }
        }
    }

    fn run_style_mut(
        defaults: &mut PptxTextBodyStyleDefaults,
        target: ParagraphTarget,
    ) -> &mut TextStyle {
        match target {
            ParagraphTarget::Default => &mut defaults.default_run,
            ParagraphTarget::Level(level) => &mut defaults.levels.entry(level).or_default().run,
        }
    }

    fn bullet_style_mut(
        defaults: &mut PptxTextBodyStyleDefaults,
        target: ParagraphTarget,
    ) -> &mut PptxBulletDefinition {
        match target {
            ParagraphTarget::Default => &mut defaults.default_bullet,
            ParagraphTarget::Level(level) => &mut defaults.levels.entry(level).or_default().bullet,
        }
    }

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"defPPr" => {
                        active_paragraph_target = Some(ParagraphTarget::Default);
                        extract_paragraph_props(
                            e,
                            paragraph_style_mut(&mut defaults, ParagraphTarget::Default),
                        );
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        let level = parse_pptx_list_style_level(name).unwrap();
                        active_paragraph_target = Some(ParagraphTarget::Level(level));
                        extract_paragraph_props(
                            e,
                            paragraph_style_mut(&mut defaults, ParagraphTarget::Level(level)),
                        );
                    }
                    b"lnSpc" if active_paragraph_target.is_some() => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pct(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"spcPts" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pts(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"buAutoNum" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind = Some(
                                PptxBulletKind::AutoNumber(parse_pptx_auto_numbering(e, level)),
                            );
                        }
                    }
                    b"buChar" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind =
                                parse_pptx_bullet_marker(e, level);
                        }
                    }
                    b"buNone" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).kind =
                                Some(PptxBulletKind::None);
                        }
                    }
                    b"buFontTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::FollowText);
                        }
                    }
                    b"buFont" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(typeface) = get_attr_str(e, b"typeface")
                        {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::Explicit(resolve_theme_font(
                                    &typeface, theme,
                                )));
                        }
                    }
                    b"buClrTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).color =
                                Some(PptxBulletColorSource::FollowText);
                        }
                    }
                    b"buClr" if active_paragraph_target.is_some() => {
                        in_bullet_fill = true;
                    }
                    b"buSzTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::FollowText);
                        }
                    }
                    b"buSzPct" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"defRPr" if active_paragraph_target.is_some() => {
                        active_run_target = active_paragraph_target;
                        if let Some(target) = active_run_target {
                            extract_rpr_attributes(e, run_style_mut(&mut defaults, target));
                        }
                    }
                    b"solidFill" if active_run_target.is_some() => {
                        in_run_fill = true;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_run_fill => {
                        let parsed = parse_color_from_start(reader, e, theme, color_map);
                        if let Some(target) = active_run_target {
                            run_style_mut(&mut defaults, target).color = parsed.color;
                        }
                    }
                    b"latin" | b"ea" | b"cs" if active_run_target.is_some() => {
                        if let Some(target) = active_run_target {
                            apply_typeface_to_style(e, run_style_mut(&mut defaults, target), theme);
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_bullet_fill => {
                        if let Some(target) = active_paragraph_target {
                            let parsed = parse_color_from_start(reader, e, theme, color_map);
                            bullet_style_mut(&mut defaults, target).color =
                                parsed.color.map(PptxBulletColorSource::Explicit);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"defPPr" => {
                        extract_paragraph_props(e, &mut defaults.default_paragraph);
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        let level = parse_pptx_list_style_level(name).unwrap();
                        extract_paragraph_props(
                            e,
                            &mut defaults.levels.entry(level).or_default().paragraph,
                        );
                    }
                    b"spcPct" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pct(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"spcPts" if in_ln_spc => {
                        if let Some(target) = active_paragraph_target {
                            extract_pptx_line_spacing_pts(
                                e,
                                paragraph_style_mut(&mut defaults, target),
                            );
                        }
                    }
                    b"buAutoNum" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind = Some(
                                PptxBulletKind::AutoNumber(parse_pptx_auto_numbering(e, level)),
                            );
                        }
                    }
                    b"buChar" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            let level = match target {
                                ParagraphTarget::Default => 0,
                                ParagraphTarget::Level(level) => level,
                            };
                            bullet_style_mut(&mut defaults, target).kind =
                                parse_pptx_bullet_marker(e, level);
                        }
                    }
                    b"buNone" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).kind =
                                Some(PptxBulletKind::None);
                        }
                    }
                    b"buFontTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::FollowText);
                        }
                    }
                    b"buFont" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(typeface) = get_attr_str(e, b"typeface")
                        {
                            bullet_style_mut(&mut defaults, target).font =
                                Some(PptxBulletFontSource::Explicit(resolve_theme_font(
                                    &typeface, theme,
                                )));
                        }
                    }
                    b"buClrTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).color =
                                Some(PptxBulletColorSource::FollowText);
                        }
                    }
                    b"buClr" if active_paragraph_target.is_some() => {}
                    b"buSzTx" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::FollowText);
                        }
                    }
                    b"buSzPct" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target
                            && let Some(val) = get_attr_i64(e, b"val")
                        {
                            bullet_style_mut(&mut defaults, target).size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"defRPr" if active_paragraph_target.is_some() => {
                        if let Some(target) = active_paragraph_target {
                            extract_rpr_attributes(e, run_style_mut(&mut defaults, target));
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_run_fill => {
                        let parsed = parse_color_from_empty(e, theme, color_map);
                        if let Some(target) = active_run_target {
                            run_style_mut(&mut defaults, target).color = parsed.color;
                        }
                    }
                    b"latin" | b"ea" | b"cs" if active_run_target.is_some() => {
                        if let Some(target) = active_run_target {
                            apply_typeface_to_style(e, run_style_mut(&mut defaults, target), theme);
                        }
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if in_bullet_fill => {
                        if let Some(target) = active_paragraph_target {
                            let parsed = parse_color_from_empty(e, theme, color_map);
                            bullet_style_mut(&mut defaults, target).color =
                                parsed.color.map(PptxBulletColorSource::Explicit);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"lstStyle" | b"otherStyle" => break,
                    b"defPPr" => {
                        active_paragraph_target = None;
                        in_ln_spc = false;
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        active_paragraph_target = None;
                        in_ln_spc = false;
                    }
                    b"defRPr" => {
                        active_run_target = None;
                        in_run_fill = false;
                    }
                    b"solidFill" if in_run_fill => {
                        in_run_fill = false;
                    }
                    b"buClr" if in_bullet_fill => {
                        in_bullet_fill = false;
                    }
                    b"lnSpc" if in_ln_spc => {
                        in_ln_spc = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    defaults
}

pub(super) fn extract_paragraph_props(
    e: &quick_xml::events::BytesStart,
    style: &mut ParagraphStyle,
) {
    if let Some(algn) = get_attr_str(e, b"algn") {
        style.alignment = match algn.as_str() {
            "l" => Some(Alignment::Left),
            "ctr" => Some(Alignment::Center),
            "r" => Some(Alignment::Right),
            "just" => Some(Alignment::Justify),
            _ => None,
        };
    }
    if let Some(val) = get_attr_str(e, b"rtl")
        && (val == "1" || val == "true")
    {
        style.direction = Some(TextDirection::Rtl);
    }
    if let Some(value) = get_attr_i64(e, b"marL") {
        style.indent_left = Some(emu_to_pt(value));
    }
    if let Some(value) = get_attr_i64(e, b"marR") {
        style.indent_right = Some(emu_to_pt(value));
    }
    if let Some(value) = get_attr_i64(e, b"indent") {
        style.indent_first_line = Some(emu_to_pt(value));
    }
}

pub(super) fn extract_pptx_line_spacing_pct(
    e: &quick_xml::events::BytesStart,
    style: &mut ParagraphStyle,
) {
    if let Some(value) = get_attr_i64(e, b"val") {
        style.line_spacing = Some(LineSpacing::Proportional(value as f64 / 100_000.0));
    }
}

pub(super) fn extract_pptx_line_spacing_pts(
    e: &quick_xml::events::BytesStart,
    style: &mut ParagraphStyle,
) {
    if let Some(value) = get_attr_i64(e, b"val") {
        style.line_spacing = Some(LineSpacing::Exact(value as f64 / 100.0));
    }
}

pub(super) fn extract_pptx_text_box_body_props(
    e: &quick_xml::events::BytesStart,
    padding: &mut Insets,
    vertical_align: &mut TextBoxVerticalAlign,
) {
    if let Some(value) = get_attr_i64(e, b"lIns") {
        padding.left = emu_to_pt(value);
    }
    if let Some(value) = get_attr_i64(e, b"rIns") {
        padding.right = emu_to_pt(value);
    }
    if let Some(value) = get_attr_i64(e, b"tIns") {
        padding.top = emu_to_pt(value);
    }
    if let Some(value) = get_attr_i64(e, b"bIns") {
        padding.bottom = emu_to_pt(value);
    }
    if let Some(anchor) = get_attr_str(e, b"anchor") {
        *vertical_align = match anchor.as_str() {
            "ctr" => TextBoxVerticalAlign::Center,
            "b" => TextBoxVerticalAlign::Bottom,
            _ => TextBoxVerticalAlign::Top,
        };
    }
}

pub(super) fn extract_pptx_table_cell_props(
    e: &quick_xml::events::BytesStart,
    vertical_align: &mut Option<CellVerticalAlign>,
    padding: &mut Option<Insets>,
) {
    if let Some(anchor) = get_attr_str(e, b"anchor") {
        *vertical_align = Some(match anchor.as_str() {
            "ctr" => CellVerticalAlign::Center,
            "b" => CellVerticalAlign::Bottom,
            _ => CellVerticalAlign::Top,
        });
    }

    let mut cell_padding = (*padding).unwrap_or_default();
    let mut has_padding = false;
    if let Some(value) = get_attr_i64(e, b"marL") {
        cell_padding.left = emu_to_pt(value);
        has_padding = true;
    }
    if let Some(value) = get_attr_i64(e, b"marR") {
        cell_padding.right = emu_to_pt(value);
        has_padding = true;
    }
    if let Some(value) = get_attr_i64(e, b"marT") {
        cell_padding.top = emu_to_pt(value);
        has_padding = true;
    }
    if let Some(value) = get_attr_i64(e, b"marB") {
        cell_padding.bottom = emu_to_pt(value);
        has_padding = true;
    }
    if has_padding {
        *padding = Some(cell_padding);
    }
}

pub(super) fn push_pptx_run(runs: &mut Vec<Run>, run: Run) {
    if let Some(previous) = runs.last_mut()
        && previous.style == run.style
        && previous.href == run.href
        && previous.footnote == run.footnote
    {
        previous.text.push_str(&run.text);
        return;
    }

    let mut run = run;
    normalize_pptx_run_boundary_spacing(runs.last(), &mut run);
    runs.push(run);
}

pub(super) fn push_pptx_soft_line_break(runs: &mut Vec<Run>, style: &TextStyle) {
    push_pptx_run(
        runs,
        Run {
            text: PPTX_SOFT_LINE_BREAK_CHAR.to_string(),
            style: style.clone(),
            href: None,
            footnote: None,
        },
    );
}

pub(super) fn decode_pptx_text_event(text: &quick_xml::events::BytesText<'_>) -> Option<String> {
    let decoded = text.decode().ok()?;
    let unescaped = unescape_xml_text(decoded.as_ref()).ok()?;
    Some(unescaped.into_owned())
}

pub(super) fn decode_pptx_general_ref(
    reference: &quick_xml::events::BytesRef<'_>,
) -> Option<String> {
    let decoded = reference.decode().ok()?;
    let wrapped = format!("&{};", decoded.as_ref());
    let unescaped = unescape_xml_text(&wrapped).ok()?;
    Some(unescaped.into_owned())
}

fn normalize_pptx_run_boundary_spacing(previous: Option<&Run>, run: &mut Run) {
    let Some(previous) = previous else {
        return;
    };

    if previous.href != run.href
        || previous.footnote.is_some()
        || run.footnote.is_some()
        || previous
            .text
            .chars()
            .last()
            .is_some_and(char::is_whitespace)
    {
        return;
    }

    let mut chars = run.text.chars();
    let Some(first_char) = chars.next() else {
        return;
    };
    let Some(next_char) = chars.next() else {
        return;
    };

    if first_char == ' ' && should_preserve_pptx_run_boundary_space(next_char) {
        // PowerPoint often splits styled phrases into adjacent runs such as
        // `K` + ` = 100)`. Preserve that boundary space as non-breaking so
        // Typst does not wrap at the style change and spill punctuation.
        run.text.replace_range(0..1, "\u{00A0}");
    }
}

fn should_preserve_pptx_run_boundary_space(next_char: char) -> bool {
    matches!(
        next_char,
        '=' | '+' | '-' | '/' | '%' | ')' | ']' | '}' | ':' | ';' | ',' | '.'
    )
}

pub(super) fn first_pptx_visible_run_style(runs: &[Run]) -> Option<TextStyle> {
    runs.iter()
        .find(|run| !run.text.is_empty() && run.footnote.is_none())
        .map(|run| run.style.clone())
}

fn resolve_pptx_marker_base_style(
    runs: &[Run],
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> TextStyle {
    first_pptx_visible_run_style(runs)
        .or_else(|| {
            (end_para_run_style != &TextStyle::default()).then(|| end_para_run_style.clone())
        })
        .unwrap_or_else(|| default_run_style.clone())
}

fn finalize_pptx_marker_style(style: TextStyle) -> Option<TextStyle> {
    (style != TextStyle::default()).then_some(style)
}

pub(super) fn resolve_pptx_marker_style(
    bullet: &PptxBulletDefinition,
    runs: &[Run],
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> Option<TextStyle> {
    let mut style = resolve_pptx_marker_base_style(runs, end_para_run_style, default_run_style);

    match bullet.font.as_ref() {
        Some(PptxBulletFontSource::FollowText) | None => {}
        Some(PptxBulletFontSource::Explicit(font_family)) => {
            style.font_family = Some(font_family.clone());
        }
    }

    match bullet.color.as_ref() {
        Some(PptxBulletColorSource::FollowText) | None => {}
        Some(PptxBulletColorSource::Explicit(color)) => {
            style.color = Some(*color);
        }
    }

    match bullet.size.as_ref() {
        Some(PptxBulletSizeSource::FollowText) | None => {}
        Some(PptxBulletSizeSource::Points(points)) => {
            style.font_size = Some(*points);
        }
        Some(PptxBulletSizeSource::Percent(percent)) => {
            style.font_size = style.font_size.map(|size| size * percent);
        }
    }

    finalize_pptx_marker_style(style)
}

pub(super) fn resolve_pptx_list_marker(
    bullet: &PptxBulletDefinition,
    level: u32,
    runs: &[Run],
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> Option<PptxListMarker> {
    let marker_style =
        resolve_pptx_marker_style(bullet, runs, end_para_run_style, default_run_style);
    match bullet.kind.as_ref()? {
        PptxBulletKind::None => None,
        PptxBulletKind::Character(character) => Some(PptxListMarker::Unordered {
            level,
            marker_text: character.clone(),
            marker_style,
        }),
        PptxBulletKind::AutoNumber(auto_numbering) => Some(PptxListMarker::Ordered {
            auto_numbering: auto_numbering.clone(),
            marker_style,
        }),
    }
}

pub(super) fn extract_paragraph_level(e: &quick_xml::events::BytesStart) -> u32 {
    get_attr_i64(e, b"lvl")
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0)
}

pub(super) fn parse_pptx_auto_numbering(
    e: &quick_xml::events::BytesStart,
    level: u32,
) -> PptxAutoNumbering {
    let numbering_pattern: Option<String> = get_attr_str(e, b"type")
        .as_deref()
        .and_then(pptx_auto_numbering_pattern)
        .map(str::to_string);
    let start_at: Option<u32> = get_attr_i64(e, b"startAt").and_then(|value| value.try_into().ok());

    PptxAutoNumbering {
        level,
        numbering_pattern,
        start_at,
    }
}

pub(super) fn parse_pptx_bullet_marker(
    e: &quick_xml::events::BytesStart,
    level: u32,
) -> Option<PptxBulletKind> {
    get_attr_str(e, b"char")
        .map(PptxBulletKind::Character)
        .or_else(|| (level == 0).then(|| PptxBulletKind::Character("•".to_string())))
}

fn pptx_auto_numbering_pattern(numbering_type: &str) -> Option<&'static str> {
    match numbering_type {
        "arabicPeriod" => Some("1."),
        "arabicParenR" => Some("1)"),
        "arabicParenBoth" => Some("(1)"),
        "alphaLcPeriod" => Some("a."),
        "alphaUcPeriod" => Some("A."),
        "alphaLcParenR" => Some("a)"),
        "alphaUcParenR" => Some("A)"),
        "romanLcPeriod" => Some("i."),
        "romanUcPeriod" => Some("I."),
        "romanLcParenR" => Some("i)"),
        "romanUcParenR" => Some("I)"),
        _ => None,
    }
}

pub(super) fn group_pptx_text_blocks(entries: Vec<PptxParagraphEntry>) -> Vec<Block> {
    let mut entries = entries;
    trim_trailing_empty_pptx_list_entries(&mut entries);

    let mut blocks: Vec<Block> = Vec::new();
    let mut pending_list: Option<PendingPptxList> = None;

    for entry in entries {
        match entry.list_marker {
            Some(list_marker) => {
                if pending_list
                    .as_ref()
                    .is_some_and(|list| !list.can_extend(&list_marker))
                {
                    blocks.push(pending_list.take().unwrap().into_block());
                }

                let paragraph: Paragraph = entry.paragraph;
                pending_list
                    .get_or_insert_with(|| PendingPptxList::new(&list_marker))
                    .push(paragraph, list_marker);
            }
            None => {
                if let Some(list) = pending_list.take() {
                    blocks.push(list.into_block());
                }
                blocks.push(Block::Paragraph(entry.paragraph));
            }
        }
    }

    if let Some(list) = pending_list {
        blocks.push(list.into_block());
    }

    blocks
}

fn trim_trailing_empty_pptx_list_entries(entries: &mut Vec<PptxParagraphEntry>) {
    while entries.len() > 1 {
        let Some(last_entry) = entries.last() else {
            break;
        };
        if last_entry.list_marker.is_none()
            || pptx_paragraph_has_visible_content(&last_entry.paragraph)
        {
            break;
        }
        entries.pop();
    }
}

fn pptx_paragraph_has_visible_content(paragraph: &Paragraph) -> bool {
    paragraph.runs.iter().any(|run| {
        run.footnote.is_some()
            || run.text.chars().any(|character| {
                character != PPTX_SOFT_LINE_BREAK_CHAR && !character.is_whitespace()
            })
    })
}

pub(super) fn extract_rpr_attributes(e: &quick_xml::events::BytesStart, style: &mut TextStyle) {
    if let Some(val) = get_attr_str(e, b"b") {
        style.bold = Some(val == "1" || val == "true");
    }
    if let Some(val) = get_attr_str(e, b"i") {
        style.italic = Some(val == "1" || val == "true");
    }
    if let Some(val) = get_attr_str(e, b"u") {
        style.underline = Some(val != "none");
    }
    if let Some(val) = get_attr_str(e, b"strike") {
        style.strikethrough = Some(val != "noStrike");
    }
    if let Some(sz) = get_attr_i64(e, b"sz") {
        // Font size in hundredths of a point (e.g. 1200 = 12pt)
        style.font_size = Some(sz as f64 / 100.0);
    }
}
