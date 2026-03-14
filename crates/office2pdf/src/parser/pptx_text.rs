use super::*;

/// Overwrite each `Option` field in `target` with `source` when the source is `Some`.
/// Fields that require `.clone()` must be listed after a `;` separator.
/// Kept only for PPTX-specific types that don't live in the IR layer.
macro_rules! merge_option_fields {
    ($target:expr, $source:expr, $($copy_field:ident),* $(; $($clone_field:ident),*)?) => {
        $(
            if $source.$copy_field.is_some() {
                $target.$copy_field = $source.$copy_field;
            }
        )*
        $($(
            if $source.$clone_field.is_some() {
                $target.$clone_field = $source.$clone_field.clone();
            }
        )*)?
    };
}

pub(super) fn merge_pptx_bullet_definition(
    target: &mut PptxBulletDefinition,
    source: &PptxBulletDefinition,
) {
    merge_option_fields!(target, source, ; kind, font, color, size);
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

// ── List style parser state machine ──────────────────────────────────

/// Which paragraph-level container the parser is currently inside.
#[derive(Clone, Copy)]
enum ParagraphTarget {
    Default,
    Level(u32),
}

impl ParagraphTarget {
    /// Return the numeric level (0 for `Default`).
    fn level(self) -> u32 {
        match self {
            Self::Default => 0,
            Self::Level(level) => level,
        }
    }
}

/// Tracks which XML context the parser is currently inside while
/// walking the children of `<a:lstStyle>` / `<a:otherStyle>`.
struct ListStyleParseState {
    defaults: PptxTextBodyStyleDefaults,
    active_paragraph_target: Option<ParagraphTarget>,
    active_run_target: Option<ParagraphTarget>,
    is_in_line_spacing: bool,
    is_in_run_fill: bool,
    is_in_bullet_fill: bool,
}

impl ListStyleParseState {
    fn new() -> Self {
        Self {
            defaults: PptxTextBodyStyleDefaults::default(),
            active_paragraph_target: None,
            active_run_target: None,
            is_in_line_spacing: false,
            is_in_run_fill: false,
            is_in_bullet_fill: false,
        }
    }

    fn paragraph_style_mut(&mut self, target: ParagraphTarget) -> &mut ParagraphStyle {
        match target {
            ParagraphTarget::Default => &mut self.defaults.default_paragraph,
            ParagraphTarget::Level(level) => {
                &mut self.defaults.levels.entry(level).or_default().paragraph
            }
        }
    }

    fn run_style_mut(&mut self, target: ParagraphTarget) -> &mut TextStyle {
        match target {
            ParagraphTarget::Default => &mut self.defaults.default_run,
            ParagraphTarget::Level(level) => {
                &mut self.defaults.levels.entry(level).or_default().run
            }
        }
    }

    fn bullet_style_mut(&mut self, target: ParagraphTarget) -> &mut PptxBulletDefinition {
        match target {
            ParagraphTarget::Default => &mut self.defaults.default_bullet,
            ParagraphTarget::Level(level) => {
                &mut self.defaults.levels.entry(level).or_default().bullet
            }
        }
    }

    // ── Paragraph-level element handlers ─────────────────────────────

    /// Enter a `<defPPr>` or `<lvlNpPr>` element (Start or Empty).
    fn enter_paragraph_target(
        &mut self,
        target: ParagraphTarget,
        e: &quick_xml::events::BytesStart,
    ) {
        self.active_paragraph_target = Some(target);
        extract_paragraph_props(e, self.paragraph_style_mut(target));
    }

    /// Handle `<spcPct>` / `<spcPts>` inside `<lnSpc>`.
    fn handle_line_spacing_element(&mut self, e: &quick_xml::events::BytesStart, is_pct: bool) {
        if let Some(target) = self.active_paragraph_target {
            let style: &mut ParagraphStyle = self.paragraph_style_mut(target);
            if is_pct {
                extract_pptx_line_spacing_pct(e, style);
            } else {
                extract_pptx_line_spacing_pts(e, style);
            }
        }
    }

    // ── Bullet element handlers ──────────────────────────────────────

    fn handle_bullet_auto_num(&mut self, e: &quick_xml::events::BytesStart) {
        if let Some(target) = self.active_paragraph_target {
            let level: u32 = target.level();
            self.bullet_style_mut(target).kind = Some(PptxBulletKind::AutoNumber(
                parse_pptx_auto_numbering(e, level),
            ));
        }
    }

    fn handle_bullet_char(&mut self, e: &quick_xml::events::BytesStart) {
        if let Some(target) = self.active_paragraph_target {
            let level: u32 = target.level();
            self.bullet_style_mut(target).kind = parse_pptx_bullet_marker(e, level);
        }
    }

    fn handle_bullet_none(&mut self) {
        if let Some(target) = self.active_paragraph_target {
            self.bullet_style_mut(target).kind = Some(PptxBulletKind::None);
        }
    }

    fn handle_bullet_font_follow_text(&mut self) {
        if let Some(target) = self.active_paragraph_target {
            self.bullet_style_mut(target).font = Some(PptxBulletFontSource::FollowText);
        }
    }

    fn handle_bullet_font_explicit(
        &mut self,
        e: &quick_xml::events::BytesStart,
        theme: &ThemeData,
    ) {
        if let Some(target) = self.active_paragraph_target
            && let Some(typeface) = get_attr_str(e, b"typeface")
        {
            self.bullet_style_mut(target).font = Some(PptxBulletFontSource::Explicit(
                resolve_theme_font(&typeface, theme),
            ));
        }
    }

    fn handle_bullet_color_follow_text(&mut self) {
        if let Some(target) = self.active_paragraph_target {
            self.bullet_style_mut(target).color = Some(PptxBulletColorSource::FollowText);
        }
    }

    fn handle_bullet_size_follow_text(&mut self) {
        if let Some(target) = self.active_paragraph_target {
            self.bullet_style_mut(target).size = Some(PptxBulletSizeSource::FollowText);
        }
    }

    fn handle_bullet_size_pct(&mut self, e: &quick_xml::events::BytesStart) {
        if let Some(target) = self.active_paragraph_target
            && let Some(val) = get_attr_i64(e, b"val")
        {
            self.bullet_style_mut(target).size =
                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
        }
    }

    fn handle_bullet_size_pts(&mut self, e: &quick_xml::events::BytesStart) {
        if let Some(target) = self.active_paragraph_target
            && let Some(val) = get_attr_i64(e, b"val")
        {
            self.bullet_style_mut(target).size =
                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
        }
    }

    // ── Run-level element handlers ───────────────────────────────────

    /// Enter `<defRPr>` as a Start element (sets `active_run_target`).
    fn enter_default_run_props(&mut self, e: &quick_xml::events::BytesStart) {
        self.active_run_target = self.active_paragraph_target;
        if let Some(target) = self.active_run_target {
            extract_rpr_attributes(e, self.run_style_mut(target));
        }
    }

    /// Handle `<defRPr/>` as an Empty element (no run target activation).
    fn handle_default_run_props_empty(&mut self, e: &quick_xml::events::BytesStart) {
        if let Some(target) = self.active_paragraph_target {
            extract_rpr_attributes(e, self.run_style_mut(target));
        }
    }

    fn handle_run_color_start(
        &mut self,
        reader: &mut Reader<&[u8]>,
        e: &quick_xml::events::BytesStart,
        theme: &ThemeData,
        color_map: &ColorMapData,
    ) {
        let parsed: ParsedColor = parse_color_from_start(reader, e, theme, color_map);
        if let Some(target) = self.active_run_target {
            self.run_style_mut(target).color = parsed.color;
        }
    }

    fn handle_run_color_empty(
        &mut self,
        e: &quick_xml::events::BytesStart,
        theme: &ThemeData,
        color_map: &ColorMapData,
    ) {
        let parsed: ParsedColor = parse_color_from_empty(e, theme, color_map);
        if let Some(target) = self.active_run_target {
            self.run_style_mut(target).color = parsed.color;
        }
    }

    fn handle_typeface(&mut self, e: &quick_xml::events::BytesStart, theme: &ThemeData) {
        if let Some(target) = self.active_run_target {
            apply_typeface_to_style(e, self.run_style_mut(target), theme);
        }
    }

    fn handle_bullet_color_start(
        &mut self,
        reader: &mut Reader<&[u8]>,
        e: &quick_xml::events::BytesStart,
        theme: &ThemeData,
        color_map: &ColorMapData,
    ) {
        if let Some(target) = self.active_paragraph_target {
            let parsed: ParsedColor = parse_color_from_start(reader, e, theme, color_map);
            self.bullet_style_mut(target).color = parsed.color.map(PptxBulletColorSource::Explicit);
        }
    }

    fn handle_bullet_color_empty(
        &mut self,
        e: &quick_xml::events::BytesStart,
        theme: &ThemeData,
        color_map: &ColorMapData,
    ) {
        if let Some(target) = self.active_paragraph_target {
            let parsed: ParsedColor = parse_color_from_empty(e, theme, color_map);
            self.bullet_style_mut(target).color = parsed.color.map(PptxBulletColorSource::Explicit);
        }
    }

    // ── Dispatch: shared bullet/run element handling ─────────────────

    /// Handle elements that appear identically in both `Start` and `Empty`
    /// contexts for bullet properties. Returns `true` if the element was handled.
    fn dispatch_bullet_element(
        &mut self,
        local_name: &[u8],
        e: &quick_xml::events::BytesStart,
        theme: &ThemeData,
    ) -> bool {
        if self.active_paragraph_target.is_none() {
            return false;
        }
        match local_name {
            b"buAutoNum" => self.handle_bullet_auto_num(e),
            b"buChar" => self.handle_bullet_char(e),
            b"buNone" => self.handle_bullet_none(),
            b"buFontTx" => self.handle_bullet_font_follow_text(),
            b"buFont" => self.handle_bullet_font_explicit(e, theme),
            b"buClrTx" => self.handle_bullet_color_follow_text(),
            b"buSzTx" => self.handle_bullet_size_follow_text(),
            b"buSzPct" => self.handle_bullet_size_pct(e),
            b"buSzPts" => self.handle_bullet_size_pts(e),
            _ => return false,
        }
        true
    }

    // ── End-element state transitions ────────────────────────────────

    /// Process an `Event::End` element. Returns `true` when the outer container
    /// (`lstStyle` / `otherStyle`) is closed and parsing should stop.
    fn handle_end_element(&mut self, local_name: &[u8]) -> bool {
        match local_name {
            b"lstStyle" | b"otherStyle" => return true,
            b"defPPr" => {
                self.active_paragraph_target = None;
                self.is_in_line_spacing = false;
            }
            name if parse_pptx_list_style_level(name).is_some() => {
                self.active_paragraph_target = None;
                self.is_in_line_spacing = false;
            }
            b"defRPr" => {
                self.active_run_target = None;
                self.is_in_run_fill = false;
            }
            b"solidFill" if self.is_in_run_fill => {
                self.is_in_run_fill = false;
            }
            b"buClr" if self.is_in_bullet_fill => {
                self.is_in_bullet_fill = false;
            }
            b"lnSpc" if self.is_in_line_spacing => {
                self.is_in_line_spacing = false;
            }
            _ => {}
        }
        false
    }
}

pub(super) fn parse_pptx_list_style(
    reader: &mut Reader<&[u8]>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> PptxTextBodyStyleDefaults {
    let mut state = ListStyleParseState::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let local_name: &[u8] = local.as_ref();
                match local_name {
                    b"defPPr" => {
                        state.enter_paragraph_target(ParagraphTarget::Default, e);
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        let level: u32 = parse_pptx_list_style_level(name).unwrap();
                        state.enter_paragraph_target(ParagraphTarget::Level(level), e);
                    }
                    b"lnSpc" if state.active_paragraph_target.is_some() => {
                        state.is_in_line_spacing = true;
                    }
                    b"spcPct" if state.is_in_line_spacing => {
                        state.handle_line_spacing_element(e, true);
                    }
                    b"spcPts" if state.is_in_line_spacing => {
                        state.handle_line_spacing_element(e, false);
                    }
                    b"buClr" if state.active_paragraph_target.is_some() => {
                        state.is_in_bullet_fill = true;
                    }
                    b"defRPr" if state.active_paragraph_target.is_some() => {
                        state.enter_default_run_props(e);
                    }
                    b"solidFill" if state.active_run_target.is_some() => {
                        state.is_in_run_fill = true;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if state.is_in_run_fill => {
                        state.handle_run_color_start(reader, e, theme, color_map);
                    }
                    b"latin" | b"ea" | b"cs" if state.active_run_target.is_some() => {
                        state.handle_typeface(e, theme);
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if state.is_in_bullet_fill => {
                        state.handle_bullet_color_start(reader, e, theme, color_map);
                    }
                    _ => {
                        state.dispatch_bullet_element(local_name, e, theme);
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let local_name: &[u8] = local.as_ref();
                match local_name {
                    b"defPPr" => {
                        extract_paragraph_props(e, &mut state.defaults.default_paragraph);
                    }
                    name if parse_pptx_list_style_level(name).is_some() => {
                        let level: u32 = parse_pptx_list_style_level(name).unwrap();
                        extract_paragraph_props(
                            e,
                            &mut state.defaults.levels.entry(level).or_default().paragraph,
                        );
                    }
                    b"spcPct" if state.is_in_line_spacing => {
                        state.handle_line_spacing_element(e, true);
                    }
                    b"spcPts" if state.is_in_line_spacing => {
                        state.handle_line_spacing_element(e, false);
                    }
                    b"buClr" if state.active_paragraph_target.is_some() => {
                        // Empty `<buClr/>` — no color data to extract.
                    }
                    b"defRPr" if state.active_paragraph_target.is_some() => {
                        state.handle_default_run_props_empty(e);
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if state.is_in_run_fill => {
                        state.handle_run_color_empty(e, theme, color_map);
                    }
                    b"latin" | b"ea" | b"cs" if state.active_run_target.is_some() => {
                        state.handle_typeface(e, theme);
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr" if state.is_in_bullet_fill => {
                        state.handle_bullet_color_empty(e, theme, color_map);
                    }
                    _ => {
                        state.dispatch_bullet_element(local_name, e, theme);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if state.handle_end_element(e.local_name().as_ref()) {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    state.defaults
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
    no_wrap: &mut bool,
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
    if get_attr_str(e, b"wrap").as_deref() == Some("none") {
        *no_wrap = true;
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
