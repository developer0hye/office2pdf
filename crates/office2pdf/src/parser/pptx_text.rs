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
    if typeface.trim().is_empty() {
        return;
    }
    let resolved = || resolve_theme_font(&typeface, theme);
    match element.local_name().as_ref() {
        b"latin" if style.font_family.is_none() => style.font_family = Some(resolved()),
        b"ea" if style.east_asian_font_family.is_none() => {
            style.east_asian_font_family = Some(resolved());
        }
        // TextStyle does not yet carry DrawingML's complex-script slot. It
        // must not be folded into the Latin slot: the issue #1333 fixture puts
        // `<a:cs typeface="Times New Roman"/>` on text that inherits Calibri,
        // while only its final punctuation declares `<a:latin>`.
        b"cs" => {}
        _ => {}
    }
}

/// Whether a paragraph wrote an `<a:endParaRPr>` at all, and — when it did
/// not — the runs whose face its mark then takes.
pub(super) enum PptxParagraphMark<'a> {
    /// The paragraph carries an `<a:endParaRPr>` element, declared typeface or
    /// not.
    Declared,
    /// The paragraph carries no `<a:endParaRPr>` element.
    Undeclared { runs: &'a [Run] },
}

impl<'a> PptxParagraphMark<'a> {
    /// The mark of a paragraph that did (`declares_end_para_rpr`) or did not
    /// write an `<a:endParaRPr>`, holding `runs` for the absent case.
    pub(super) fn for_paragraph(declares_end_para_rpr: bool, runs: &'a [Run]) -> Self {
        if declares_end_para_rpr {
            Self::Declared
        } else {
            Self::Undeclared { runs }
        }
    }
}

/// The family a paragraph's mark — the empty run `<a:endParaRPr>` describes —
/// ends up set in.
///
/// A mark that is **present** but declares no typeface inherits the run
/// properties the paragraph resolved, ending at the presentation's default
/// text style, whose `<a:latin typeface="+mn-lt"/>` names the theme's minor
/// Latin font. That fallback has to be the real face rather than the
/// renderer's own default, because the mark shares the final physical line's
/// box with its text fonts. For the golden mocks' unwrapped Arial paragraphs,
/// bare marks change the only line's ascent share from 0.97238em to 0.94377em
/// (#1176). Earlier wrapped lines exclude the mark (#1177).
///
/// A mark that is **absent** inherits nothing: it is typed in the face of the
/// text it follows, so it puts no family on the line the runs have not put
/// there already. Measured with two native one-factor probes over eight sizes
/// from 11pt to 53pt (`scripts/probes/issue-1645-*.json`, exported through
/// `probe_harness.py --backend office`): with the run naming its own typeface
/// and no `<a:endParaRPr>` present, replacing the body list style's
/// `<a:latin>` with Meiryo or Calibri — or the theme's minor Latin font with
/// either — moves no baseline at all, while declaring
/// `<a:endParaRPr><a:latin typeface="Meiryo"/></a:endParaRPr>` moves all eight
/// by up to 5.04pt, and a present-but-bare `<a:endParaRPr lang="en-US"/>` over
/// a Meiryo list style moves them by exactly as much. Inheriting the list
/// style's face into an absent mark seated the Arial role lines of issue
/// #1645 one point high, because the shared and unshared boxes straddle the
/// whole point PowerPoint rounds the story position to.
///
/// A paragraph with no runs has no such face, so its absent mark keeps the
/// inherited resolution — the blank line still has to be set in something.
pub(super) fn pptx_paragraph_mark_font_family(
    end_run_style: &TextStyle,
    theme: &ThemeData,
    mark: PptxParagraphMark<'_>,
) -> Option<Box<str>> {
    if let PptxParagraphMark::Undeclared { runs } = mark
        && !runs.is_empty()
    {
        return runs
            .iter()
            .rev()
            .find_map(|run| run.style.font_family.as_deref())
            .map(Box::from);
    }
    end_run_style
        .font_family
        .clone()
        .or_else(|| theme.minor_font.clone())
        .map(String::into_boxed_str)
}

// ── List style parser state machine ──────────────────────────────────

/// Which paragraph-level container the parser is currently inside.
#[derive(Clone, Copy)]
enum ParagraphTarget {
    Default,
    Level(u32),
}

#[derive(Clone, Copy)]
enum ParagraphSpacingTarget {
    Before,
    After,
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
    paragraph_spacing_target: Option<ParagraphSpacingTarget>,
    is_in_tab_list: bool,
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
            paragraph_spacing_target: None,
            is_in_tab_list: false,
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

    fn handle_paragraph_spacing_element(
        &mut self,
        e: &quick_xml::events::BytesStart,
        is_percent: bool,
    ) {
        let (Some(paragraph_target), Some(spacing_target)) =
            (self.active_paragraph_target, self.paragraph_spacing_target)
        else {
            return;
        };
        let style: &mut ParagraphStyle = self.paragraph_style_mut(paragraph_target);
        match spacing_target {
            ParagraphSpacingTarget::Before => {
                if is_percent {
                    extract_pptx_space_percent(e, &mut style.space_before_percent);
                    style.space_before = None;
                } else {
                    extract_pptx_space_points(e, &mut style.space_before);
                    style.space_before_percent = None;
                }
            }
            ParagraphSpacingTarget::After => {
                if is_percent {
                    extract_pptx_space_percent(e, &mut style.space_after_percent);
                    style.space_after = None;
                } else {
                    extract_pptx_space_points(e, &mut style.space_after);
                    style.space_after_percent = None;
                }
            }
        }
    }

    fn begin_tab_list(&mut self) {
        if let Some(target) = self.active_paragraph_target {
            self.paragraph_style_mut(target).tab_stops = Some(Vec::new());
            self.is_in_tab_list = true;
        }
    }

    fn handle_tab_stop(&mut self, e: &quick_xml::events::BytesStart) {
        if self.is_in_tab_list
            && let Some(target) = self.active_paragraph_target
        {
            extract_pptx_tab_stop(e, self.paragraph_style_mut(target));
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
            let style: &mut TextStyle = self.run_style_mut(target);
            style.color = parsed.color;
            style.color_alpha = parsed.alpha;
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
            let style: &mut TextStyle = self.run_style_mut(target);
            style.color = parsed.color;
            // A self-closing colour carries no `<a:alpha>` child, so it states
            // full opacity and clears any inherited half-tone.
            style.color_alpha = parsed.alpha;
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
    /// (`lstStyle` or a master `txStyles` bucket) is closed and parsing should stop.
    fn handle_end_element(&mut self, local_name: &[u8]) -> bool {
        match local_name {
            b"lstStyle" | b"otherStyle" | b"titleStyle" | b"bodyStyle" => return true,
            b"defPPr" => {
                self.active_paragraph_target = None;
                self.is_in_line_spacing = false;
                self.paragraph_spacing_target = None;
                self.is_in_tab_list = false;
            }
            name if parse_pptx_list_style_level(name).is_some() => {
                self.active_paragraph_target = None;
                self.is_in_line_spacing = false;
                self.paragraph_spacing_target = None;
                self.is_in_tab_list = false;
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
            b"spcBef" | b"spcAft" if self.paragraph_spacing_target.is_some() => {
                self.paragraph_spacing_target = None;
            }
            b"tabLst" if self.is_in_tab_list => {
                self.is_in_tab_list = false;
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
                    b"spcBef" if state.active_paragraph_target.is_some() => {
                        state.paragraph_spacing_target = Some(ParagraphSpacingTarget::Before);
                    }
                    b"spcAft" if state.active_paragraph_target.is_some() => {
                        state.paragraph_spacing_target = Some(ParagraphSpacingTarget::After);
                    }
                    b"tabLst" if state.active_paragraph_target.is_some() => {
                        state.begin_tab_list();
                    }
                    b"tab" if state.is_in_tab_list => {
                        state.handle_tab_stop(e);
                    }
                    b"spcPct" if state.is_in_line_spacing => {
                        state.handle_line_spacing_element(e, true);
                    }
                    b"spcPts" if state.is_in_line_spacing => {
                        state.handle_line_spacing_element(e, false);
                    }
                    b"spcPts" if state.paragraph_spacing_target.is_some() => {
                        state.handle_paragraph_spacing_element(e, false);
                    }
                    b"spcPct" if state.paragraph_spacing_target.is_some() => {
                        state.handle_paragraph_spacing_element(e, true);
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
                    b"spcPts" if state.paragraph_spacing_target.is_some() => {
                        state.handle_paragraph_spacing_element(e, false);
                    }
                    b"spcPct" if state.paragraph_spacing_target.is_some() => {
                        state.handle_paragraph_spacing_element(e, true);
                    }
                    b"tabLst" if state.active_paragraph_target.is_some() => {
                        state.begin_tab_list();
                        state.is_in_tab_list = false;
                    }
                    b"tab" if state.is_in_tab_list => {
                        state.handle_tab_stop(e);
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
    if let Some(value) = get_attr_i64(e, b"defTabSz")
        && value > 0
    {
        style.default_tab_stop_pt = Some(emu_to_pt(value));
    }
}

pub(super) fn extract_pptx_tab_stop(e: &quick_xml::events::BytesStart, style: &mut ParagraphStyle) {
    let Some(position_emu) = get_attr_i64(e, b"pos") else {
        return;
    };
    let alignment: TabAlignment = match get_attr_str(e, b"algn").as_deref() {
        Some("ctr") => TabAlignment::Center,
        Some("r") => TabAlignment::Right,
        Some("dec") => TabAlignment::Decimal,
        _ => TabAlignment::Left,
    };
    let tab_stop: TabStop = TabStop {
        position: emu_to_pt(position_emu),
        alignment,
        leader: TabLeader::None,
    };
    let tab_stops: &mut Vec<TabStop> = style.tab_stops.get_or_insert_with(Vec::new);
    tab_stops.push(tab_stop);
    tab_stops.sort_by(|left, right| left.position.total_cmp(&right.position));
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

/// `a:spcBef`/`a:spcAft` points value: hundredths of a point.
pub(super) fn extract_pptx_space_points(
    e: &quick_xml::events::BytesStart,
    target: &mut Option<f64>,
) {
    if let Some(value) = get_attr_i64(e, b"val") {
        *target = Some(value as f64 / 100.0);
    }
}

/// `a:spcBef`/`a:spcAft` percentage as a fraction of the paragraph's plain
/// line advance. Its point value cannot be resolved until the run size and
/// any saved normal-autofit font scale are known.
pub(super) fn extract_pptx_space_percent(
    e: &quick_xml::events::BytesStart,
    target: &mut Option<f64>,
) {
    if let Some(value) = get_attr_i64(e, b"val") {
        *target = Some(value as f64 / 100_000.0);
    }
}

/// Text-box layout settings accumulated from `<a:bodyPr>` and autofit hints.
#[derive(Debug, Clone, Copy)]
pub(super) struct PptxTextBoxSettings {
    pub(super) padding: Insets,
    pub(super) vertical_align: TextBoxVerticalAlign,
    pub(super) no_wrap: bool,
    pub(super) auto_fit: bool,
    /// PowerPoint's completed normal-autofit answer, as a fraction of each
    /// run's declared font size; the scaled size paints at a whole point (see
    /// [`round_pptx_scaled_font_size`]). `None` leaves font sizes unscaled;
    /// dynamic fitting is requested only when both saved autofit values are
    /// absent.
    pub(super) normal_autofit_font_scale: Option<f64>,
    /// Fraction subtracted from the original percentage line spacing by
    /// normal autofit.
    pub(super) normal_autofit_line_spacing_reduction: Option<f64>,
    pub(super) text_rotation_deg: Option<f64>,
}

impl Default for PptxTextBoxSettings {
    fn default() -> Self {
        Self {
            padding: default_pptx_text_box_padding(),
            vertical_align: TextBoxVerticalAlign::Top,
            no_wrap: false,
            auto_fit: false,
            normal_autofit_font_scale: None,
            normal_autofit_line_spacing_reduction: None,
            text_rotation_deg: None,
        }
    }
}

impl PptxTextBoxSettings {
    pub(super) fn requests_dynamic_autofit(&self) -> bool {
        self.auto_fit
            && self.normal_autofit_font_scale.is_none()
            && self.normal_autofit_line_spacing_reduction.is_none()
    }
}

fn parse_drawingml_percentage(value: &str) -> Option<f64> {
    let percentage: f64 = if let Some(percent) = value.strip_suffix('%') {
        percent.trim().parse::<f64>().ok()? / 100.0
    } else {
        value.trim().parse::<i64>().ok()? as f64 / 100_000.0
    };
    percentage.is_finite().then_some(percentage)
}

pub(super) fn extract_pptx_normal_autofit(
    element: &quick_xml::events::BytesStart,
    settings: &mut PptxTextBoxSettings,
) {
    settings.auto_fit = true;
    settings.normal_autofit_font_scale = get_attr_str(element, b"fontScale")
        .and_then(|value| parse_drawingml_percentage(&value))
        .filter(|value| (0.01..=1.0).contains(value));
    settings.normal_autofit_line_spacing_reduction = get_attr_str(element, b"lnSpcReduction")
        .and_then(|value| parse_drawingml_percentage(&value))
        .filter(|value| (0.0..=1.0).contains(value));
}

/// PowerPoint paints a saved `fontScale` at the nearest whole point, half up.
///
/// One-factor native probe (#1375: `fontScale` varied on a slide whose body
/// declares 22pt and its heading 35pt; macOS PowerPoint 16.x, Quartz
/// PDFContext): 62.5% paints 14/22pt, 70% 15/25pt, 75% 17/26pt, 80% 18/28pt,
/// 85% 19/30pt, 92.5% 20/32pt, 97.5% 21/34pt, and the #1303 fixture's 36pt
/// title at 90% paints 32pt. The exact halves 24.5 -> 25 and 16.5 -> 17 rule
/// out truncation and round-half-even, and the line advance follows the
/// rounded size too. The product is first settled on a 1/1000pt grid because
/// `35 * 0.7` is `24.499999...` in binary and would otherwise round down.
fn round_pptx_scaled_font_size(declared_size_pt: f64, font_scale: f64) -> f64 {
    const SETTLE_GRID_PER_PT: f64 = 1000.0;
    let scaled_pt: f64 =
        (declared_size_pt * font_scale * SETTLE_GRID_PER_PT).round() / SETTLE_GRID_PER_PT;
    scaled_pt.round()
}

fn scale_pptx_text_style_font_size(style: &mut TextStyle, font_scale: f64) {
    if let Some(font_size) = style.font_size.as_mut() {
        *font_size = round_pptx_scaled_font_size(*font_size, font_scale);
    }
}

/// Apply PowerPoint's saved normal-autofit result after run inheritance and
/// bullet sizing have resolved, so every painted size takes the same scale.
pub(super) fn apply_pptx_saved_normal_autofit(
    entries: &mut [PptxParagraphEntry],
    settings: &PptxTextBoxSettings,
) {
    for entry in entries {
        if let Some(font_scale) = settings.normal_autofit_font_scale {
            for run in &mut entry.paragraph.runs {
                scale_pptx_text_style_font_size(&mut run.style, font_scale);
            }
            match entry.list_marker.as_mut() {
                Some(PptxListMarker::Ordered {
                    marker_style: Some(style),
                    ..
                })
                | Some(PptxListMarker::Unordered {
                    marker_style: Some(style),
                    ..
                }) => scale_pptx_text_style_font_size(style, font_scale),
                _ => {}
            }
            if let Some(size) = entry.paragraph_mark_font_size_pt.as_mut() {
                *size = round_pptx_scaled_font_size(*size, font_scale);
            }
        }

        let Some(reduction) = settings.normal_autofit_line_spacing_reduction else {
            continue;
        };
        // ECMA-376 21.1.2.1.2 subtracts the reduction from the percentage:
        // 110% less 20% paces at 90%, not at 110% x 0.8 = 88%. The native
        // export of the #1375 deck paints that paragraph at the same pitch as
        // its unreduced 90% neighbours. A paragraph stating no percentage
        // starts from the 100% default.
        entry.paragraph.style.line_spacing = match entry.paragraph.style.line_spacing {
            Some(LineSpacing::Proportional(factor)) => {
                Some(LineSpacing::Proportional((factor - reduction).max(0.0)))
            }
            // ECMA-376 limits lnSpcReduction to percentage line spacing;
            // point-based a:spcPts is an absolute rule and stays unchanged.
            Some(LineSpacing::Exact(points)) => Some(LineSpacing::Exact(points)),
            None => Some(LineSpacing::Proportional(1.0 - reduction)),
        };
    }
}

pub(super) fn extract_pptx_text_box_body_props(
    e: &quick_xml::events::BytesStart,
    settings: &mut PptxTextBoxSettings,
) {
    parse_pptx_body_props(e).apply_to(settings);
}

/// One `<a:bodyPr>`'s attributes, each left `None` when the element does not
/// state it.
///
/// A placeholder's `<a:bodyPr>` inherits attribute by attribute from the
/// matching layout placeholder and then the master's (ECMA-376 §19.3.1.36),
/// so "absent" and "stated as the built-in default" are different answers and
/// [`PptxTextBoxSettings`] cannot represent the difference.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub(super) struct PptxBodyProps {
    /// Outer `Some` means `vert` was stated; the inner value is the rotation
    /// it resolves to, which is `None` for the horizontal writing modes.
    pub(super) text_rotation_deg: Option<Option<f64>>,
    pub(super) padding_left: Option<f64>,
    pub(super) padding_right: Option<f64>,
    pub(super) padding_top: Option<f64>,
    pub(super) padding_bottom: Option<f64>,
    pub(super) vertical_align: Option<TextBoxVerticalAlign>,
    pub(super) no_wrap: Option<bool>,
}

impl PptxBodyProps {
    /// Overlay `nearer`'s stated attributes onto `self`, which holds the ones
    /// from further up the inheritance chain.
    pub(super) fn overlay(&mut self, nearer: &Self) {
        let PptxBodyProps {
            text_rotation_deg,
            padding_left,
            padding_right,
            padding_top,
            padding_bottom,
            vertical_align,
            no_wrap,
        } = *nearer;
        self.text_rotation_deg = text_rotation_deg.or(self.text_rotation_deg);
        self.padding_left = padding_left.or(self.padding_left);
        self.padding_right = padding_right.or(self.padding_right);
        self.padding_top = padding_top.or(self.padding_top);
        self.padding_bottom = padding_bottom.or(self.padding_bottom);
        self.vertical_align = vertical_align.or(self.vertical_align);
        self.no_wrap = no_wrap.or(self.no_wrap);
    }

    pub(super) fn apply_to(&self, settings: &mut PptxTextBoxSettings) {
        if let Some(rotation) = self.text_rotation_deg {
            settings.text_rotation_deg = rotation;
        }
        if let Some(value) = self.padding_left {
            settings.padding.left = value;
        }
        if let Some(value) = self.padding_right {
            settings.padding.right = value;
        }
        if let Some(value) = self.padding_top {
            settings.padding.top = value;
        }
        if let Some(value) = self.padding_bottom {
            settings.padding.bottom = value;
        }
        if let Some(align) = self.vertical_align {
            settings.vertical_align = align;
        }
        // `wrap="square"` is the default and says nothing, so only the
        // no-wrap answer is carried; a nearer layer cannot turn it back off.
        if self.no_wrap == Some(true) {
            settings.no_wrap = true;
        }
    }
}

pub(super) fn parse_pptx_body_props(e: &quick_xml::events::BytesStart) -> PptxBodyProps {
    PptxBodyProps {
        text_rotation_deg: get_attr_str(e, b"vert").map(|vert| match vert.as_str() {
            // "vert" runs top-to-bottom (90° cw), "vert270" bottom-to-top.
            "vert" | "eaVert" | "mongolianVert" => Some(270.0),
            "vert270" => Some(90.0),
            _ => None,
        }),
        padding_left: get_attr_i64(e, b"lIns").map(emu_to_pt),
        padding_right: get_attr_i64(e, b"rIns").map(emu_to_pt),
        padding_top: get_attr_i64(e, b"tIns").map(emu_to_pt),
        padding_bottom: get_attr_i64(e, b"bIns").map(emu_to_pt),
        vertical_align: get_attr_str(e, b"anchor").map(|anchor| match anchor.as_str() {
            "ctr" => TextBoxVerticalAlign::Center,
            "b" => TextBoxVerticalAlign::Bottom,
            _ => TextBoxVerticalAlign::Top,
        }),
        no_wrap: get_attr_str(e, b"wrap").map(|wrap| wrap == "none"),
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
        // Slides carry no notes, so both sides are `None` here; comparing
              // presence keeps the merge total without requiring `Run: PartialEq`.
        && previous.footnote.is_none() == run.footnote.is_none()
    {
        previous.text.push_str(&run.text);
        return;
    }

    let mut run = run;
    normalize_pptx_run_boundary_spacing(runs.last(), &mut run);
    runs.push(run);
}

/// IR-only sentinel between a Hangul syllable and the terminal
/// punctuation that follows it (issue #438). UAX #14 LB13 forbids a break
/// before such a mark, so an overflowing mark drags the syllable onto the
/// next line; PowerPoint instead keeps the syllable in place and either
/// hangs the mark past the margin (Windows) or breaks before it (macOS).
/// The renderer converts the sentinel to an empty `#box[]` — a Contingent
/// Break in UAX #14, which restores the break opportunity — or drops it in
/// no-wrap runs; it never reaches the Typst source or the PDF text layer.
const HANGUL_KINSOKU_BREAK_CHAR: char = '\u{200B}';

fn is_hangul_syllable(character: char) -> bool {
    matches!(character, '\u{AC00}'..='\u{D7A3}')
}

/// Terminal marks PowerPoint refuses to push down with a preceding Hangul
/// syllable, measured against native exports in issue #438. `%` is
/// deliberately absent: PowerPoint keeps it glued to the syllable.
fn is_hangul_terminal_punct(character: char) -> bool {
    matches!(
        character,
        '?' | '!'
            | '.'
            | ','
            | ':'
            | ')'
            | '\u{201D}'
            | '\u{2026}'
            | '\u{FF1F}'
            | '\u{FF01}'
            | '\u{3002}'
            | '\u{3001}'
            | '\u{FF0E}'
            | '\u{FF0C}'
            | '\u{FF1A}'
            | '\u{FF09}'
    )
}

/// Grant a line-break opportunity between each Hangul syllable and the
/// terminal punctuation that follows it, including across run boundaries.
pub(super) fn insert_hangul_kinsoku_break_markers(runs: &mut [Run]) {
    let mut previous_character: Option<char> = None;
    for run in runs {
        if run.footnote.is_some() {
            previous_character = None;
            continue;
        }
        let needs_marker = {
            let mut previous = previous_character;
            run.text.chars().any(|character| {
                let pair =
                    previous.is_some_and(is_hangul_syllable) && is_hangul_terminal_punct(character);
                previous = Some(character);
                pair
            })
        };
        if needs_marker {
            let mut marked = String::with_capacity(run.text.len() + 8);
            let mut previous = previous_character;
            for character in run.text.chars() {
                if previous.is_some_and(is_hangul_syllable) && is_hangul_terminal_punct(character) {
                    marked.push(HANGUL_KINSOKU_BREAK_CHAR);
                }
                marked.push(character);
                previous = Some(character);
            }
            run.text = marked;
        }
        previous_character = run.text.chars().last().or(previous_character);
    }
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
    first_run_style_override: Option<&TextStyle>,
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> TextStyle {
    first_run_style_override
        .cloned()
        .or_else(|| first_pptx_visible_run_style(runs))
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
    first_run_style_override: Option<&TextStyle>,
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> Option<TextStyle> {
    let mut style = resolve_pptx_marker_base_style(
        runs,
        first_run_style_override,
        end_para_run_style,
        default_run_style,
    );
    // PowerPoint lets a bullet follow the visible run's font, size, and
    // colour, but its underline belongs to the paragraph's list-level style.
    // A hyperlink run and its trailing `<a:endParaRPr u="sng">` therefore
    // underline the text without drawing a rule below an inherited bullet
    // (issue #1353). Conversely, a list-level underline remains on the marker
    // even when the first run explicitly turns its own underline off.
    style.underline = default_run_style.underline;

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
            style.color_alpha = None;
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
    first_run_style_override: Option<&TextStyle>,
    end_para_run_style: &TextStyle,
    default_run_style: &TextStyle,
) -> Option<PptxListMarker> {
    let marker_style = resolve_pptx_marker_style(
        bullet,
        runs,
        first_run_style_override,
        end_para_run_style,
        default_run_style,
    );
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
        // CT_TextBulletCharacter is one character. Normalize malformed cache
        // values that repeat the marker so they cannot become multiple marks.
        .and_then(|marker| marker.chars().next().map(|character| character.to_string()))
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
    for entry in &mut entries {
        resolve_pptx_percentage_paragraph_spacing(entry);
        preserve_blank_pptx_list_item(entry);
    }

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

/// Resolve DrawingML's percentage before/after spacing against the
/// paragraph's unmodified 1.2em line advance. `lnSpcReduction` changes the
/// spacing *between lines of the paragraph*; it does not shrink this separate
/// paragraph gap. Visible text decides the size, while an empty paragraph
/// falls back to its inherited paragraph mark.
fn resolve_pptx_percentage_paragraph_spacing(entry: &mut PptxParagraphEntry) {
    const PLAIN_LINE_ADVANCE_FACTOR: f64 = 1.2;

    let font_size_pt: f64 = entry
        .paragraph
        .runs
        .iter()
        .filter_map(|run| run.style.font_size)
        .reduce(f64::max)
        .or(entry.paragraph_mark_font_size_pt)
        .unwrap_or(crate::defaults::TYPST_DEFAULT_FONT_SIZE_PT);
    let plain_line_advance_pt: f64 = PLAIN_LINE_ADVANCE_FACTOR * font_size_pt;
    let style: &mut ParagraphStyle = &mut entry.paragraph.style;
    if let Some(percent) = style.space_before_percent.take() {
        style.space_before = Some(percent * plain_line_advance_pt);
    }
    if let Some(percent) = style.space_after_percent.take() {
        style.space_after = Some(percent * plain_line_advance_pt);
    }
}

/// Keep an inherited list paragraph in the IR while marking that its list
/// marker is absent. PowerPoint still reserves the paragraph mark's line box;
/// its resolved family and size travel on the sentinel run so the renderer can
/// create that exact blank line without leaking a private-use glyph.
fn preserve_blank_pptx_list_item(entry: &mut PptxParagraphEntry) {
    if entry.list_marker.is_none() || pptx_paragraph_has_visible_content(&entry.paragraph) {
        return;
    }

    entry.paragraph.runs.push(Run {
        text: PPTX_BLANK_LIST_ITEM_SENTINEL.to_string(),
        style: TextStyle {
            font_family: entry
                .paragraph
                .style
                .paragraph_mark_font_family
                .as_deref()
                .map(str::to_string),
            font_size: entry.paragraph_mark_font_size_pt,
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
    });
}

fn pptx_paragraph_has_visible_content(paragraph: &Paragraph) -> bool {
    if pptx_paragraph_is_blank_list_item(paragraph) {
        return false;
    }
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
    // PowerPoint cases the run at render time and leaves the stored text
    // alone, so a title written mixed-case prints uppercase (issue #875).
    // `none` is an explicit override of an inherited `cap` and has to state
    // both answers rather than stay silent.
    if let Some(val) = get_attr_str(e, b"cap") {
        style.all_caps = Some(val == "all");
        style.small_caps = Some(val == "small");
    }
    if let Some(sz) = get_attr_i64(e, b"sz") {
        // Font size in hundredths of a point (e.g. 1200 = 12pt)
        style.font_size = Some(sz as f64 / 100.0);
    }
    if let Some(baseline) = get_attr_i64(e, b"baseline") {
        // DrawingML stores the displacement in thousandths of a percent.
        // Keeping it font-relative lets inherited run sizes resolve first.
        style.baseline_shift = Some(BaselineShiftEm(baseline as f64 / 100_000.0));
    }
    if let Some(spc) = get_attr_i64(e, b"spc") {
        // Character tracking, also in hundredths of a point, added to every
        // character gap in the run. Negative values tighten (e.g. -100 = -1pt).
        style.letter_spacing = Some(spc as f64 / 100.0);
    }
    if let Some(kern) = get_attr_i64(e, b"kern") {
        // DrawingML's `kern` is the size *threshold* pair kerning starts at,
        // in hundredths of a point — not a switch. A master's `titleStyle`
        // typically states `kern="1200"`, so PowerPoint kerns every title from
        // 12pt up; the deck of issue #1073 sets its 38pt titles that way and
        // tightens `TA`/`AT` by ~1.9pt each. Reading it here covers `a:rPr`
        // and the `a:defRPr` of every list style, which is where decks
        // actually state it (issue #1073).
        style.pair_kerning = Some(PairKerning::from_threshold_pt(kern as f64 / 100.0));
    }
}

/// Apply PowerPoint's hyperlink colour precedence while preserving an authored underline choice.
pub(super) fn apply_pptx_hyperlink_style(
    style: &mut TextStyle,
    has_explicit_underline: bool,
    theme: &ThemeData,
    color_map: &ColorMapData,
) {
    if let Some(color) = resolve_scheme_color(theme, color_map, "hlink") {
        style.color = Some(color);
        // The resolved theme colour retains no alpha, so clear any inherited
        // or run-local half-tone rather than compositing the theme RGB at it.
        style.color_alpha = None;
    }
    if !has_explicit_underline {
        style.underline = Some(true);
    }
}
