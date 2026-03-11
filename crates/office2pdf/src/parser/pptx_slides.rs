use super::package::{
    load_chart_data, load_slide_images, load_smartart_data, resolve_layout_master_paths,
    scan_chart_refs,
};
use super::*;

// ── Slide inheritance chain ─────────────────────────────────────────────

/// Resolved XML content and color maps for the master -> layout -> slide chain.
struct SlideInheritanceChain {
    slide_xml: String,
    slide_color_map: ColorMapData,
    layout_path: Option<String>,
    layout_xml: Option<String>,
    layout_color_map: Option<ColorMapData>,
    master_path: Option<String>,
    master_xml: Option<String>,
    master_color_map: ColorMapData,
    master_text_style_defaults: PptxTextBodyStyleDefaults,
}

/// Build the full inheritance chain by reading master/layout/slide XML and
/// resolving each layer's effective color map from a single master base.
fn resolve_inheritance_chain<R: Read + std::io::Seek>(
    slide_path: &str,
    theme: &ThemeData,
    archive: &mut ZipArchive<R>,
) -> Result<SlideInheritanceChain, ConvertError> {
    let slide_xml: String = read_zip_entry(archive, slide_path)?;
    let (layout_path, master_path) = resolve_layout_master_paths(slide_path, archive);

    let master_xml: Option<String> = master_path
        .as_ref()
        .and_then(|path| read_zip_entry(archive, path).ok());
    let layout_xml: Option<String> = layout_path
        .as_ref()
        .and_then(|path| read_zip_entry(archive, path).ok());

    let master_color_map: ColorMapData = master_xml
        .as_deref()
        .map(parse_master_color_map)
        .unwrap_or_else(default_color_map);
    let master_text_style_defaults: PptxTextBodyStyleDefaults = master_xml
        .as_deref()
        .map(|xml| parse_master_other_style(xml, theme, &master_color_map))
        .unwrap_or_default();

    let slide_color_map: ColorMapData = resolve_effective_color_map(&slide_xml, &master_color_map);
    let layout_color_map: Option<ColorMapData> = layout_xml
        .as_deref()
        .map(|xml| resolve_effective_color_map(xml, &master_color_map));

    Ok(SlideInheritanceChain {
        slide_xml,
        slide_color_map,
        layout_path,
        layout_xml,
        layout_color_map,
        master_path,
        master_xml,
        master_color_map,
        master_text_style_defaults,
    })
}

/// Parse elements from a single inheritance layer (master or layout).
/// Broken layers are non-fatal and silently return empty results.
fn parse_layer_elements<R: Read + std::io::Seek>(
    layer_path: &str,
    layer_xml: &str,
    color_map: &ColorMapData,
    theme: &ThemeData,
    label: &str,
    text_style_defaults: &PptxTextBodyStyleDefaults,
    archive: &mut ZipArchive<R>,
) -> (Vec<FixedElement>, Vec<ConvertWarning>) {
    let images: SlideImageMap = load_slide_images(layer_path, archive);
    parse_slide_xml(
        layer_xml,
        &images,
        theme,
        color_map,
        label,
        text_style_defaults,
    )
    .unwrap_or_default()
}

// ── Embedded object helpers ─────────────────────────────────────────────

/// Collect SmartArt elements referenced by the slide XML.
fn collect_smartart_elements<R: Read + std::io::Seek>(
    slide_xml: &str,
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> Vec<FixedElement> {
    let smartart_refs = smartart::scan_smartart_refs(slide_xml);
    if smartart_refs.is_empty() {
        return Vec::new();
    }

    let smartart_data = load_smartart_data(slide_path, archive);
    smartart_refs
        .iter()
        .filter_map(|sa_ref| {
            smartart_data
                .get(&sa_ref.data_rid)
                .map(|items| FixedElement {
                    x: emu_to_pt(sa_ref.x),
                    y: emu_to_pt(sa_ref.y),
                    width: emu_to_pt(sa_ref.cx),
                    height: emu_to_pt(sa_ref.cy),
                    kind: FixedElementKind::SmartArt(SmartArt {
                        items: items.clone(),
                    }),
                })
        })
        .collect()
}

/// Collect Chart elements referenced by the slide XML.
fn collect_chart_elements<R: Read + std::io::Seek>(
    slide_xml: &str,
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> Vec<FixedElement> {
    let chart_refs = scan_chart_refs(slide_xml);
    if chart_refs.is_empty() {
        return Vec::new();
    }

    let chart_data = load_chart_data(slide_path, archive);
    chart_refs
        .iter()
        .filter_map(|c_ref| {
            chart_data.get(&c_ref.chart_rid).map(|chart| FixedElement {
                x: emu_to_pt(c_ref.x),
                y: emu_to_pt(c_ref.y),
                width: emu_to_pt(c_ref.cx),
                height: emu_to_pt(c_ref.cy),
                kind: FixedElementKind::Chart(chart.clone()),
            })
        })
        .collect()
}

// ── Background resolution ───────────────────────────────────────────────

/// Resolve the slide background by checking slide -> layout -> master in order.
/// If a gradient is found on the slide, its first stop color is used as the
/// solid fallback; otherwise the first solid color found in the chain wins.
fn resolve_slide_background(
    chain: &SlideInheritanceChain,
    theme: &ThemeData,
) -> (Option<Color>, Option<GradientFill>) {
    let gradient = parse_background_gradient(&chain.slide_xml, theme, &chain.slide_color_map);

    if gradient.is_some() {
        let fallback_color = gradient
            .as_ref()
            .and_then(|g| g.stops.first().map(|s| s.color));
        return (fallback_color, gradient);
    }

    let solid_color = parse_background_color(&chain.slide_xml, theme, &chain.slide_color_map)
        .or_else(|| {
            chain.layout_xml.as_deref().and_then(|xml| {
                chain
                    .layout_color_map
                    .as_ref()
                    .and_then(|map| parse_background_color(xml, theme, map))
            })
        })
        .or_else(|| {
            chain
                .master_xml
                .as_deref()
                .and_then(|xml| parse_background_color(xml, theme, &chain.master_color_map))
        });

    (solid_color, None)
}

// ── Public entry point ──────────────────────────────────────────────────

/// Parse a single slide from the archive, returning a Page or an error.
///
/// Resolves the inheritance chain (slide -> layout -> master) and
/// prepends master/layout elements behind slide elements.
pub(super) fn parse_single_slide<R: Read + std::io::Seek>(
    slide_path: &str,
    slide_label: &str,
    slide_size: PageSize,
    theme: &ThemeData,
    archive: &mut ZipArchive<R>,
) -> Result<(Page, Vec<ConvertWarning>), ConvertError> {
    let chain: SlideInheritanceChain = resolve_inheritance_chain(slide_path, theme, archive)?;

    let slide_images: SlideImageMap = load_slide_images(slide_path, archive);
    let mut warnings: Vec<ConvertWarning> = Vec::new();

    let (slide_elements, slide_warnings) = parse_slide_xml(
        &chain.slide_xml,
        &slide_images,
        theme,
        &chain.slide_color_map,
        slide_label,
        &chain.master_text_style_defaults,
    )?;
    warnings.extend(slide_warnings);

    let mut elements: Vec<FixedElement> = Vec::new();

    // Master layer (bottom)
    if let Some(ref path) = chain.master_path
        && let Some(ref xml) = chain.master_xml
    {
        let master_label: String = format!("{slide_label} master");
        let (master_elems, master_warnings) = parse_layer_elements(
            path,
            xml,
            &chain.master_color_map,
            theme,
            &master_label,
            &chain.master_text_style_defaults,
            archive,
        );
        elements.extend(master_elems);
        warnings.extend(master_warnings);
    }

    // Layout layer (middle)
    if let Some(ref path) = chain.layout_path
        && let Some(ref xml) = chain.layout_xml
        && let Some(ref color_map) = chain.layout_color_map
    {
        let layout_label: String = format!("{slide_label} layout");
        let (layout_elems, layout_warnings) = parse_layer_elements(
            path,
            xml,
            color_map,
            theme,
            &layout_label,
            &chain.master_text_style_defaults,
            archive,
        );
        elements.extend(layout_elems);
        warnings.extend(layout_warnings);
    }

    // Slide layer (top)
    elements.extend(slide_elements);

    // Embedded objects
    elements.extend(collect_smartart_elements(
        &chain.slide_xml,
        slide_path,
        archive,
    ));
    elements.extend(collect_chart_elements(
        &chain.slide_xml,
        slide_path,
        archive,
    ));

    let (background_color, background_gradient) = resolve_slide_background(&chain, theme);

    Ok((
        Page::Fixed(FixedPage {
            size: slide_size,
            elements,
            background_color,
            background_gradient,
        }),
        warnings,
    ))
}

fn describe_assets(assets: impl IntoIterator<Item = String>) -> String {
    assets.into_iter().collect::<Vec<_>>().join(", ")
}

fn pick_supported_asset(rid: &str, images: &SlideImageMap) -> Option<SlideImageAsset> {
    images
        .get(rid)
        .filter(|asset| asset.is_supported())
        .cloned()
}

fn select_picture_asset(
    images: &SlideImageMap,
    warning_context: &str,
    base_rid: Option<&str>,
    svg_rid: Option<&str>,
    img_layer_rids: &[String],
) -> (Option<SlideImageAsset>, Vec<ConvertWarning>) {
    let mut warnings = Vec::new();

    let unsupported_layers: Vec<String> = img_layer_rids
        .iter()
        .filter_map(|rid| images.get(rid))
        .filter(|asset| !asset.is_supported())
        .map(|asset| asset.file_name().to_string())
        .collect();
    if !unsupported_layers.is_empty() {
        warnings.push(ConvertWarning::PartialElement {
            format: "PPTX".to_string(),
            element: format!("{warning_context} picture"),
            detail: format!(
                "unsupported image layer omitted: {}",
                describe_assets(unsupported_layers)
            ),
        });
    }

    let selected = svg_rid
        .and_then(|rid| pick_supported_asset(rid, images))
        .or_else(|| base_rid.and_then(|rid| pick_supported_asset(rid, images)))
        .or_else(|| {
            img_layer_rids
                .iter()
                .find_map(|rid| pick_supported_asset(rid, images))
        });
    if selected.is_some() {
        return (selected, warnings);
    }

    let omitted_assets = svg_rid
        .into_iter()
        .chain(base_rid)
        .chain(img_layer_rids.iter().map(String::as_str))
        .filter_map(|rid| images.get(rid))
        .map(|asset| asset.file_name().to_string())
        .collect::<Vec<_>>();
    if !omitted_assets.is_empty() {
        warnings.push(ConvertWarning::UnsupportedElement {
            format: "PPTX".to_string(),
            element: format!(
                "{warning_context} image omitted: {}",
                describe_assets(omitted_assets)
            ),
        });
    }

    (None, warnings)
}

// ── State structs ───────────────────────────────────────────────────────

/// Accumulated state for a `<p:pic>` element.
#[derive(Default)]
struct PictureState {
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    blip_embed: Option<String>,
    svg_blip_embed: Option<String>,
    img_layer_embeds: Vec<String>,
    crop: Option<ImageCrop>,
    in_xfrm: bool,
}

impl PictureState {
    fn reset(&mut self) {
        *self = Self::default();
    }
}

/// Accumulated state for a `<p:graphicFrame>` element.
#[derive(Default)]
struct GraphicFrameState {
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    in_xfrm: bool,
}

impl GraphicFrameState {
    fn reset(&mut self) {
        *self = Self::default();
    }
}

/// Accumulated state for a `<p:sp>` (shape) element and its nested properties.
struct ShapeState {
    depth: usize,
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    has_placeholder: bool,
    rotation_deg: Option<f64>,
    opacity: Option<f64>,
    shadow: Option<Shadow>,
    in_sp_pr: bool,
    prst_geom: Option<String>,
    fill: Option<Color>,
    gradient_fill: Option<GradientFill>,
    in_xfrm: bool,
    in_ln: bool,
    ln_width_emu: i64,
    ln_color: Option<Color>,
    ln_dash_style: BorderLineStyle,
}

impl Default for ShapeState {
    fn default() -> Self {
        Self {
            depth: 0,
            x: 0,
            y: 0,
            cx: 0,
            cy: 0,
            has_placeholder: false,
            rotation_deg: None,
            opacity: None,
            shadow: None,
            in_sp_pr: false,
            prst_geom: None,
            fill: None,
            gradient_fill: None,
            in_xfrm: false,
            in_ln: false,
            ln_width_emu: 0,
            ln_color: None,
            ln_dash_style: BorderLineStyle::Solid,
        }
    }
}

impl ShapeState {
    fn reset(&mut self) {
        *self = Self::default();
    }
}

// ── Finalization helpers ────────────────────────────────────────────────

/// Finalize a shape element when `</p:sp>` is reached.
/// Returns either a TextBox (if the shape has text) or a Shape geometry element.
fn finalize_shape(
    shape: &mut ShapeState,
    paragraphs: &mut Vec<PptxParagraphEntry>,
    text_box_padding: Insets,
    text_box_vertical_align: TextBoxVerticalAlign,
) -> Option<FixedElement> {
    let has_text = paragraphs
        .iter()
        .any(|entry| !entry.paragraph.runs.is_empty());

    if has_text {
        let blocks: Vec<Block> = group_pptx_text_blocks(std::mem::take(paragraphs));
        Some(FixedElement {
            x: emu_to_pt(shape.x),
            y: emu_to_pt(shape.y),
            width: emu_to_pt(shape.cx),
            height: emu_to_pt(shape.cy),
            kind: FixedElementKind::TextBox(TextBoxData {
                content: blocks,
                padding: text_box_padding,
                vertical_align: text_box_vertical_align,
            }),
        })
    } else if let Some(ref geom) = shape.prst_geom {
        let kind = prst_to_shape_kind(geom, emu_to_pt(shape.cx), emu_to_pt(shape.cy));
        let stroke = shape.ln_color.map(|color| BorderSide {
            width: shape.ln_width_emu as f64 / 12700.0,
            color,
            style: shape.ln_dash_style,
        });
        Some(FixedElement {
            x: emu_to_pt(shape.x),
            y: emu_to_pt(shape.y),
            width: emu_to_pt(shape.cx),
            height: emu_to_pt(shape.cy),
            kind: FixedElementKind::Shape(Shape {
                kind,
                fill: shape.fill,
                gradient_fill: shape.gradient_fill.take(),
                stroke,
                rotation_deg: shape.rotation_deg,
                opacity: shape.opacity,
                shadow: shape.shadow.take(),
            }),
        })
    } else {
        None
    }
}

/// Finalize a picture element when `</p:pic>` is reached.
fn finalize_picture(
    pic: &PictureState,
    images: &SlideImageMap,
    warning_context: &str,
) -> (Option<FixedElement>, Vec<ConvertWarning>) {
    let (selected_asset, picture_warnings) = select_picture_asset(
        images,
        warning_context,
        pic.blip_embed.as_deref(),
        pic.svg_blip_embed.as_deref(),
        &pic.img_layer_embeds,
    );
    let element = selected_asset.and_then(|asset| {
        asset.format().map(|format| FixedElement {
            x: emu_to_pt(pic.x),
            y: emu_to_pt(pic.y),
            width: emu_to_pt(pic.cx),
            height: emu_to_pt(pic.cy),
            kind: FixedElementKind::Image(ImageData {
                data: asset.data.clone(),
                format,
                width: Some(emu_to_pt(pic.cx)),
                height: Some(emu_to_pt(pic.cy)),
                crop: pic.crop,
            }),
        })
    });
    (element, picture_warnings)
}

/// Apply a parsed solid fill color to the appropriate target based on the current context.
fn apply_solid_fill_color(
    ctx: SolidFillCtx,
    parsed: &ParsedColor,
    shape: &mut ShapeState,
    run_style: &mut TextStyle,
    end_run_style: &mut TextStyle,
    bullet_def: &mut PptxBulletDefinition,
) {
    match ctx {
        SolidFillCtx::ShapeFill => {
            shape.fill = parsed.color;
            if let Some(alpha) = parsed.alpha {
                shape.opacity = Some(alpha);
            }
        }
        SolidFillCtx::LineFill => shape.ln_color = parsed.color,
        SolidFillCtx::RunFill => run_style.color = parsed.color,
        SolidFillCtx::EndParaFill => end_run_style.color = parsed.color,
        SolidFillCtx::BulletFill => {
            bullet_def.color = parsed.color.map(PptxBulletColorSource::Explicit);
        }
        SolidFillCtx::None => {}
    }
}

// ── Main parse function ─────────────────────────────────────────────────

/// Parse a slide XML to extract positioned elements (text boxes, shapes, images).
pub(super) fn parse_slide_xml(
    xml: &str,
    images: &SlideImageMap,
    theme: &ThemeData,
    color_map: &ColorMapData,
    warning_context: &str,
    inherited_text_body_defaults: &PptxTextBodyStyleDefaults,
) -> Result<(Vec<FixedElement>, Vec<ConvertWarning>), ConvertError> {
    let mut reader = Reader::from_str(xml);
    let mut elements = Vec::new();
    let mut warnings = Vec::new();

    let mut in_shape = false;
    let mut shape = ShapeState::default();

    let mut in_txbody = false;
    let mut paragraphs: Vec<PptxParagraphEntry> = Vec::new();
    let mut text_box_padding: Insets = default_pptx_text_box_padding();
    let mut text_box_vertical_align: TextBoxVerticalAlign = TextBoxVerticalAlign::Top;
    let mut text_body_style_defaults = PptxTextBodyStyleDefaults::default();

    let mut in_para = false;
    let mut para_style = ParagraphStyle::default();
    let mut para_level: u32 = 0;
    let mut para_default_run_style = TextStyle::default();
    let mut para_end_run_style = TextStyle::default();
    let mut para_bullet_definition = PptxBulletDefinition::default();
    let mut in_ln_spc = false;
    let mut runs: Vec<Run> = Vec::new();

    let mut in_run = false;
    let mut run_style = TextStyle::default();
    let mut run_text = String::new();

    let mut in_text = false;
    let mut in_rpr = false;
    let mut in_end_para_rpr = false;
    let mut solid_fill_ctx = SolidFillCtx::None;

    let mut in_pic = false;
    let mut pic = PictureState::default();

    let mut in_graphic_frame = false;
    let mut gf = GraphicFrameState::default();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"graphicFrame" if !in_shape && !in_pic && !in_graphic_frame => {
                        in_graphic_frame = true;
                        gf.reset();
                    }
                    b"xfrm" if in_graphic_frame && !in_shape => {
                        gf.in_xfrm = true;
                    }
                    b"tbl" if in_graphic_frame => {
                        if let Ok(mut table) = parse_pptx_table(&mut reader, theme, color_map) {
                            scale_pptx_table_geometry_to_frame(
                                &mut table,
                                emu_to_pt(gf.cx),
                                emu_to_pt(gf.cy),
                            );
                            elements.push(FixedElement {
                                x: emu_to_pt(gf.x),
                                y: emu_to_pt(gf.y),
                                width: emu_to_pt(gf.cx),
                                height: emu_to_pt(gf.cy),
                                kind: FixedElementKind::Table(table),
                            });
                        }
                    }
                    b"grpSp" if !in_shape && !in_pic && !in_graphic_frame => {
                        if let Ok((group_elems, group_warnings)) = parse_group_shape(
                            &mut reader,
                            xml,
                            images,
                            theme,
                            color_map,
                            warning_context,
                            inherited_text_body_defaults,
                        ) {
                            elements.extend(group_elems);
                            warnings.extend(group_warnings);
                        }
                    }
                    b"sp" if !in_shape && !in_pic => {
                        in_shape = true;
                        shape.reset();
                        shape.depth = 1;
                        in_txbody = false;
                        paragraphs.clear();
                        text_box_padding = default_pptx_text_box_padding();
                        text_box_vertical_align = TextBoxVerticalAlign::Top;
                    }
                    b"sp" if in_shape => {
                        shape.depth += 1;
                    }
                    b"spPr" if in_shape && !in_txbody => {
                        shape.in_sp_pr = true;
                    }
                    b"xfrm" if in_shape && shape.in_sp_pr => {
                        shape.in_xfrm = true;
                        if let Some(rot) = get_attr_i64(e, b"rot") {
                            shape.rotation_deg = Some(rot as f64 / 60_000.0);
                        }
                    }
                    b"prstGeom" if shape.in_sp_pr => {
                        if let Some(prst) = get_attr_str(e, b"prst") {
                            shape.prst_geom = Some(prst);
                        }
                    }
                    b"solidFill" if shape.in_sp_pr && !shape.in_ln && !in_rpr => {
                        solid_fill_ctx = SolidFillCtx::ShapeFill;
                    }
                    b"gradFill" if shape.in_sp_pr && !shape.in_ln && !in_rpr => {
                        shape.gradient_fill =
                            parse_shape_gradient_fill(&mut reader, theme, color_map);
                        if let Some(ref gradient_fill) = shape.gradient_fill
                            && shape.fill.is_none()
                        {
                            shape.fill = gradient_fill.stops.first().map(|stop| stop.color);
                        }
                    }
                    b"effectLst" if shape.in_sp_pr && !shape.in_ln => {
                        shape.shadow = parse_effect_list(&mut reader, theme, color_map);
                    }
                    b"ln" if shape.in_sp_pr => {
                        shape.in_ln = true;
                        shape.ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                        shape.ln_dash_style = BorderLineStyle::Solid;
                    }
                    b"prstDash" if shape.in_ln => {
                        shape.ln_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }
                    b"solidFill" if shape.in_ln => {
                        solid_fill_ctx = SolidFillCtx::LineFill;
                    }
                    b"ph" if in_shape => {
                        shape.has_placeholder = true;
                    }
                    b"txBody" if in_shape => {
                        in_txbody = true;
                        text_body_style_defaults = if shape.has_placeholder {
                            PptxTextBodyStyleDefaults::default()
                        } else {
                            inherited_text_body_defaults.clone()
                        };
                    }
                    b"bodyPr" if in_shape && in_txbody => {
                        extract_pptx_text_box_body_props(
                            e,
                            &mut text_box_padding,
                            &mut text_box_vertical_align,
                        );
                    }
                    b"lstStyle" if in_shape && in_txbody => {
                        let local_defaults = parse_pptx_list_style(&mut reader, theme, color_map);
                        text_body_style_defaults.merge_from(&local_defaults);
                    }
                    b"p" if in_txbody => {
                        in_para = true;
                        para_level = 0;
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        in_ln_spc = false;
                        runs.clear();
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_ln_spc => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_ctx = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"r" if in_para => {
                        in_run = true;
                        run_style = para_default_run_style.clone();
                        run_text.clear();
                    }
                    b"rPr" if in_run => {
                        in_rpr = true;
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        in_end_para_rpr = true;
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"solidFill" if in_rpr => {
                        solid_fill_ctx = SolidFillCtx::RunFill;
                    }
                    b"solidFill" if in_end_para_rpr => {
                        solid_fill_ctx = SolidFillCtx::EndParaFill;
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_ctx != SolidFillCtx::None =>
                    {
                        let parsed = parse_color_from_start(&mut reader, e, theme, color_map);
                        apply_solid_fill_color(
                            solid_fill_ctx,
                            &parsed,
                            &mut shape,
                            &mut run_style,
                            &mut para_end_run_style,
                            &mut para_bullet_definition,
                        );
                    }
                    b"t" if in_run => {
                        in_text = true;
                    }
                    b"pic" if !in_shape && !in_pic => {
                        in_pic = true;
                        pic.reset();
                    }
                    b"spPr" if in_pic => {}
                    b"xfrm" if in_pic => {
                        pic.in_xfrm = true;
                    }
                    b"blipFill" if in_pic => {}
                    b"blip" if in_pic => {
                        pic.blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"svgBlip" if in_pic => {
                        pic.svg_blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"imgLayer" if in_pic => {
                        if let Some(rid) = get_attr_str(e, b"r:embed") {
                            pic.img_layer_embeds.push(rid);
                        }
                    }
                    b"srcRect" if in_pic => {
                        pic.crop = parse_src_rect(e);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"off" if shape.in_xfrm => {
                        shape.x = get_attr_i64(e, b"x").unwrap_or(0);
                        shape.y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if shape.in_xfrm => {
                        shape.cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        shape.cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }
                    b"off" if pic.in_xfrm => {
                        pic.x = get_attr_i64(e, b"x").unwrap_or(0);
                        pic.y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if pic.in_xfrm => {
                        pic.cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        pic.cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }
                    b"off" if gf.in_xfrm => {
                        gf.x = get_attr_i64(e, b"x").unwrap_or(0);
                        gf.y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if gf.in_xfrm => {
                        gf.cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        gf.cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }
                    b"blip" if in_pic => {
                        pic.blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"svgBlip" if in_pic => {
                        pic.svg_blip_embed = get_attr_str(e, b"r:embed");
                    }
                    b"imgLayer" if in_pic => {
                        if let Some(rid) = get_attr_str(e, b"r:embed") {
                            pic.img_layer_embeds.push(rid);
                        }
                    }
                    b"srcRect" if in_pic => {
                        pic.crop = parse_src_rect(e);
                    }
                    b"prstGeom" if shape.in_sp_pr => {
                        if let Some(prst) = get_attr_str(e, b"prst") {
                            shape.prst_geom = Some(prst);
                        }
                    }
                    b"ln" if shape.in_sp_pr => {
                        shape.ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                    }
                    b"prstDash" if shape.in_ln => {
                        shape.ln_dash_style = get_attr_str(e, b"val")
                            .as_deref()
                            .map(pptx_dash_to_border_style)
                            .unwrap_or(BorderLineStyle::Solid);
                    }
                    b"srgbClr" | b"schemeClr" | b"sysClr"
                        if solid_fill_ctx != SolidFillCtx::None =>
                    {
                        let parsed = parse_color_from_empty(e, theme, color_map);
                        apply_solid_fill_color(
                            solid_fill_ctx,
                            &parsed,
                            &mut shape,
                            &mut run_style,
                            &mut para_end_run_style,
                            &mut para_bullet_definition,
                        );
                    }
                    b"rPr" if in_run => {
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"endParaRPr" if in_para && !in_run => {
                        para_end_run_style = para_default_run_style.clone();
                        extract_rpr_attributes(e, &mut para_end_run_style);
                    }
                    b"pPr" if in_para && !in_run => {
                        para_level = extract_paragraph_level(e);
                        para_style = text_body_style_defaults.paragraph_style_for_level(para_level);
                        para_default_run_style =
                            text_body_style_defaults.run_style_for_level(para_level);
                        para_end_run_style = para_default_run_style.clone();
                        para_bullet_definition =
                            text_body_style_defaults.bullet_for_level(para_level);
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"lnSpc" if in_para && !in_run => {
                        in_ln_spc = true;
                    }
                    b"spcPct" if in_ln_spc => {
                        extract_pptx_line_spacing_pct(e, &mut para_style);
                    }
                    b"spcPts" if in_ln_spc => {
                        extract_pptx_line_spacing_pts(e, &mut para_style);
                    }
                    b"buAutoNum" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                            parse_pptx_auto_numbering(e, para_level),
                        ));
                    }
                    b"buChar" if in_para && !in_run => {
                        para_bullet_definition.kind = parse_pptx_bullet_marker(e, para_level);
                    }
                    b"buNone" if in_para && !in_run => {
                        para_bullet_definition.kind = Some(PptxBulletKind::None);
                    }
                    b"buFontTx" if in_para && !in_run => {
                        para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
                    }
                    b"buFont" if in_para && !in_run => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                                resolve_theme_font(&typeface, theme),
                            ));
                        }
                    }
                    b"buClrTx" if in_para && !in_run => {
                        para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
                    }
                    b"buClr" if in_para && !in_run => {
                        solid_fill_ctx = SolidFillCtx::BulletFill;
                    }
                    b"buSzTx" if in_para && !in_run => {
                        para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
                    }
                    b"buSzPct" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                        }
                    }
                    b"buSzPts" if in_para && !in_run => {
                        if let Some(val) = get_attr_i64(e, b"val") {
                            para_bullet_definition.size =
                                Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                        }
                    }
                    b"br" if in_para && !in_run => {
                        push_pptx_soft_line_break(&mut runs, &para_default_run_style);
                    }
                    b"latin" | b"ea" | b"cs" if in_rpr => {
                        apply_typeface_to_style(e, &mut run_style, theme);
                    }
                    b"latin" | b"ea" | b"cs" if in_end_para_rpr => {
                        apply_typeface_to_style(e, &mut para_end_run_style, theme);
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_text && let Some(text) = decode_pptx_text_event(t) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::GeneralRef(ref reference)) => {
                if in_text && let Some(text) = decode_pptx_general_ref(reference) {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"sp" if in_shape => {
                        shape.depth -= 1;
                        if shape.depth == 0 {
                            if let Some(element) = finalize_shape(
                                &mut shape,
                                &mut paragraphs,
                                text_box_padding,
                                text_box_vertical_align,
                            ) {
                                elements.push(element);
                            }
                            in_shape = false;
                        }
                    }
                    b"spPr" if shape.in_sp_pr => {
                        shape.in_sp_pr = false;
                    }
                    b"xfrm" if shape.in_xfrm => {
                        shape.in_xfrm = false;
                    }
                    b"ln" if shape.in_ln => {
                        shape.in_ln = false;
                    }
                    b"txBody" if in_txbody => {
                        in_txbody = false;
                    }
                    b"p" if in_para => {
                        let resolved_list_marker = resolve_pptx_list_marker(
                            &para_bullet_definition,
                            para_level,
                            &runs,
                            &para_end_run_style,
                            &para_default_run_style,
                        );
                        let paragraph_runs = std::mem::take(&mut runs);
                        paragraphs.push(PptxParagraphEntry {
                            paragraph: Paragraph {
                                style: para_style.clone(),
                                runs: paragraph_runs,
                            },
                            list_marker: resolved_list_marker,
                        });
                        in_para = false;
                    }
                    b"r" if in_run => {
                        if !run_text.is_empty() {
                            push_pptx_run(
                                &mut runs,
                                Run {
                                    text: std::mem::take(&mut run_text),
                                    style: run_style.clone(),
                                    href: None,
                                    footnote: None,
                                },
                            );
                        }
                        in_run = false;
                    }
                    b"rPr" if in_rpr => {
                        in_rpr = false;
                    }
                    b"endParaRPr" if in_end_para_rpr => {
                        in_end_para_rpr = false;
                    }
                    b"lnSpc" if in_ln_spc => {
                        in_ln_spc = false;
                    }
                    b"solidFill" if solid_fill_ctx != SolidFillCtx::None => {
                        solid_fill_ctx = SolidFillCtx::None;
                    }
                    b"t" if in_text => {
                        in_text = false;
                    }
                    b"pic" if in_pic => {
                        let (element, picture_warnings) =
                            finalize_picture(&pic, images, warning_context);
                        warnings.extend(picture_warnings);
                        if let Some(element) = element {
                            elements.push(element);
                        }
                        in_pic = false;
                    }
                    b"xfrm" if pic.in_xfrm => {
                        pic.in_xfrm = false;
                    }
                    b"graphicFrame" if in_graphic_frame => {
                        in_graphic_frame = false;
                    }
                    b"xfrm" if gf.in_xfrm => {
                        gf.in_xfrm = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(ConvertError::Parse(format!("XML error in slide: {error}"))),
            _ => {}
        }
    }

    Ok((elements, warnings))
}
