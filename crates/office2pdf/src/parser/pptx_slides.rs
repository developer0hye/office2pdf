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
            width: emu_to_pt(shape.ln_width_emu),
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

// ── SlideXmlParser state machine ────────────────────────────────────────

/// Bundles the 20+ mutable state variables of the slide XML event loop
/// into a single struct, with methods for each event type.
///
/// The XML reader is passed to each handler rather than stored, because
/// several sub-parsers (`parse_pptx_table`, `parse_group_shape`, etc.)
/// need `&mut Reader` to consume nested elements.
struct SlideXmlParser<'a> {
    // ── Context references (immutable for the parse lifetime) ────────
    xml: &'a str,
    images: &'a SlideImageMap,
    theme: &'a ThemeData,
    color_map: &'a ColorMapData,
    warning_context: &'a str,
    inherited_text_body_defaults: &'a PptxTextBodyStyleDefaults,

    // ── Output accumulators ─────────────────────────────────────────
    elements: Vec<FixedElement>,
    warnings: Vec<ConvertWarning>,

    // ── Shape state (`<p:sp>`) ──────────────────────────────────────
    in_shape: bool,
    shape: ShapeState,

    // ── Text body state (`<p:txBody>`) ──────────────────────────────
    in_txbody: bool,
    paragraphs: Vec<PptxParagraphEntry>,
    text_box_padding: Insets,
    text_box_vertical_align: TextBoxVerticalAlign,
    text_body_style_defaults: PptxTextBodyStyleDefaults,

    // ── Paragraph state (`<a:p>`) ───────────────────────────────────
    in_para: bool,
    para_style: ParagraphStyle,
    para_level: u32,
    para_default_run_style: TextStyle,
    para_end_run_style: TextStyle,
    para_bullet_definition: PptxBulletDefinition,
    in_ln_spc: bool,
    runs: Vec<Run>,

    // ── Run state (`<a:r>`) ─────────────────────────────────────────
    in_run: bool,
    run_style: TextStyle,
    run_text: String,

    // ── Inline tracking flags ───────────────────────────────────────
    in_text: bool,
    in_rpr: bool,
    in_end_para_rpr: bool,
    solid_fill_ctx: SolidFillCtx,

    // ── Picture state (`<p:pic>`) ───────────────────────────────────
    in_pic: bool,
    pic: PictureState,

    // ── Graphic frame state (`<p:graphicFrame>`) ────────────────────
    in_graphic_frame: bool,
    gf: GraphicFrameState,
}

impl<'a> SlideXmlParser<'a> {
    fn new(
        xml: &'a str,
        images: &'a SlideImageMap,
        theme: &'a ThemeData,
        color_map: &'a ColorMapData,
        warning_context: &'a str,
        inherited_text_body_defaults: &'a PptxTextBodyStyleDefaults,
    ) -> Self {
        Self {
            xml,
            images,
            theme,
            color_map,
            warning_context,
            inherited_text_body_defaults,

            elements: Vec::new(),
            warnings: Vec::new(),

            in_shape: false,
            shape: ShapeState::default(),

            in_txbody: false,
            paragraphs: Vec::new(),
            text_box_padding: default_pptx_text_box_padding(),
            text_box_vertical_align: TextBoxVerticalAlign::Top,
            text_body_style_defaults: PptxTextBodyStyleDefaults::default(),

            in_para: false,
            para_style: ParagraphStyle::default(),
            para_level: 0,
            para_default_run_style: TextStyle::default(),
            para_end_run_style: TextStyle::default(),
            para_bullet_definition: PptxBulletDefinition::default(),
            in_ln_spc: false,
            runs: Vec::new(),

            in_run: false,
            run_style: TextStyle::default(),
            run_text: String::new(),

            in_text: false,
            in_rpr: false,
            in_end_para_rpr: false,
            solid_fill_ctx: SolidFillCtx::None,

            in_pic: false,
            pic: PictureState::default(),

            in_graphic_frame: false,
            gf: GraphicFrameState::default(),
        }
    }

    /// Handle an `Event::Start` element.
    fn handle_start(&mut self, reader: &mut Reader<&[u8]>, e: &BytesStart<'_>) {
        let local = e.local_name();
        match local.as_ref() {
            b"graphicFrame" if !self.in_shape && !self.in_pic && !self.in_graphic_frame => {
                self.in_graphic_frame = true;
                self.gf.reset();
            }
            b"xfrm" if self.in_graphic_frame && !self.in_shape => {
                self.gf.in_xfrm = true;
            }
            b"tbl" if self.in_graphic_frame => {
                if let Ok(mut table) = parse_pptx_table(reader, self.theme, self.color_map) {
                    scale_pptx_table_geometry_to_frame(
                        &mut table,
                        emu_to_pt(self.gf.cx),
                        emu_to_pt(self.gf.cy),
                    );
                    self.elements.push(FixedElement {
                        x: emu_to_pt(self.gf.x),
                        y: emu_to_pt(self.gf.y),
                        width: emu_to_pt(self.gf.cx),
                        height: emu_to_pt(self.gf.cy),
                        kind: FixedElementKind::Table(table),
                    });
                }
            }
            b"grpSp" if !self.in_shape && !self.in_pic && !self.in_graphic_frame => {
                if let Ok((group_elems, group_warnings)) = parse_group_shape(
                    reader,
                    self.xml,
                    self.images,
                    self.theme,
                    self.color_map,
                    self.warning_context,
                    self.inherited_text_body_defaults,
                ) {
                    self.elements.extend(group_elems);
                    self.warnings.extend(group_warnings);
                }
            }
            b"sp" if !self.in_shape && !self.in_pic => {
                self.in_shape = true;
                self.shape.reset();
                self.shape.depth = 1;
                self.in_txbody = false;
                self.paragraphs.clear();
                self.text_box_padding = default_pptx_text_box_padding();
                self.text_box_vertical_align = TextBoxVerticalAlign::Top;
            }
            b"sp" if self.in_shape => {
                self.shape.depth += 1;
            }
            b"spPr" if self.in_shape && !self.in_txbody => {
                self.shape.in_sp_pr = true;
            }
            b"xfrm" if self.in_shape && self.shape.in_sp_pr => {
                self.shape.in_xfrm = true;
                if let Some(rot) = get_attr_i64(e, b"rot") {
                    self.shape.rotation_deg = Some(rot as f64 / 60_000.0);
                }
            }
            b"prstGeom" if self.shape.in_sp_pr => {
                if let Some(prst) = get_attr_str(e, b"prst") {
                    self.shape.prst_geom = Some(prst);
                }
            }
            b"solidFill" if self.shape.in_sp_pr && !self.shape.in_ln && !self.in_rpr => {
                self.solid_fill_ctx = SolidFillCtx::ShapeFill;
            }
            b"gradFill" if self.shape.in_sp_pr && !self.shape.in_ln && !self.in_rpr => {
                self.shape.gradient_fill =
                    parse_shape_gradient_fill(reader, self.theme, self.color_map);
                if let Some(ref gradient_fill) = self.shape.gradient_fill
                    && self.shape.fill.is_none()
                {
                    self.shape.fill = gradient_fill.stops.first().map(|stop| stop.color);
                }
            }
            b"effectLst" if self.shape.in_sp_pr && !self.shape.in_ln => {
                self.shape.shadow = parse_effect_list(reader, self.theme, self.color_map);
            }
            b"ln" if self.shape.in_sp_pr => {
                self.shape.in_ln = true;
                self.shape.ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                self.shape.ln_dash_style = BorderLineStyle::Solid;
            }
            b"prstDash" if self.shape.in_ln => {
                self.shape.ln_dash_style = get_attr_str(e, b"val")
                    .as_deref()
                    .map(pptx_dash_to_border_style)
                    .unwrap_or(BorderLineStyle::Solid);
            }
            b"solidFill" if self.shape.in_ln => {
                self.solid_fill_ctx = SolidFillCtx::LineFill;
            }
            b"ph" if self.in_shape => {
                self.shape.has_placeholder = true;
            }
            b"txBody" if self.in_shape => {
                self.in_txbody = true;
                self.text_body_style_defaults = if self.shape.has_placeholder {
                    PptxTextBodyStyleDefaults::default()
                } else {
                    self.inherited_text_body_defaults.clone()
                };
            }
            b"bodyPr" if self.in_shape && self.in_txbody => {
                extract_pptx_text_box_body_props(
                    e,
                    &mut self.text_box_padding,
                    &mut self.text_box_vertical_align,
                );
            }
            b"lstStyle" if self.in_shape && self.in_txbody => {
                let local_defaults = parse_pptx_list_style(reader, self.theme, self.color_map);
                self.text_body_style_defaults.merge_from(&local_defaults);
            }
            b"p" if self.in_txbody => {
                self.in_para = true;
                self.para_level = 0;
                self.para_style = self
                    .text_body_style_defaults
                    .paragraph_style_for_level(self.para_level);
                self.para_default_run_style = self
                    .text_body_style_defaults
                    .run_style_for_level(self.para_level);
                self.para_end_run_style = self.para_default_run_style.clone();
                self.para_bullet_definition = self
                    .text_body_style_defaults
                    .bullet_for_level(self.para_level);
                self.in_ln_spc = false;
                self.runs.clear();
            }
            b"pPr" if self.in_para && !self.in_run => {
                self.para_level = extract_paragraph_level(e);
                self.para_style = self
                    .text_body_style_defaults
                    .paragraph_style_for_level(self.para_level);
                self.para_default_run_style = self
                    .text_body_style_defaults
                    .run_style_for_level(self.para_level);
                self.para_end_run_style = self.para_default_run_style.clone();
                self.para_bullet_definition = self
                    .text_body_style_defaults
                    .bullet_for_level(self.para_level);
                extract_paragraph_props(e, &mut self.para_style);
            }
            b"lnSpc" if self.in_para && !self.in_run => {
                self.in_ln_spc = true;
            }
            b"spcPct" if self.in_ln_spc => {
                extract_pptx_line_spacing_pct(e, &mut self.para_style);
            }
            b"spcPts" if self.in_ln_spc => {
                extract_pptx_line_spacing_pts(e, &mut self.para_style);
            }
            b"buAutoNum" if self.in_para && !self.in_run => {
                self.para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                    parse_pptx_auto_numbering(e, self.para_level),
                ));
            }
            b"buChar" if self.in_para && !self.in_run => {
                self.para_bullet_definition.kind = parse_pptx_bullet_marker(e, self.para_level);
            }
            b"buNone" if self.in_para && !self.in_run => {
                self.para_bullet_definition.kind = Some(PptxBulletKind::None);
            }
            b"buFontTx" if self.in_para && !self.in_run => {
                self.para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
            }
            b"buFont" if self.in_para && !self.in_run => {
                if let Some(typeface) = get_attr_str(e, b"typeface") {
                    self.para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                        resolve_theme_font(&typeface, self.theme),
                    ));
                }
            }
            b"buClrTx" if self.in_para && !self.in_run => {
                self.para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
            }
            b"buClr" if self.in_para && !self.in_run => {
                self.solid_fill_ctx = SolidFillCtx::BulletFill;
            }
            b"buSzTx" if self.in_para && !self.in_run => {
                self.para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
            }
            b"buSzPct" if self.in_para && !self.in_run => {
                if let Some(val) = get_attr_i64(e, b"val") {
                    self.para_bullet_definition.size =
                        Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                }
            }
            b"buSzPts" if self.in_para && !self.in_run => {
                if let Some(val) = get_attr_i64(e, b"val") {
                    self.para_bullet_definition.size =
                        Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                }
            }
            b"br" if self.in_para && !self.in_run => {
                push_pptx_soft_line_break(&mut self.runs, &self.para_default_run_style);
            }
            b"r" if self.in_para => {
                self.in_run = true;
                self.run_style = self.para_default_run_style.clone();
                self.run_text.clear();
            }
            b"rPr" if self.in_run => {
                self.in_rpr = true;
                extract_rpr_attributes(e, &mut self.run_style);
            }
            b"endParaRPr" if self.in_para && !self.in_run => {
                self.in_end_para_rpr = true;
                self.para_end_run_style = self.para_default_run_style.clone();
                extract_rpr_attributes(e, &mut self.para_end_run_style);
            }
            b"solidFill" if self.in_rpr => {
                self.solid_fill_ctx = SolidFillCtx::RunFill;
            }
            b"solidFill" if self.in_end_para_rpr => {
                self.solid_fill_ctx = SolidFillCtx::EndParaFill;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.solid_fill_ctx != SolidFillCtx::None => {
                let parsed = parse_color_from_start(reader, e, self.theme, self.color_map);
                apply_solid_fill_color(
                    self.solid_fill_ctx,
                    &parsed,
                    &mut self.shape,
                    &mut self.run_style,
                    &mut self.para_end_run_style,
                    &mut self.para_bullet_definition,
                );
            }
            b"t" if self.in_run => {
                self.in_text = true;
            }
            b"pic" if !self.in_shape && !self.in_pic => {
                self.in_pic = true;
                self.pic.reset();
            }
            b"spPr" if self.in_pic => {}
            b"xfrm" if self.in_pic => {
                self.pic.in_xfrm = true;
            }
            b"blipFill" if self.in_pic => {}
            b"blip" if self.in_pic => {
                self.pic.blip_embed = get_attr_str(e, b"r:embed");
            }
            b"svgBlip" if self.in_pic => {
                self.pic.svg_blip_embed = get_attr_str(e, b"r:embed");
            }
            b"imgLayer" if self.in_pic => {
                if let Some(rid) = get_attr_str(e, b"r:embed") {
                    self.pic.img_layer_embeds.push(rid);
                }
            }
            b"srcRect" if self.in_pic => {
                self.pic.crop = parse_src_rect(e);
            }
            _ => {}
        }
    }

    /// Handle an `Event::Empty` element.
    fn handle_empty(&mut self, e: &BytesStart<'_>) {
        let local = e.local_name();
        match local.as_ref() {
            b"off" if self.shape.in_xfrm => {
                self.shape.x = get_attr_i64(e, b"x").unwrap_or(0);
                self.shape.y = get_attr_i64(e, b"y").unwrap_or(0);
            }
            b"ext" if self.shape.in_xfrm => {
                self.shape.cx = get_attr_i64(e, b"cx").unwrap_or(0);
                self.shape.cy = get_attr_i64(e, b"cy").unwrap_or(0);
            }
            b"off" if self.pic.in_xfrm => {
                self.pic.x = get_attr_i64(e, b"x").unwrap_or(0);
                self.pic.y = get_attr_i64(e, b"y").unwrap_or(0);
            }
            b"ext" if self.pic.in_xfrm => {
                self.pic.cx = get_attr_i64(e, b"cx").unwrap_or(0);
                self.pic.cy = get_attr_i64(e, b"cy").unwrap_or(0);
            }
            b"off" if self.gf.in_xfrm => {
                self.gf.x = get_attr_i64(e, b"x").unwrap_or(0);
                self.gf.y = get_attr_i64(e, b"y").unwrap_or(0);
            }
            b"ext" if self.gf.in_xfrm => {
                self.gf.cx = get_attr_i64(e, b"cx").unwrap_or(0);
                self.gf.cy = get_attr_i64(e, b"cy").unwrap_or(0);
            }
            b"blip" if self.in_pic => {
                self.pic.blip_embed = get_attr_str(e, b"r:embed");
            }
            b"svgBlip" if self.in_pic => {
                self.pic.svg_blip_embed = get_attr_str(e, b"r:embed");
            }
            b"imgLayer" if self.in_pic => {
                if let Some(rid) = get_attr_str(e, b"r:embed") {
                    self.pic.img_layer_embeds.push(rid);
                }
            }
            b"srcRect" if self.in_pic => {
                self.pic.crop = parse_src_rect(e);
            }
            b"prstGeom" if self.shape.in_sp_pr => {
                if let Some(prst) = get_attr_str(e, b"prst") {
                    self.shape.prst_geom = Some(prst);
                }
            }
            b"ln" if self.shape.in_sp_pr => {
                self.shape.ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
            }
            b"prstDash" if self.shape.in_ln => {
                self.shape.ln_dash_style = get_attr_str(e, b"val")
                    .as_deref()
                    .map(pptx_dash_to_border_style)
                    .unwrap_or(BorderLineStyle::Solid);
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.solid_fill_ctx != SolidFillCtx::None => {
                let parsed = parse_color_from_empty(e, self.theme, self.color_map);
                apply_solid_fill_color(
                    self.solid_fill_ctx,
                    &parsed,
                    &mut self.shape,
                    &mut self.run_style,
                    &mut self.para_end_run_style,
                    &mut self.para_bullet_definition,
                );
            }
            b"rPr" if self.in_run => {
                extract_rpr_attributes(e, &mut self.run_style);
            }
            b"endParaRPr" if self.in_para && !self.in_run => {
                self.para_end_run_style = self.para_default_run_style.clone();
                extract_rpr_attributes(e, &mut self.para_end_run_style);
            }
            b"pPr" if self.in_para && !self.in_run => {
                self.para_level = extract_paragraph_level(e);
                self.para_style = self
                    .text_body_style_defaults
                    .paragraph_style_for_level(self.para_level);
                self.para_default_run_style = self
                    .text_body_style_defaults
                    .run_style_for_level(self.para_level);
                self.para_end_run_style = self.para_default_run_style.clone();
                self.para_bullet_definition = self
                    .text_body_style_defaults
                    .bullet_for_level(self.para_level);
                extract_paragraph_props(e, &mut self.para_style);
            }
            b"lnSpc" if self.in_para && !self.in_run => {
                self.in_ln_spc = true;
            }
            b"spcPct" if self.in_ln_spc => {
                extract_pptx_line_spacing_pct(e, &mut self.para_style);
            }
            b"spcPts" if self.in_ln_spc => {
                extract_pptx_line_spacing_pts(e, &mut self.para_style);
            }
            b"buAutoNum" if self.in_para && !self.in_run => {
                self.para_bullet_definition.kind = Some(PptxBulletKind::AutoNumber(
                    parse_pptx_auto_numbering(e, self.para_level),
                ));
            }
            b"buChar" if self.in_para && !self.in_run => {
                self.para_bullet_definition.kind = parse_pptx_bullet_marker(e, self.para_level);
            }
            b"buNone" if self.in_para && !self.in_run => {
                self.para_bullet_definition.kind = Some(PptxBulletKind::None);
            }
            b"buFontTx" if self.in_para && !self.in_run => {
                self.para_bullet_definition.font = Some(PptxBulletFontSource::FollowText);
            }
            b"buFont" if self.in_para && !self.in_run => {
                if let Some(typeface) = get_attr_str(e, b"typeface") {
                    self.para_bullet_definition.font = Some(PptxBulletFontSource::Explicit(
                        resolve_theme_font(&typeface, self.theme),
                    ));
                }
            }
            b"buClrTx" if self.in_para && !self.in_run => {
                self.para_bullet_definition.color = Some(PptxBulletColorSource::FollowText);
            }
            b"buClr" if self.in_para && !self.in_run => {
                self.solid_fill_ctx = SolidFillCtx::BulletFill;
            }
            b"buSzTx" if self.in_para && !self.in_run => {
                self.para_bullet_definition.size = Some(PptxBulletSizeSource::FollowText);
            }
            b"buSzPct" if self.in_para && !self.in_run => {
                if let Some(val) = get_attr_i64(e, b"val") {
                    self.para_bullet_definition.size =
                        Some(PptxBulletSizeSource::Percent(val as f64 / 100_000.0));
                }
            }
            b"buSzPts" if self.in_para && !self.in_run => {
                if let Some(val) = get_attr_i64(e, b"val") {
                    self.para_bullet_definition.size =
                        Some(PptxBulletSizeSource::Points(val as f64 / 100.0));
                }
            }
            b"br" if self.in_para && !self.in_run => {
                push_pptx_soft_line_break(&mut self.runs, &self.para_default_run_style);
            }
            b"latin" | b"ea" | b"cs" if self.in_rpr => {
                apply_typeface_to_style(e, &mut self.run_style, self.theme);
            }
            b"latin" | b"ea" | b"cs" if self.in_end_para_rpr => {
                apply_typeface_to_style(e, &mut self.para_end_run_style, self.theme);
            }
            _ => {}
        }
    }

    /// Handle an `Event::Text` element.
    fn handle_text(&mut self, text: &str) {
        if self.in_text {
            self.run_text.push_str(text);
        }
    }

    /// Handle an `Event::End` element.
    fn handle_end(&mut self, local_name: &[u8]) {
        match local_name {
            b"sp" if self.in_shape => {
                self.shape.depth -= 1;
                if self.shape.depth == 0 {
                    if let Some(element) = finalize_shape(
                        &mut self.shape,
                        &mut self.paragraphs,
                        self.text_box_padding,
                        self.text_box_vertical_align,
                    ) {
                        self.elements.push(element);
                    }
                    self.in_shape = false;
                }
            }
            b"spPr" if self.shape.in_sp_pr => {
                self.shape.in_sp_pr = false;
            }
            b"xfrm" if self.shape.in_xfrm => {
                self.shape.in_xfrm = false;
            }
            b"ln" if self.shape.in_ln => {
                self.shape.in_ln = false;
            }
            b"txBody" if self.in_txbody => {
                self.in_txbody = false;
            }
            b"p" if self.in_para => {
                let resolved_list_marker = resolve_pptx_list_marker(
                    &self.para_bullet_definition,
                    self.para_level,
                    &self.runs,
                    &self.para_end_run_style,
                    &self.para_default_run_style,
                );
                let paragraph_runs = std::mem::take(&mut self.runs);
                self.paragraphs.push(PptxParagraphEntry {
                    paragraph: Paragraph {
                        style: self.para_style.clone(),
                        runs: paragraph_runs,
                    },
                    list_marker: resolved_list_marker,
                });
                self.in_para = false;
            }
            b"r" if self.in_run => {
                if !self.run_text.is_empty() {
                    push_pptx_run(
                        &mut self.runs,
                        Run {
                            text: std::mem::take(&mut self.run_text),
                            style: self.run_style.clone(),
                            href: None,
                            footnote: None,
                        },
                    );
                }
                self.in_run = false;
            }
            b"rPr" if self.in_rpr => {
                self.in_rpr = false;
            }
            b"endParaRPr" if self.in_end_para_rpr => {
                self.in_end_para_rpr = false;
            }
            b"lnSpc" if self.in_ln_spc => {
                self.in_ln_spc = false;
            }
            b"solidFill" if self.solid_fill_ctx != SolidFillCtx::None => {
                self.solid_fill_ctx = SolidFillCtx::None;
            }
            b"t" if self.in_text => {
                self.in_text = false;
            }
            b"pic" if self.in_pic => {
                let (element, picture_warnings) =
                    finalize_picture(&self.pic, self.images, self.warning_context);
                self.warnings.extend(picture_warnings);
                if let Some(element) = element {
                    self.elements.push(element);
                }
                self.in_pic = false;
            }
            b"xfrm" if self.pic.in_xfrm => {
                self.pic.in_xfrm = false;
            }
            b"graphicFrame" if self.in_graphic_frame => {
                self.in_graphic_frame = false;
            }
            b"xfrm" if self.gf.in_xfrm => {
                self.gf.in_xfrm = false;
            }
            _ => {}
        }
    }

    /// Consume the parser and return the accumulated results.
    fn finish(self) -> (Vec<FixedElement>, Vec<ConvertWarning>) {
        (self.elements, self.warnings)
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
    let mut parser = SlideXmlParser::new(
        xml,
        images,
        theme,
        color_map,
        warning_context,
        inherited_text_body_defaults,
    );

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                parser.handle_start(&mut reader, e);
            }
            Ok(Event::Empty(ref e)) => {
                parser.handle_empty(e);
            }
            Ok(Event::Text(ref t)) => {
                if let Some(text) = decode_pptx_text_event(t) {
                    parser.handle_text(&text);
                }
            }
            Ok(Event::GeneralRef(ref reference)) => {
                if let Some(text) = decode_pptx_general_ref(reference) {
                    parser.handle_text(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                parser.handle_end(e.local_name().as_ref());
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(crate::parser::parse_err(format!(
                    "XML error in slide: {error}"
                )));
            }
            _ => {}
        }
    }

    Ok(parser.finish())
}
