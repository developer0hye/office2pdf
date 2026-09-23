use super::package::{
    load_chart_data, load_slide_images, load_smartart_data, parse_rels_xml, rels_path_for,
    resolve_layout_master_paths, resolve_relative_path, scan_chart_refs,
};
use super::placeholders::PlaceholderGeometryMap;
use super::*;
use crate::parser::xml_util::OOXML_XML_VERSION;

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
    master_text_styles: PptxMasterTextStyles,
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
    let master_text_styles: PptxMasterTextStyles = master_xml
        .as_deref()
        .map(|xml| parse_master_text_styles(xml, theme, &master_color_map))
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
        master_text_styles,
    })
}

/// One inheritance layer (master or layout) to pull elements from.
struct SlideLayer<'a> {
    path: &'a str,
    xml: &'a str,
    color_map: &'a ColorMapData,
    label: &'a str,
    text_style_defaults: &'a PptxTextBodyStyleDefaults,
}

/// Parse elements from a single inheritance layer (master or layout).
/// Broken layers are non-fatal and silently return empty results.
fn parse_layer_elements<R: Read + std::io::Seek>(
    layer: SlideLayer<'_>,
    theme: &ThemeData,
    slide_number: u32,
    default_text_size_pt: Option<f64>,
    archive: &mut ZipArchive<R>,
) -> (Vec<FixedElement>, Vec<ConvertWarning>) {
    let images: SlideImageMap = load_slide_images(layer.path, archive);
    let empty_table_styles: table_styles::TableStyleMap = table_styles::TableStyleMap::new();
    let ctx = SlideParseContext {
        images: &images,
        slide_number,
        theme,
        color_map: layer.color_map,
        warning_context: layer.label,
        inherited_text_body_defaults: layer.text_style_defaults,
        table_styles: &empty_table_styles,
        default_text_size_pt,
    };
    // Skip placeholder shapes in master/layout layers.
    parse_slide_xml_inner(layer.xml, &ctx, true, None).unwrap_or_default()
}

// ── Embedded object helpers ─────────────────────────────────────────────

/// Collect SmartArt elements referenced by the slide XML.
fn collect_smartart_elements<R: Read + std::io::Seek>(
    slide_xml: &str,
    slide_path: &str,
    archive: &mut ZipArchive<R>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Vec<FixedElement> {
    let smartart_refs = smartart::scan_smartart_refs(slide_xml);
    if smartart_refs.is_empty() {
        return Vec::new();
    }

    let smartart_data = load_smartart_data(slide_path, archive);
    let mut elements: Vec<FixedElement> = Vec::new();
    for sa_ref in &smartart_refs {
        // Prefer the pre-rendered drawing cache (the real shapes PowerPoint
        // laid out); fall back to a structured node list when absent.
        let drawing_elems: Vec<FixedElement> =
            load_smartart_drawing_xml(slide_path, archive, &sa_ref.data_rid)
                .map(|xml| {
                    parse_smartart_drawing(
                        &xml,
                        theme,
                        color_map,
                        emu_to_pt(sa_ref.x),
                        emu_to_pt(sa_ref.y),
                    )
                })
                .unwrap_or_default();
        if !drawing_elems.is_empty() {
            elements.extend(drawing_elems);
        } else if let Some(items) = smartart_data.get(&sa_ref.data_rid) {
            elements.push(FixedElement {
                x: emu_to_pt(sa_ref.x),
                y: emu_to_pt(sa_ref.y),
                width: emu_to_pt(sa_ref.cx),
                height: emu_to_pt(sa_ref.cy),
                kind: FixedElementKind::SmartArt(SmartArt {
                    items: items.clone(),
                }),
            });
        }
    }
    elements
}

/// Resolve the SmartArt drawing cache (`diagrams/drawingN.xml`) for a
/// diagram: slide rels(data_rid) → data XML → dataModelExt relId → slide
/// rels(drawing_rid) → drawing XML.
fn load_smartart_drawing_xml<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
    data_rid: &str,
) -> Option<String> {
    let rels_xml: String = read_zip_entry(archive, &rels_path_for(slide_path)).ok()?;
    let rels: HashMap<String, String> = parse_rels_xml(&rels_xml);
    let slide_dir: &str = slide_path
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");

    let data_target: &str = rels.get(data_rid)?;
    let data_path: String = match data_target.strip_prefix('/') {
        Some(stripped) => stripped.to_string(),
        None => resolve_relative_path(slide_dir, data_target),
    };
    let data_xml: String = read_zip_entry(archive, &data_path).ok()?;

    // <dsp:dataModelExt relId="rIdN"> names the drawing relationship (in the
    // slide's rels, not the data part's).
    let drawing_rid: String = extract_data_model_ext_rel_id(&data_xml)?;
    let drawing_target: &str = rels.get(&drawing_rid)?;
    let drawing_path: String = match drawing_target.strip_prefix('/') {
        Some(stripped) => stripped.to_string(),
        None => resolve_relative_path(slide_dir, drawing_target),
    };
    read_zip_entry(archive, &drawing_path).ok()
}

fn extract_data_model_ext_rel_id(data_xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(data_xml);
    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"dataModelExt" =>
            {
                return e.attributes().flatten().find_map(|attr| {
                    (attr.key.local_name().as_ref() == b"relId")
                        .then(|| attr.normalized_value(OOXML_XML_VERSION).ok())
                        .flatten()
                        .map(|v| v.to_string())
                });
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
    }
}

/// Parse the SmartArt drawing cache's `<dsp:sp>` shapes with the same
/// DrawingML shape/text pipeline as ordinary slide shapes. The cache uses the
/// same coordinate space as the frame extent, so its local offsets add
/// directly to the frame origin.
pub(super) fn parse_smartart_drawing(
    drawing_xml: &str,
    theme: &ThemeData,
    color_map: &ColorMapData,
    frame_x_pt: f64,
    frame_y_pt: f64,
) -> Vec<FixedElement> {
    let images: SlideImageMap = SlideImageMap::new();
    let table_styles: table_styles::TableStyleMap = table_styles::TableStyleMap::new();
    let inherited_text_body_defaults: PptxTextBodyStyleDefaults =
        PptxTextBodyStyleDefaults::default();
    let ctx = SlideParseContext {
        images: &images,
        slide_number: 1,
        theme,
        color_map,
        warning_context: "SmartArt drawing cache",
        inherited_text_body_defaults: &inherited_text_body_defaults,
        table_styles: &table_styles,
        default_text_size_pt: None,
    };
    let Ok((mut elements, _warnings)) = parse_slide_xml_inner(drawing_xml, &ctx, false, None)
    else {
        return Vec::new();
    };
    for element in &mut elements {
        element.x += frame_x_pt;
        element.y += frame_y_pt;
    }
    elements
}

/// Collect Chart elements referenced by the slide XML.
fn collect_chart_elements<R: Read + std::io::Seek>(
    slide_xml: &str,
    slide_path: &str,
    archive: &mut ZipArchive<R>,
    theme: &ThemeData,
    color_map: &ColorMapData,
) -> Vec<FixedElement> {
    let chart_refs = scan_chart_refs(slide_xml);
    if chart_refs.is_empty() {
        return Vec::new();
    }

    // A series with no fill of its own takes the deck's theme accents rather
    // than the renderer's built-in palette (issue #670).
    let theme_accents: Vec<Color> = crate::parser::drawingml::theme_accent_palette(&theme.colors);
    // Chart text resolves its face the same way: the chart part names no theme
    // of its own, so the deck's font scheme settles both `+mn-lt` and the far
    // commoner case of a chart that names no face at all (issue #668).
    let theme_fonts: crate::parser::drawingml::ThemeFontScheme =
        crate::parser::drawingml::ThemeFontScheme {
            major_latin: theme.major_font.clone(),
            minor_latin: theme.minor_font.clone(),
        };
    // A series' own `<a:schemeClr>` fill resolves against the deck's theme —
    // the chart part declares none of its own (issue #876). The slide's colour
    // map comes with it, since `bg1`/`tx1` are aliases the map settles.
    let chart_scheme = crate::parser::drawingml::SchemeColors {
        colors: &theme.colors,
        aliases: &color_map.aliases,
    };
    let chart_data = load_chart_data(slide_path, archive, &chart_scheme, &theme_fonts);
    chart_refs
        .iter()
        .filter_map(|c_ref| {
            chart_data.get(&c_ref.chart_rid).map(|chart| {
                let mut chart: Chart = chart.clone();
                chart.theme_accent_colors = theme_accents.clone();
                chart.host = crate::ir::ChartHost::Presentation;
                theme_fonts.resolve_chart_text_fonts(&mut chart);
                FixedElement {
                    x: emu_to_pt(c_ref.x),
                    y: emu_to_pt(c_ref.y),
                    width: emu_to_pt(c_ref.cx),
                    height: emu_to_pt(c_ref.cy),
                    kind: FixedElementKind::Chart(Box::new(chart)),
                }
            })
        })
        .collect()
}

// ── Background resolution ───────────────────────────────────────────────

/// Resolved slide background: an optional solid color and gradient, plus an
/// optional picture fill given as (owning layer part path, image rel id).
struct ResolvedBackground {
    color: Option<Color>,
    gradient: Option<GradientFill>,
    image: Option<(String, String)>,
}

/// Resolve the slide background by checking slide -> layout -> master in
/// order. Within a layer, a `<p:bgPr>` gradient wins over a solid fill, then
/// a picture fill, then `<p:bgRef>` theme references resolved through the
/// theme fill style lists. The first layer with a resolvable background wins.
fn resolve_slide_background(
    chain: &SlideInheritanceChain,
    slide_path: &str,
    theme: &ThemeData,
) -> ResolvedBackground {
    let layers: [(Option<&str>, &str, &ColorMapData); 3] = [
        (
            Some(chain.slide_xml.as_str()),
            slide_path,
            &chain.slide_color_map,
        ),
        (
            chain.layout_xml.as_deref(),
            chain.layout_path.as_deref().unwrap_or(""),
            chain
                .layout_color_map
                .as_ref()
                .unwrap_or(&chain.master_color_map),
        ),
        (
            chain.master_xml.as_deref(),
            chain.master_path.as_deref().unwrap_or(""),
            &chain.master_color_map,
        ),
    ];

    for (layer_xml, layer_path, color_map) in layers {
        let Some(xml) = layer_xml else { continue };

        if let Some(gradient) = parse_background_gradient(xml, theme, color_map) {
            return ResolvedBackground {
                color: gradient.stops.first().map(|s| s.color),
                gradient: Some(gradient),
                image: None,
            };
        }
        if let Some(color) = parse_background_color(xml, theme, color_map) {
            return ResolvedBackground {
                color: Some(color),
                gradient: None,
                image: None,
            };
        }
        if let Some(rid) = parse_background_image_rid(xml) {
            return ResolvedBackground {
                color: None,
                gradient: None,
                image: Some((layer_path.to_string(), rid)),
            };
        }
        if let Some((color, gradient)) = parse_background_ref(xml, theme, color_map) {
            return ResolvedBackground {
                color,
                gradient,
                image: None,
            };
        }
    }

    ResolvedBackground {
        color: None,
        gradient: None,
        image: None,
    }
}

/// Build a full-page image element for a picture-fill background.
fn build_background_image_element<R: Read + std::io::Seek>(
    layer_path: &str,
    rid: &str,
    slide_size: PageSize,
    archive: &mut ZipArchive<R>,
) -> Option<FixedElement> {
    let images: SlideImageMap = load_slide_images(layer_path, archive);
    let asset = images.get(rid)?;
    let format = asset.format()?;
    Some(FixedElement {
        x: 0.0,
        y: 0.0,
        width: slide_size.width,
        height: slide_size.height,
        kind: FixedElementKind::Image(ImageData {
            data: asset.data.clone(),
            format,
            rotation_deg: None,
            flip_h: false,
            flip_v: false,
            width: Some(slide_size.width),
            height: Some(slide_size.height),
            crop: None,
            stroke: None,
            alignment: None,
            clip_shape: None,
            shadow: None,
            paragraph_spacing: None,
        }),
    })
}

// ── Public entry point ──────────────────────────────────────────────────

/// True when the slide's root `<p:sld>` element carries `show="0"` or
/// `show="false"` — PowerPoint omits such hidden slides from PDF export.
fn is_hidden_slide(slide_xml: &str) -> bool {
    let mut reader: Reader<&[u8]> = Reader::from_str(slide_xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                return e.local_name().as_ref() == b"sld"
                    && get_attr_str(e, b"show").is_some_and(|v| v == "0" || v == "false");
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
    }
}

/// Parse a single slide from the archive, returning a Page or an error.
/// Returns `Ok(None)` for hidden slides when `include_hidden_slides` is false,
/// which matches PowerPoint's default PDF export behavior.
///
/// Resolves the inheritance chain (slide -> layout -> master) and
/// prepends master/layout elements behind slide elements.
pub(super) fn parse_single_slide<R: Read + std::io::Seek>(
    slide_path: &str,
    slide_label: &str,
    slide_number: u32,
    slide_size: PageSize,
    presentation: &PresentationResources<'_>,
    archive: &mut ZipArchive<R>,
    include_hidden_slides: bool,
) -> Result<Option<(Page, Vec<ConvertWarning>)>, ConvertError> {
    let PresentationResources {
        theme,
        table_styles,
        default_text_size_pt,
    } = *presentation;
    let chain: SlideInheritanceChain = resolve_inheritance_chain(slide_path, theme, archive)?;

    if !include_hidden_slides && is_hidden_slide(&chain.slide_xml) {
        tracing::debug!(slide = slide_label, "skipping hidden slide");
        return Ok(None);
    }

    let slide_images: SlideImageMap = load_slide_images(slide_path, archive);
    let mut warnings: Vec<ConvertWarning> = Vec::new();

    let placeholder_geometry: PlaceholderGeometryMap = PlaceholderGeometryMap::build(
        chain.layout_xml.as_deref(),
        chain.master_xml.as_deref(),
        theme,
        chain
            .layout_color_map
            .as_ref()
            .unwrap_or(&chain.master_color_map),
        &chain.master_color_map,
        chain.master_text_styles.clone(),
    );

    let slide_ctx = SlideParseContext {
        images: &slide_images,
        slide_number,
        theme,
        color_map: &chain.slide_color_map,
        warning_context: slide_label,
        inherited_text_body_defaults: &chain.master_text_styles.other,
        table_styles,
        default_text_size_pt,
    };
    let (slide_elements, slide_warnings) =
        parse_slide_xml(&chain.slide_xml, &slide_ctx, Some(&placeholder_geometry))?;
    warnings.extend(slide_warnings);

    let mut elements: Vec<FixedElement> = Vec::new();

    // Master layer (bottom)
    if let Some(ref path) = chain.master_path
        && let Some(ref xml) = chain.master_xml
    {
        let master_label: String = format!("{slide_label} master");
        let (master_elems, master_warnings) = parse_layer_elements(
            SlideLayer {
                path,
                xml,
                color_map: &chain.master_color_map,
                label: &master_label,
                text_style_defaults: &chain.master_text_styles.other,
            },
            theme,
            slide_number,
            default_text_size_pt,
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
            SlideLayer {
                path,
                xml,
                color_map,
                label: &layout_label,
                text_style_defaults: &chain.master_text_styles.other,
            },
            theme,
            slide_number,
            default_text_size_pt,
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
        theme,
        &chain.slide_color_map,
    ));
    elements.extend(collect_chart_elements(
        &chain.slide_xml,
        slide_path,
        archive,
        theme,
        &chain.slide_color_map,
    ));

    let background: ResolvedBackground = resolve_slide_background(&chain, slide_path, theme);
    if let Some((layer_path, rid)) = &background.image
        && let Some(element) = build_background_image_element(layer_path, rid, slide_size, archive)
    {
        // Picture-fill backgrounds render as a full-page image behind
        // everything else on the slide.
        elements.insert(0, element);
    }

    Ok(Some((
        Page::Fixed(FixedPage {
            size: slide_size,
            elements,
            background_color: background.color,
            background_gradient: background.gradient,
        }),
        warnings,
    )))
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
    has_placeholder: bool,
    ph_type: Option<String>,
    ph_idx: Option<String>,
    /// True when the slide itself provides `<a:xfrm>`; placeholders without
    /// one inherit geometry from the layout/master chain.
    has_explicit_xfrm: bool,
    blip_embed: Option<String>,
    /// Fill alpha from `<a:blip><a:alphaModFix amt>` (0.0-1.0).
    blip_alpha: Option<f64>,
    /// Preset geometry name from `<a:prstGeom prst>` ("crop to shape").
    prst_geom: Option<String>,
    /// Subpaths flattened from `<a:custGeom>`, normalized to the picture box.
    /// A picture may crop to a custom shape just as it can to a preset one
    /// (issue #872).
    custom_geometry: Vec<crate::ir::Subpath>,
    /// Outer shadow from the picture's `<a:effectLst>` (issue #360).
    shadow: Option<Shadow>,
    /// First `<a:gd>` adjust value inside the picture's prstGeom avLst.
    prst_adj: Option<f64>,
    in_prst_geom: bool,
    svg_blip_embed: Option<String>,
    /// Office 2021 live-feed pictures (PowerPoint Cameo) use cover sizing,
    /// unlike an ordinary DrawingML `a:stretch` picture (issue #976).
    has_live_feed_properties: bool,
    img_layer_embeds: Vec<String>,
    crop: Option<ImageCrop>,
    /// Clockwise rotation from `a:xfrm/@rot` (issue #682).
    rotation_deg: Option<f64>,
    /// Picture-frame mirrors from `a:xfrm` (issue #1017).
    flip_h: bool,
    flip_v: bool,
    in_xfrm: bool,
    in_sp_pr: bool,
    in_ln: bool,
    ln_width_emu: i64,
    ln_color: Option<Color>,
    ln_dash_style: BorderLineStyle,
    /// The `a:ln` corner join, or `None` when it names none (issue #1090).
    ln_join: Option<LineJoin>,
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

/// Accumulated state for a `<p:sp>` or `<p:cxnSp>` element and its nested properties.
struct ShapeState {
    depth: usize,
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
    has_placeholder: bool,
    ph_type: Option<String>,
    ph_idx: Option<String>,
    /// True when the slide itself provides `<a:xfrm>`; placeholders without
    /// one inherit geometry from the layout/master chain.
    has_explicit_xfrm: bool,
    rotation_deg: Option<f64>,
    flip_h: bool,
    flip_v: bool,
    opacity: Option<f64>,
    shadow: Option<Shadow>,
    top_bevel: Option<TopBevel>,
    /// `<a:effectRef idx>` from `<p:style>`, resolved against the theme once
    /// the shape closes (issue #740).
    style_effect_idx: Option<i64>,
    style_effect_color: Option<Color>,
    /// Whether `<p:spPr>` stated an `<a:effectLst>` of its own. An empty one
    /// means "no effect" and must not fall through to the style reference,
    /// which an absent shadow alone cannot distinguish.
    has_direct_effect_lst: bool,
    in_sp_pr: bool,
    prst_geom: Option<String>,
    /// Subpaths flattened from `<a:custGeom>`, each normalized to the shape
    /// box. Empty when the geometry yielded nothing usable, in which case the
    /// rectangle fallback stands (issues #855, #866).
    custom_geometry: Vec<crate::ir::Subpath>,
    /// `<a:custGeom><a:rect>` text bounds, normalized to the shape box.
    custom_text_rect: Option<super::custom_geometry::GeometryTextRect>,
    fill: Option<Color>,
    gradient_fill: Option<GradientFill>,
    pattern_fill: Option<PatternFill>,
    /// Relationship id from an ordinary shape's `<a:blipFill>`. DrawingML
    /// permits a bitmap to paint a `<p:sp>` just as it permits solid,
    /// gradient, and pattern fills (issue #1220).
    blip_embed: Option<String>,
    /// Fill alpha from `<a:blip><a:alphaModFix amt>` (0.0-1.0).
    blip_alpha: Option<f64>,
    /// Source crop from the shape fill's `<a:srcRect>`.
    blip_crop: Option<ImageCrop>,
    in_blip_fill: bool,
    in_xfrm: bool,
    in_ln: bool,
    ln_width_emu: i64,
    ln_color: Option<Color>,
    ln_dash_style: BorderLineStyle,
    /// The `a:ln` corner join, or `None` when it names none, in which case the
    /// `<a:lnRef>` theme line decides and DrawingML's round default backs it
    /// (issue #1090).
    ln_join: Option<LineJoin>,
    /// Arrowhead at line start.
    head_end: ArrowHead,
    /// Arrowhead at line end.
    tail_end: ArrowHead,
    /// Preset adjustment values from `<a:avLst><a:gd>`, in document order.
    adj_values: Vec<f64>,
    /// Fallback line color from `<p:style><a:lnRef>` scheme reference.
    style_ln_color: Option<Color>,
    /// `<a:lnRef idx>` (1-based) into the theme line style list, for the
    /// fallback outline width when no explicit `<a:ln w>` is present.
    style_ln_idx: Option<usize>,
    /// Fallback fill color from `<p:style><a:fillRef>` scheme reference.
    style_fill_color: Option<Color>,
    /// 1-based theme `fillStyleLst` entry named by `<a:fillRef idx>`.
    style_fill_idx: Option<usize>,
    /// Fallback text color from `<p:style><a:fontRef>` scheme reference.
    style_font_color: Option<Color>,
    /// Theme face selected by `<p:style><a:fontRef idx>`.
    style_font_family: Option<String>,
    /// True when `<a:noFill/>` is explicitly set in `<p:spPr>`, preventing style fallback.
    explicit_no_fill: bool,
    /// True when `<a:noFill/>` sits inside `<a:ln>`: PowerPoint's "No line",
    /// which must also defeat the `<p:style><a:lnRef>` fallback (issue #516).
    explicit_no_line: bool,
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
            ph_type: None,
            ph_idx: None,
            has_explicit_xfrm: false,
            style_effect_idx: None,
            style_effect_color: None,
            has_direct_effect_lst: false,
            rotation_deg: None,
            flip_h: false,
            flip_v: false,
            opacity: None,
            shadow: None,
            top_bevel: None,
            in_sp_pr: false,
            prst_geom: None,
            custom_geometry: Vec::new(),
            custom_text_rect: None,
            fill: None,
            gradient_fill: None,
            pattern_fill: None,
            blip_embed: None,
            blip_alpha: None,
            blip_crop: None,
            in_blip_fill: false,
            in_xfrm: false,
            in_ln: false,
            ln_width_emu: 0,
            ln_color: None,
            ln_dash_style: BorderLineStyle::Solid,
            ln_join: None,
            head_end: ArrowHead::None,
            tail_end: ArrowHead::None,
            adj_values: Vec::new(),
            style_ln_color: None,
            style_ln_idx: None,
            style_fill_color: None,
            style_fill_idx: None,
            style_font_color: None,
            style_font_family: None,
            explicit_no_fill: false,
            explicit_no_line: false,
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
/// Returns elements to add. Text shapes that require the shape/picture
/// renderer — non-rectangular geometry, gradient, pattern, or picture fills,
/// shadows, or top bevels — return a background plus a transparent text
/// overlay.
fn finalize_shape(
    shape: &mut ShapeState,
    paragraphs: &mut Vec<PptxParagraphEntry>,
    text_box: PptxTextBoxSettings,
    theme_line_styles: &[ThemeLineStyle],
    images: &SlideImageMap,
    warning_context: &str,
    warnings: &mut Vec<ConvertWarning>,
) -> Vec<FixedElement> {
    let referenced_line_style: Option<&ThemeLineStyle> = shape
        .style_ln_idx
        .and_then(|idx| theme_line_styles.get(idx - 1));
    // Outline width: explicit `<a:ln w>` when present, otherwise the theme
    // line style referenced by `<a:lnRef idx>` (issue #318).
    let effective_ln_width_emu: i64 = if shape.ln_width_emu > 0 {
        shape.ln_width_emu
    } else {
        referenced_line_style
            .map(|style| style.width_emu)
            .unwrap_or(shape.ln_width_emu)
    };
    let effective_ln_width_pt: f64 = emu_to_pt(effective_ln_width_emu);
    // Corner join: the shape's own `a:ln` child, else the referenced theme
    // line's, else DrawingML's round default (issue #1090).
    let effective_ln_join: LineJoin = shape
        .ln_join
        .or_else(|| referenced_line_style.and_then(|style| style.join))
        .unwrap_or_default();

    // Resolve effective fill: explicit > noFill > style fallback.
    let effective_fill: Option<Color> = if shape.fill.is_some() {
        shape.fill
    } else if shape.explicit_no_fill {
        None
    } else {
        shape.style_fill_color
    };

    // The outline is shared by solid/gradient shape paint and by an ordinary
    // shape's picture fill. Building it once also lets the picture reuse the
    // established `<p:pic>` finalizer instead of growing a parallel image
    // renderer (issue #1220).
    let effective_ln_color: Option<Color> = if shape.explicit_no_line {
        None
    } else {
        shape.ln_color.or(shape.style_ln_color)
    };
    let stroke: Option<BorderSide> = effective_ln_color.map(|color| BorderSide {
        width: effective_ln_width_pt,
        color,
        style: shape.ln_dash_style,
        join: effective_ln_join,
    });
    let mut picture_fill: Option<FixedElement> = if shape.blip_embed.is_some() {
        let picture = PictureState {
            x: shape.x,
            y: shape.y,
            cx: shape.cx,
            cy: shape.cy,
            blip_embed: shape.blip_embed.clone(),
            blip_alpha: shape.blip_alpha,
            prst_geom: shape.prst_geom.clone(),
            custom_geometry: shape.custom_geometry.clone(),
            crop: shape.blip_crop,
            rotation_deg: shape.rotation_deg,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
            ln_width_emu: effective_ln_width_emu,
            ln_color: effective_ln_color,
            ln_dash_style: shape.ln_dash_style,
            ln_join: Some(effective_ln_join),
            shadow: shape.shadow.clone(),
            ..PictureState::default()
        };
        let (element, picture_warnings) = finalize_picture(&picture, images, warning_context);
        warnings.extend(picture_warnings);
        element
    } else {
        None
    };

    let has_text = paragraphs
        .iter()
        .any(|entry| !entry.paragraph.runs.is_empty());

    if has_text {
        apply_pptx_saved_normal_autofit(paragraphs, &text_box);
        let blocks: Vec<Block> = group_pptx_text_blocks(std::mem::take(paragraphs));
        // Use explicit line color, falling back to style-based color from
        // <p:style><a:lnRef> - unless <a:ln><a:noFill/> disabled the line.
        // For shapes with text that need a background of their own — a
        // non-rectangular geometry, a gradient, pattern, or picture fill, a
        // shadow, or a top bevel — emit the shape background first, then
        // overlay a transparent text box, so the paint or geometry goes
        // through the proven shape/picture renderer. A plain rectangle
        // otherwise skips this and becomes a text box with a background fill,
        // which is cheaper and lays the text out the same.
        // A shadow needs something to cast it. A plain rectangle with text is
        // otherwise drawn as a text box with a background fill, and a text box
        // has nowhere to put a shadow, so the theme shadow of issue #740 was
        // resolved and then dropped here. Emitting the shape background makes
        // the existing shape renderer draw it.
        let needs_shape_background = shape.gradient_fill.is_some()
            || shape.pattern_fill.is_some()
            || picture_fill.is_some()
            || shape.shadow.is_some()
            || shape.top_bevel.is_some();
        let shape_width = emu_to_pt(shape.cx);
        let shape_height = emu_to_pt(shape.cy);
        let has_geometry_text_rect = shape.custom_text_rect.is_some()
            || shape.prst_geom.as_deref().is_some_and(|geom| {
                preset_text_rect_insets(geom, shape_width, shape_height, &shape.adj_values)
                    .is_some()
            });
        let text_shape_kind: Option<ShapeKind> = shape
            .prst_geom
            .as_deref()
            .and_then(|geom| {
                if let Some(kind) = custom_geometry_kind(shape) {
                    return Some(kind);
                }
                let kind: ShapeKind = prst_to_shape_kind(
                    geom,
                    shape_width,
                    shape_height,
                    shape.flip_h,
                    shape.flip_v,
                    shape.head_end,
                    shape.tail_end,
                    &shape.adj_values,
                );
                match kind {
                    ShapeKind::Rectangle if !needs_shape_background && !has_geometry_text_rect => {
                        None
                    }
                    other => Some(other),
                }
            })
            .or_else(|| picture_fill.is_some().then_some(ShapeKind::Rectangle));
        let mut elements: Vec<FixedElement> = Vec::new();
        if let Some(kind) = text_shape_kind {
            // Shape background element (picture, or vector fill + geometry).
            if let Some(picture) = picture_fill.take() {
                elements.push(picture);
            } else {
                elements.push(FixedElement {
                    x: emu_to_pt(shape.x),
                    y: emu_to_pt(shape.y),
                    width: emu_to_pt(shape.cx),
                    height: emu_to_pt(shape.cy),
                    kind: FixedElementKind::Shape(Shape {
                        kind,
                        fill: effective_fill,
                        gradient_fill: shape.gradient_fill.take(),
                        pattern_fill: shape.pattern_fill.take(),
                        stroke: stroke.clone(),
                        rotation_deg: shape.rotation_deg,
                        opacity: shape.opacity,
                        shadow: shape.shadow.take(),
                        top_bevel: shape.top_bevel.take(),
                    }),
                });
            }
            // Transparent text overlay (no fill, no stroke). DrawingML
            // anchors shape text inside the preset geometry's text rectangle,
            // in addition to the bodyPr margins (issues #286 and #676).
            let shape_x = emu_to_pt(shape.x);
            let shape_y = emu_to_pt(shape.y);
            let geometry_text_rect = shape.prst_geom.as_deref().and_then(|geom| {
                let rotation_deg = shape.rotation_deg.unwrap_or(0.0).rem_euclid(360.0);
                // A transparent overlay cannot represent an independently
                // oriented geometry rectangle. Limit this exact model to
                // transforms that preserve its axes; other rotations retain
                // the previous full-box safety fallback.
                let preserves_axes =
                    rotation_deg.abs() < 1e-9 || (rotation_deg - 180.0).abs() < 1e-9;
                if !preserves_axes {
                    return None;
                }
                let insets = shape
                    .custom_text_rect
                    .map(|rect| Insets {
                        left: rect.left * shape_width,
                        top: rect.top * shape_height,
                        right: (1.0 - rect.right) * shape_width,
                        bottom: (1.0 - rect.bottom) * shape_height,
                    })
                    .or_else(|| {
                        preset_text_rect_insets(geom, shape_width, shape_height, &shape.adj_values)
                    });
                insets.map(|insets| {
                    let rect_width = (shape_width - insets.left - insets.right).max(0.0);
                    let rect_height = (shape_height - insets.top - insets.bottom).max(0.0);
                    let rect_center_x = shape_x + insets.left + rect_width / 2.0;
                    let rect_center_y = shape_y + insets.top + rect_height / 2.0;
                    let shape_center_x = shape_x + shape_width / 2.0;
                    let shape_center_y = shape_y + shape_height / 2.0;
                    // The text body's reading direction is independent from
                    // a:xfrm (issue #992), but the preset text rectangle is
                    // part of the geometry and its center follows that
                    // transform (issue #676).
                    let rotation = rotation_deg.to_radians();
                    let (sin, cos) = rotation.sin_cos();
                    let dx = rect_center_x - shape_center_x;
                    let dy = rect_center_y - shape_center_y;
                    let rotated_center_x = shape_center_x + dx * cos - dy * sin;
                    let rotated_center_y = shape_center_y + dx * sin + dy * cos;
                    (
                        rotated_center_x - rect_width / 2.0,
                        rotated_center_y - rect_height / 2.0,
                        rect_width,
                        rect_height,
                    )
                })
            });
            let (overlay_x, overlay_y, overlay_width, overlay_height) =
                geometry_text_rect.unwrap_or((shape_x, shape_y, shape_width, shape_height));
            // Preserve the old safety approximation for presets whose text
            // rectangle is not modelled yet: edge-anchoring rotated text to
            // the full shape box can put it on a sloped boundary.
            let overlay_vertical_align =
                if text_box.text_rotation_deg.is_some() && geometry_text_rect.is_none() {
                    TextBoxVerticalAlign::Center
                } else {
                    text_box.vertical_align
                };
            elements.push(FixedElement {
                x: overlay_x,
                y: overlay_y,
                width: overlay_width,
                height: overlay_height,
                kind: FixedElementKind::TextBox(TextBoxData {
                    content: blocks,
                    padding: text_box.padding,
                    vertical_align: overlay_vertical_align,
                    fill: None,
                    opacity: None,
                    stroke: None,
                    shape_kind: None,
                    no_wrap: text_box.no_wrap,
                    auto_fit: text_box.requests_dynamic_autofit(),
                    text_rotation_deg: text_box.text_rotation_deg,
                    // A preset's geometry rotates independently from an
                    // explicit vertical text body. Composing the same xfrm
                    // rotation into that transparent overlay reverses the
                    // vertical reading direction (issue #992).
                    shape_rotation_deg: if text_box.text_rotation_deg.is_some() {
                        None
                    } else {
                        shape.rotation_deg
                    },
                }),
            });
        } else {
            // Simple rectangular text box with fill/stroke directly on the block.
            elements.push(FixedElement {
                x: emu_to_pt(shape.x),
                y: emu_to_pt(shape.y),
                width: emu_to_pt(shape.cx),
                height: emu_to_pt(shape.cy),
                kind: FixedElementKind::TextBox(TextBoxData {
                    content: blocks,
                    padding: text_box.padding,
                    vertical_align: text_box.vertical_align,
                    fill: effective_fill,
                    opacity: shape.opacity,
                    stroke,
                    shape_kind: None,
                    no_wrap: text_box.no_wrap,
                    auto_fit: text_box.requests_dynamic_autofit(),
                    text_rotation_deg: text_box.text_rotation_deg,
                    shape_rotation_deg: shape.rotation_deg,
                }),
            });
        }
        elements
    } else if let Some(picture) = picture_fill {
        vec![picture]
    } else if let Some(ref geom) = shape.prst_geom {
        let width: f64 = emu_to_pt(shape.cx);
        let height: f64 = emu_to_pt(shape.cy);
        let kind: ShapeKind = custom_geometry_kind(shape).unwrap_or_else(|| {
            prst_to_shape_kind(
                geom,
                width,
                height,
                shape.flip_h,
                shape.flip_v,
                shape.head_end,
                shape.tail_end,
                &shape.adj_values,
            )
        });
        // Use explicit line color, falling back to style-based color from
        // <p:style><a:lnRef> - unless <a:ln><a:noFill/> disabled the line.
        vec![FixedElement {
            x: emu_to_pt(shape.x),
            y: emu_to_pt(shape.y),
            width,
            height,
            kind: FixedElementKind::Shape(Shape {
                kind,
                fill: effective_fill,
                gradient_fill: shape.gradient_fill.take(),
                pattern_fill: shape.pattern_fill.take(),
                stroke,
                rotation_deg: shape.rotation_deg,
                opacity: shape.opacity,
                shadow: shape.shadow.take(),
                top_bevel: shape.top_bevel.take(),
            }),
        }]
    } else {
        Vec::new()
    }
}

/// The path a shape's `<a:custGeom>` flattened to, or `None` when the geometry
/// yielded nothing usable and the rectangle fallback stands (issue #855).
///
/// Every subpath rides in one [`ShapeKind::Path`], filled under the even-odd
/// rule: a geometry is one path however many subpaths it holds, so an inner
/// boundary carves a hole (issue #870). Concatenating them into a single ring
/// instead welded each outline's end to the next one's start (issue #866).
fn custom_geometry_kind(shape: &ShapeState) -> Option<ShapeKind> {
    (!shape.custom_geometry.is_empty()).then(|| {
        let mut subpaths = shape.custom_geometry.clone();
        // DrawingML applies `flipH`/`flipV` about the shape box's centre. A
        // custom path is already normalized to that box, so mirroring each
        // normalized axis completes the shape transform (issue #1418).
        for vertex in subpaths
            .iter_mut()
            .flat_map(|subpath| &mut subpath.vertices)
        {
            if shape.flip_h {
                vertex.0 = 1.0 - vertex.0;
            }
            if shape.flip_v {
                vertex.1 = 1.0 - vertex.1;
            }
        }
        ShapeKind::Path { subpaths }
    })
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
    let stroke: Option<BorderSide> = pic.ln_color.map(|color| BorderSide {
        width: emu_to_pt(pic.ln_width_emu),
        color,
        style: pic.ln_dash_style,
        join: pic.ln_join.unwrap_or_default(),
    });
    let element = selected_asset.and_then(|asset| {
        asset.format().map(|format| {
            let mut clip_shape = picture_clip_shape(pic.prst_geom.as_deref(), pic.prst_adj);
            let (data, format) = match pic.blip_alpha {
                Some(alpha) if alpha < 1.0 => {
                    crate::parser::drawingml::apply_image_alpha(&asset.data, alpha)
                        .unwrap_or_else(|| (asset.data.clone(), format))
                }
                _ => (asset.data.clone(), format),
            };
            // DrawingML `a:stretch` normally scales non-uniformly. PowerPoint
            // makes Office 2021 live-feed (Cameo) artwork an exception: it
            // centre-crops the feed to the frame's aspect ratio. Narrow the SVG
            // first so its later custom-geometry clip still spans the frame
            // rather than being cropped together with the artwork (issue #976).
            let data: Vec<u8> = if pic.has_live_feed_properties && format == ImageFormat::Svg {
                cover_crop_svg_to_aspect_ratio(&data, pic.cx as f64, pic.cy as f64).unwrap_or(data)
            } else {
                data
            };
            // Typst's corner radius expresses a rounded rectangle and nothing
            // else, so the crops it cannot draw are clipped out here instead:
            // baked into the alpha mask for a raster, wrapped in a
            // `<clipPath>` for an SVG (issue #897).
            let (data, format) = if !pic.custom_geometry.is_empty() {
                // A custom geometry has no corner-radius equivalent at all
                // (issue #872). An SVG has no raster to mask, so it takes the
                // same path as a `<clipPath>` instead (issue #897).
                // PowerPoint crops with `a:srcRect` first and clips the
                // resulting frame second. The renderer applies the crop later,
                // so the clip written here has to live in the cropped frame's
                // coordinates or it lands on the pre-crop artwork (#1018).
                let clipped: Option<(Vec<u8>, ImageFormat)> = if format == ImageFormat::Svg {
                    clip_svg_to_path(&data, &pic.custom_geometry, pic.crop)
                        .map(|svg| (svg, ImageFormat::Svg))
                } else {
                    apply_path_mask(&data, &pic.custom_geometry, pic.crop)
                };
                match clipped {
                    Some(masked) => {
                        clip_shape = None;
                        masked
                    }
                    None => (data, format),
                }
            } else if clip_shape == Some(ImageClipShape::Ellipse) {
                // A radius cannot describe a true ellipse on a non-square box.
                match apply_ellipse_mask(&data) {
                    Some(masked) => {
                        clip_shape = None;
                        masked
                    }
                    None => (data, format),
                }
            } else {
                (data, format)
            };
            FixedElement {
                x: emu_to_pt(pic.x),
                y: emu_to_pt(pic.y),
                width: emu_to_pt(pic.cx),
                height: emu_to_pt(pic.cy),
                kind: FixedElementKind::Image(ImageData {
                    data,
                    format,
                    rotation_deg: pic.rotation_deg,
                    flip_h: pic.flip_h,
                    flip_v: pic.flip_v,
                    width: Some(emu_to_pt(pic.cx)),
                    height: Some(emu_to_pt(pic.cy)),
                    crop: pic.crop,
                    stroke: stroke.clone(),
                    alignment: None,
                    clip_shape,
                    shadow: pic.shadow.clone(),
                    paragraph_spacing: None,
                }),
            }
        })
    });
    (element, picture_warnings)
}

/// Map a picture's preset geometry to a renderable clip shape
/// (PowerPoint "crop to shape"); unsupported geometries clip nothing.
fn picture_clip_shape(
    prst: Option<&str>,
    adjust: Option<f64>,
) -> Option<crate::ir::ImageClipShape> {
    match prst? {
        "ellipse" => Some(crate::ir::ImageClipShape::Ellipse),
        "roundRect" | "round1Rect" | "round2SameRect" => Some(
            crate::ir::ImageClipShape::RoundedRect(adjust.unwrap_or(0.16667).clamp(0.0, 0.5)),
        ),
        _ => None,
    }
}

/// Zero the alpha outside the inscribed ellipse and re-encode as PNG.
fn apply_ellipse_mask(data: &[u8]) -> Option<(Vec<u8>, ImageFormat)> {
    let decoded = image::load_from_memory(data).ok()?;
    let mut rgba = decoded.into_rgba8();
    let (width, height) = rgba.dimensions();
    if width == 0 || height == 0 {
        return None;
    }
    let (cx, cy) = (f64::from(width) / 2.0, f64::from(height) / 2.0);
    for (x, y, pixel) in rgba.enumerate_pixels_mut() {
        let nx = (f64::from(x) + 0.5 - cx) / cx;
        let ny = (f64::from(y) + 0.5 - cy) / cy;
        if nx * nx + ny * ny > 1.0 {
            pixel[3] = 0;
        }
    }
    let mut out = std::io::Cursor::new(Vec::new());
    image::DynamicImage::ImageRgba8(rgba)
        .write_to(&mut out, image::ImageFormat::Png)
        .ok()?;
    Some((out.into_inner(), ImageFormat::Png))
}

/// The fraction of the source an `a:srcRect` keeps, as `(left, top, kept
/// width, kept height)`, mirroring the clamps of the renderer's crop.
///
/// `None` when there is no crop to apply — absent, empty, or degenerate
/// (nothing kept) — the three cases where the renderer leaves the asset at
/// its full extent, so a clip must span the full box too.
fn crop_kept_fractions(crop: Option<ImageCrop>) -> Option<(f64, f64, f64, f64)> {
    let crop: ImageCrop = crop.filter(|crop| !crop.is_empty())?;
    let left: f64 = crop.left.clamp(0.0, 1.0);
    let top: f64 = crop.top.clamp(0.0, 1.0);
    let kept_w: f64 = 1.0 - left - crop.right.clamp(0.0, 1.0);
    let kept_h: f64 = 1.0 - top - crop.bottom.clamp(0.0, 1.0);
    (kept_w > 0.0 && kept_h > 0.0).then_some((left, top, kept_w, kept_h))
}

/// Clip an SVG to `subpaths` by wrapping its content in a `<clipPath>`.
///
/// An SVG has no raster to mask, so [`apply_path_mask`] could not touch one
/// and a curved panel rendered as its bounding rectangle (issue #897). The
/// subpaths arrive normalised to 0..1 of the picture box, so they scale onto
/// the root's own `viewBox` to become user-space coordinates.
///
/// `crop` is the picture's `a:srcRect`, applied later by the renderer by
/// narrowing the viewBox: the picture box maps onto the region that
/// narrowing keeps, not the full viewBox (issue #1018).
///
/// Returns `None` when the root states no usable `viewBox`, leaving the asset
/// alone rather than clipping it against guessed units.
fn clip_svg_to_path(
    data: &[u8],
    subpaths: &[crate::ir::Subpath],
    crop: Option<ImageCrop>,
) -> Option<Vec<u8>> {
    use std::fmt::Write as _;

    if subpaths.is_empty() {
        return None;
    }
    let text: &str = std::str::from_utf8(data).ok()?;
    let open_end: usize = text.find('>')?;
    if !text[..open_end].trim_start().starts_with("<svg") {
        return None;
    }
    let close: usize = text.rfind("</svg>")?;

    let attr_start: usize = text[..open_end].find("viewBox=\"")? + "viewBox=\"".len();
    let attr_len: usize = text[attr_start..open_end].find('"')?;
    let values: Vec<f64> = text[attr_start..attr_start + attr_len]
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|part| !part.is_empty())
        .map(str::parse::<f64>)
        .collect::<Result<Vec<f64>, _>>()
        .ok()?;
    let [x, y, width, height] = values[..] else {
        return None;
    };
    if !(width.is_finite() && height.is_finite() && width > 0.0 && height > 0.0) {
        return None;
    }
    let (frame_x, frame_y, frame_w, frame_h) = match crop_kept_fractions(crop) {
        Some((left, top, kept_w, kept_h)) => (
            x + width * left,
            y + height * top,
            width * kept_w,
            height * kept_h,
        ),
        None => (x, y, width, height),
    };

    let mut path: String = String::new();
    // A clip region is an area however the geometry ended, so every subpath
    // closes here whether or not it stated `a:close`.
    for subpath in subpaths.iter().filter(|path| path.vertices.len() >= 3) {
        for (index, (fx, fy)) in subpath.vertices.iter().enumerate() {
            let _ = write!(
                path,
                "{} {} {} ",
                if index == 0 { "M" } else { "L" },
                format_svg_number(frame_x + fx * frame_w),
                format_svg_number(frame_y + fy * frame_h),
            );
        }
        path.push_str("Z ");
    }
    let path: &str = path.trim_end();
    if path.is_empty() {
        return None;
    }

    let mut out: String = String::with_capacity(text.len() + path.len() + 128);
    out.push_str(&text[..=open_end]);
    let _ = write!(
        out,
        "<clipPath id=\"o2pPicClip\" clipPathUnits=\"userSpaceOnUse\"><path clip-rule=\"evenodd\" d=\"{path}\"/></clipPath><g clip-path=\"url(#o2pPicClip)\" clip-rule=\"evenodd\">"
    );
    out.push_str(&text[open_end + 1..close]);
    out.push_str("</g>");
    out.push_str(&text[close..]);
    Some(out.into_bytes())
}

/// Centre-crop an SVG viewport to `target_width / target_height`.
///
/// PowerPoint Cameo preserves the live feed's proportions while covering its
/// picture frame. Rewriting both the `viewBox` and root viewport preserves that
/// behaviour through Typst's later explicit-width-and-height stretch.
fn cover_crop_svg_to_aspect_ratio(
    data: &[u8],
    target_width: f64,
    target_height: f64,
) -> Option<Vec<u8>> {
    if !(target_width.is_finite()
        && target_height.is_finite()
        && target_width > 0.0
        && target_height > 0.0)
    {
        return None;
    }

    let text: &str = std::str::from_utf8(data).ok()?;
    let open_end: usize = text.find('>')?;
    let head: &str = &text[..open_end];
    if !head.trim_start().starts_with("<svg") {
        return None;
    }
    let attr_start: usize = head.find("viewBox=\"")? + "viewBox=\"".len();
    let attr_len: usize = head[attr_start..].find('"')?;
    let values: Vec<f64> = head[attr_start..attr_start + attr_len]
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|part| !part.is_empty())
        .map(str::parse::<f64>)
        .collect::<Result<Vec<f64>, _>>()
        .ok()?;
    let [x, y, width, height] = values[..] else {
        return None;
    };
    if !(x.is_finite()
        && y.is_finite()
        && width.is_finite()
        && height.is_finite()
        && width > 0.0
        && height > 0.0)
    {
        return None;
    }

    let target_aspect_ratio: f64 = target_width / target_height;
    let source_aspect_ratio: f64 = width / height;
    let (kept_x, kept_y, kept_width, kept_height): (f64, f64, f64, f64) =
        if source_aspect_ratio > target_aspect_ratio {
            let kept_width: f64 = height * target_aspect_ratio;
            (x + (width - kept_width) / 2.0, y, kept_width, height)
        } else {
            let kept_height: f64 = width / target_aspect_ratio;
            (x, y + (height - kept_height) / 2.0, width, kept_height)
        };

    let replacement: String = format!(
        "viewBox=\"{} {} {} {}\"",
        format_svg_number(kept_x),
        format_svg_number(kept_y),
        format_svg_number(kept_width),
        format_svg_number(kept_height)
    );
    let mut out: String = String::with_capacity(text.len() + replacement.len());
    out.push_str(&text[..attr_start - "viewBox=\"".len()]);
    out.push_str(&replacement);
    out.push_str(&text[attr_start + attr_len + 1..]);

    let out: String = replace_svg_root_length(&out, "width", kept_width);
    let out: String = replace_svg_root_length(&out, "height", kept_height);
    Some(out.into_bytes())
}

/// Rewrite one root `<svg>` length while preserving its unit suffix.
fn replace_svg_root_length(text: &str, name: &str, value: f64) -> String {
    let Some(open_end) = text.find('>') else {
        return text.to_string();
    };
    let needle: String = format!("{name}=\"");
    let Some(offset) = text[..open_end].find(&needle) else {
        return text.to_string();
    };
    let start: usize = offset + needle.len();
    let Some(len) = text[start..open_end].find('"') else {
        return text.to_string();
    };
    let old: &str = &text[start..start + len];
    let unit: &str = old.trim_start_matches(|c: char| c.is_ascii_digit() || c == '.' || c == '-');
    format!(
        "{}{}{}{}",
        &text[..start],
        format_svg_number(value),
        unit,
        &text[start + len..]
    )
}

/// A number for SVG markup: no exponent, and no trailing zeros to read past.
fn format_svg_number(value: f64) -> String {
    let rounded: f64 = (value * 1000.0).round() / 1000.0;
    let mut text: String = format!("{rounded}");
    if text.contains('.') {
        text = text.trim_end_matches('0').trim_end_matches('.').to_string();
    }
    text
}

/// Clear the alpha of every pixel outside `subpaths`, and re-encode as PNG.
///
/// A picture may carry an `a:custGeom` just as a shape does — the deck on #872
/// clips its team photos to circles and its background art to discs that way —
/// and Typst's corner radius cannot express those, so the clip is baked into
/// the alpha channel the same way [`apply_ellipse_mask`] bakes an ellipse.
///
/// Subpaths are normalized to 0..1 of the picture box and filled even-odd, so
/// an inner boundary carves a hole (issues #866, #870).
fn apply_path_mask(
    data: &[u8],
    subpaths: &[crate::ir::Subpath],
    crop: Option<ImageCrop>,
) -> Option<(Vec<u8>, ImageFormat)> {
    let decoded = image::load_from_memory(data).ok()?;
    let mut rgba = decoded.into_rgba8();
    let (width, height) = rgba.dimensions();
    if width == 0 || height == 0 || subpaths.is_empty() {
        return None;
    }
    // The renderer crops the bitmap after this mask is baked, so the path is
    // evaluated in the cropped frame's coordinates: a full-bitmap mapping
    // would clip the pre-crop artwork instead (issue #1018).
    let (frame_x, frame_y, frame_w, frame_h) = match crop_kept_fractions(crop) {
        Some((left, top, kept_w, kept_h)) => (left, top, kept_w, kept_h),
        None => (0.0, 0.0, 1.0, 1.0),
    };
    for (x, y, pixel) in rgba.enumerate_pixels_mut() {
        let px = ((f64::from(x) + 0.5) / f64::from(width) - frame_x) / frame_w;
        let py = ((f64::from(y) + 0.5) / f64::from(height) - frame_y) / frame_h;
        if !point_is_inside_even_odd(subpaths, px, py) {
            pixel[3] = 0;
        }
    }
    let mut out = std::io::Cursor::new(Vec::new());
    image::DynamicImage::ImageRgba8(rgba)
        .write_to(&mut out, image::ImageFormat::Png)
        .ok()?;
    Some((out.into_inner(), ImageFormat::Png))
}

/// Whether a point falls inside a set of subpaths under the even-odd rule:
/// cast a ray and count the edges it crosses.
///
/// Each subpath is treated as closed, as a filled region is: DrawingML closes
/// an unclosed outline to fill it and only leaves it open for the stroke.
fn point_is_inside_even_odd(subpaths: &[crate::ir::Subpath], px: f64, py: f64) -> bool {
    let mut crossings: usize = 0;
    for subpath in subpaths {
        let vertices: &[(f64, f64)] = &subpath.vertices;
        if vertices.len() < 3 {
            continue;
        }
        for index in 0..vertices.len() {
            let (x1, y1) = vertices[index];
            let (x2, y2) = vertices[(index + 1) % vertices.len()];
            // A horizontal edge cannot be crossed by a horizontal ray, and the
            // half-open test keeps a vertex from counting twice.
            if (y1 > py) != (y2 > py) {
                let t: f64 = (py - y1) / (y2 - y1);
                if px < x1 + t * (x2 - x1) {
                    crossings += 1;
                }
            }
        }
    }
    crossings % 2 == 1
}

/// Apply a parsed solid fill color to the appropriate target based on the current context.
fn apply_solid_fill_color(
    ctx: SolidFillCtx,
    parsed: &ParsedColor,
    shape: &mut ShapeState,
    run_style: &mut TextStyle,
    end_run_style: &mut TextStyle,
    bullet_def: &mut PptxBulletDefinition,
    pic: &mut PictureState,
) {
    match ctx {
        SolidFillCtx::ShapeFill => {
            shape.fill = parsed.color;
            if let Some(alpha) = parsed.alpha {
                shape.opacity = Some(alpha);
            }
        }
        SolidFillCtx::LineFill => shape.ln_color = parsed.color,
        SolidFillCtx::RunFill => {
            run_style.color = parsed.color;
            run_style.color_alpha = parsed.alpha;
        }
        SolidFillCtx::EndParaFill => {
            end_run_style.color = parsed.color;
            end_run_style.color_alpha = parsed.alpha;
        }
        SolidFillCtx::BulletFill => {
            bullet_def.color = parsed.color.map(PptxBulletColorSource::Explicit);
        }
        SolidFillCtx::PicLineFill => pic.ln_color = parsed.color,
        SolidFillCtx::None => {}
    }
}

// ── SlideXmlParser state machine ────────────────────────────────────────

/// Deck-level resources that are identical for every slide.
///
/// Bundled so `parse_single_slide` takes them as one value; passing them
/// separately put it over `clippy::too_many_arguments`.
#[derive(Clone, Copy)]
pub(super) struct PresentationResources<'a> {
    pub(super) theme: &'a ThemeData,
    pub(super) table_styles: &'a table_styles::TableStyleMap,
    /// `p:defaultTextStyle/a:lvl1pPr/a:defRPr/@sz`, the size a text body falls
    /// back to when its own chain declares none (issue #675).
    pub(super) default_text_size_pt: Option<f64>,
}

/// Shared read-only inputs for parsing one slide-layer XML part.
///
/// Groups the per-slide references that every sub-parser needs, so they
/// travel as one value instead of six positional parameters.
#[derive(Clone, Copy)]
pub(super) struct SlideParseContext<'a> {
    pub(super) images: &'a SlideImageMap,
    /// The slide's 1-based position in the deck, for `<a:fld type="slidenum">`.
    pub(super) slide_number: u32,
    pub(super) theme: &'a ThemeData,
    pub(super) color_map: &'a ColorMapData,
    pub(super) warning_context: &'a str,
    pub(super) inherited_text_body_defaults: &'a PptxTextBodyStyleDefaults,
    pub(super) table_styles: &'a table_styles::TableStyleMap,
    /// `p:defaultTextStyle/a:lvl1pPr/a:defRPr/@sz`, the size a text body
    /// falls back to when its own chain declares none (issue #675).
    pub(super) default_text_size_pt: Option<f64>,
}

/// Bundles the 20+ mutable state variables of the slide XML event loop
/// into a single struct, with methods for each event type.
///
/// The XML reader is passed to each handler rather than stored, because
/// several sub-parsers (`parse_pptx_table`, `parse_group_shape`, etc.)
/// need `&mut Reader` to consume nested elements.
struct SlideXmlParser<'a> {
    // ── Context references (immutable for the parse lifetime) ────────
    xml: &'a str,
    ctx: SlideParseContext<'a>,

    // ── Options ─────────────────────────────────────────────────────
    /// When true, shapes with `<p:ph>` (placeholder) are skipped.
    /// Used when parsing master/layout layers whose placeholder content
    /// should not render unless the slide overrides it.
    skip_placeholders: bool,
    /// The layout/master placeholder chain a slide placeholder inherits from.
    /// Supplies geometry to one that omits `<a:xfrm>`, and shape fill to one
    /// that declares no fill — the two are looked up independently. None
    /// outside slide-layer parsing.
    placeholder_geometry: Option<&'a PlaceholderGeometryMap>,

    // ── Output accumulators ─────────────────────────────────────────
    elements: Vec<FixedElement>,
    warnings: Vec<ConvertWarning>,

    // ── Shape state (`<p:sp>`) ──────────────────────────────────────
    in_shape: bool,
    shape: ShapeState,

    // ── Text body state (`<p:txBody>`) ──────────────────────────────
    in_txbody: bool,
    paragraphs: Vec<PptxParagraphEntry>,
    text_box: PptxTextBoxSettings,
    text_body_style_defaults: PptxTextBodyStyleDefaults,

    // ── Paragraph state (`<a:p>`) ───────────────────────────────────
    in_para: bool,
    para_style: ParagraphStyle,
    para_level: u32,
    para_default_run_style: TextStyle,
    para_end_run_style: TextStyle,
    /// Whether the open paragraph wrote an `<a:endParaRPr>` element.
    /// An absent mark takes no inherited face onto the line (issue #1645).
    para_declares_end_para_rpr: bool,
    para_bullet_definition: PptxBulletDefinition,
    in_ln_spc: bool,
    in_spc_bef: bool,
    in_spc_aft: bool,
    runs: Vec<Run>,

    // ── Run state (`<a:r>`, `<a:fld>`) ──────────────────────────────
    in_run: bool,
    run_style: TextStyle,
    run_text: String,
    run_has_explicit_underline: bool,
    run_marker_style_before_hyperlink: Option<TextStyle>,
    first_run_marker_style_override: Option<TextStyle>,
    /// `<a:fld type>` of the run being read; `None` for a literal `<a:r>`.
    run_field_type: Option<String>,

    // ── Inline tracking flags ───────────────────────────────────────
    in_text: bool,
    in_rpr: bool,
    /// True once the current rPr/endParaRPr applied its own Latin typeface, so
    /// a duplicate slot does not override the first declaration.
    rpr_applied_latin_typeface: bool,
    /// The corresponding first-declaration guard for the East Asian slot.
    rpr_applied_east_asian_typeface: bool,
    in_end_para_rpr: bool,
    in_text_line: bool,
    solid_fill_ctx: SolidFillCtx,
    /// Inside `<a:lnRef>` within `<p:style>` — for resolving fallback line color.
    in_style_ln_ref: bool,
    /// Inside `<a:fillRef>` within `<p:style>` — for resolving fallback fill color.
    in_style_fill_ref: bool,
    in_style_effect_ref: bool,
    /// Inside `<a:fontRef>` within `<p:style>` — for resolving fallback text color.
    in_style_font_ref: bool,

    // ── Picture state (`<p:pic>`) ───────────────────────────────────
    in_pic: bool,
    pic: PictureState,

    // ── Graphic frame state (`<p:graphicFrame>`) ────────────────────
    in_graphic_frame: bool,
    gf: GraphicFrameState,
}

impl<'a> SlideXmlParser<'a> {
    fn new(xml: &'a str, ctx: SlideParseContext<'a>) -> Self {
        Self {
            xml,
            ctx,

            skip_placeholders: false,
            placeholder_geometry: None,

            elements: Vec::new(),
            warnings: Vec::new(),

            in_shape: false,
            shape: ShapeState::default(),

            in_txbody: false,
            paragraphs: Vec::new(),
            text_box: PptxTextBoxSettings::default(),
            text_body_style_defaults: PptxTextBodyStyleDefaults::default(),

            in_para: false,
            para_style: ParagraphStyle::default(),
            para_level: 0,
            para_default_run_style: TextStyle::default(),
            para_end_run_style: TextStyle::default(),
            para_declares_end_para_rpr: false,
            para_bullet_definition: PptxBulletDefinition::default(),
            in_ln_spc: false,
            in_spc_bef: false,
            in_spc_aft: false,
            runs: Vec::new(),

            in_run: false,
            run_field_type: None,
            run_style: TextStyle::default(),
            run_text: String::new(),
            run_has_explicit_underline: false,
            run_marker_style_before_hyperlink: None,
            first_run_marker_style_override: None,

            in_text: false,
            in_rpr: false,
            rpr_applied_latin_typeface: false,
            rpr_applied_east_asian_typeface: false,
            in_end_para_rpr: false,
            in_text_line: false,
            solid_fill_ctx: SolidFillCtx::None,
            in_style_ln_ref: false,
            in_style_fill_ref: false,
            in_style_effect_ref: false,
            in_style_font_ref: false,

            in_pic: false,
            pic: PictureState::default(),

            in_graphic_frame: false,
            gf: GraphicFrameState::default(),
        }
    }

    /// Handle an `Event::Start` element by trying each domain sub-handler in
    /// the original dispatch order.
    fn handle_start(&mut self, reader: &mut Reader<&[u8]>, e: &BytesStart<'_>) {
        let _ = self.handle_start_frames_tables_groups(reader, e)
            || self.handle_start_shape_tree(reader, e)
            || self.handle_start_text_body(reader, e)
            || self.handle_start_fill_colors_and_style_refs(reader, e)
            || self.handle_start_picture(e);
    }

    /// Graphic frames, embedded tables, and shape groups.
    ///
    /// Returns `true` when the element was dispatched here. The sub-handlers
    /// preserve the original single-match arm order: each owns a contiguous
    /// slice of it, so guard overlap between slices keeps its old priority.
    fn handle_start_frames_tables_groups(
        &mut self,
        reader: &mut Reader<&[u8]>,
        e: &BytesStart<'_>,
    ) -> bool {
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
                if let Ok(mut table) = parse_pptx_table(
                    reader,
                    self.ctx.theme,
                    self.ctx.color_map,
                    self.ctx.table_styles,
                    self.ctx.default_text_size_pt,
                ) {
                    scale_pptx_table_geometry_to_frame(
                        &mut table,
                        emu_to_pt(self.gf.cx),
                        emu_to_pt(self.gf.cy),
                    );
                    // PowerPoint treats each `a:tr/@h` as a floor and grows a
                    // row whose cell content is taller. The frame still scales
                    // those floors up when it explicitly exceeds their sum,
                    // but it must not turn them into clipping tracks (#1253).
                    table.use_content_driven_row_heights = true;
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
                if let Ok((group_elems, group_warnings)) =
                    parse_group_shape(reader, self.xml, &self.ctx)
                {
                    self.elements.extend(group_elems);
                    self.warnings.extend(group_warnings);
                }
            }
            _ => return false,
        }
        true
    }

    /// Shape (`sp`/`cxnSp`) tree: geometry, fills, outline, placeholders (plus picture geometry arms that share this dispatch range).
    ///
    /// Returns `true` when the element was dispatched here. The sub-handlers
    /// preserve the original single-match arm order: each owns a contiguous
    /// slice of it, so guard overlap between slices keeps its old priority.
    fn handle_start_shape_tree(&mut self, reader: &mut Reader<&[u8]>, e: &BytesStart<'_>) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            b"sp" | b"cxnSp" if !self.in_shape && !self.in_pic => {
                self.in_shape = true;
                self.shape.reset();
                self.shape.depth = 1;
                self.in_txbody = false;
                self.paragraphs.clear();
                self.text_box = PptxTextBoxSettings::default();
            }
            b"sp" | b"cxnSp" if self.in_shape => {
                self.shape.depth += 1;
            }
            b"spPr" if self.in_shape && !self.in_txbody => {
                self.shape.in_sp_pr = true;
            }
            b"xfrm" if self.in_shape && self.shape.in_sp_pr => {
                self.shape.in_xfrm = true;
                self.shape.has_explicit_xfrm = true;
                if let Some(rot) = get_attr_i64(e, b"rot") {
                    self.shape.rotation_deg = Some(rot as f64 / 60_000.0);
                }
                self.shape.flip_h =
                    get_attr_str(e, b"flipH").is_some_and(|v| v == "1" || v == "true");
                self.shape.flip_v =
                    get_attr_str(e, b"flipV").is_some_and(|v| v == "1" || v == "true");
            }
            b"prstGeom" if self.in_pic && self.pic.in_sp_pr => {
                self.pic.prst_geom = get_attr_str(e, b"prst");
                self.pic.in_prst_geom = true;
            }
            // A picture can crop to a custom shape as well as a preset one
            // (issue #872).
            b"custGeom" if self.in_pic && self.pic.in_sp_pr => {
                // `<a:xfrm>` precedes the geometry in `<p:spPr>`, so the box a
                // guide formula measures against is already known.
                self.pic.custom_geometry = super::custom_geometry::parse_custom_geometry(
                    reader,
                    super::geometry_guides::ShapeExtent::new(
                        self.pic.cx as f64,
                        self.pic.cy as f64,
                    ),
                );
            }
            b"effectLst" if self.in_pic && self.pic.in_sp_pr => {
                self.pic.shadow = parse_effect_list(reader, self.ctx.theme, self.ctx.color_map);
            }
            b"gd" if self.in_pic && self.pic.in_prst_geom => {
                if self.pic.prst_adj.is_none()
                    && let Some(formula) = get_attr_str(e, b"fmla")
                    && let Some(value) = formula.strip_prefix("val ")
                    && let Ok(value) = value.trim().parse::<f64>()
                {
                    self.pic.prst_adj = Some(value / 100_000.0);
                }
            }
            b"prstGeom" if self.shape.in_sp_pr => {
                if let Some(prst) = get_attr_str(e, b"prst") {
                    self.shape.prst_geom = Some(prst);
                }
            }
            // Custom geometry is flattened to subpaths and drawn as one
            // even-odd path. A geometry that yields none still falls back to
            // a rectangle, so its fill renders as it did before
            // (issues #855, #866, #870).
            b"custGeom" if self.shape.in_sp_pr && self.shape.prst_geom.is_none() => {
                let geometry = super::custom_geometry::parse_custom_geometry_with_text_rect(
                    reader,
                    super::geometry_guides::ShapeExtent::new(
                        self.shape.cx as f64,
                        self.shape.cy as f64,
                    ),
                );
                self.shape.custom_geometry = geometry.subpaths;
                self.shape.custom_text_rect = geometry.text_rect;
                self.shape.prst_geom = Some("rect".to_string());
            }
            b"noFill" if self.shape.in_sp_pr && !self.shape.in_ln && !self.in_rpr => {
                self.shape.explicit_no_fill = true;
            }
            b"noFill" if self.shape.in_ln && !self.in_rpr => {
                self.shape.explicit_no_line = true;
            }
            b"solidFill" if self.shape.in_sp_pr && !self.shape.in_ln && !self.in_rpr => {
                self.solid_fill_ctx = SolidFillCtx::ShapeFill;
            }
            b"gradFill" if self.shape.in_sp_pr && !self.shape.in_ln && !self.in_rpr => {
                self.shape.gradient_fill =
                    parse_shape_gradient_fill(reader, self.ctx.theme, self.ctx.color_map);
                if let Some(ref gradient_fill) = self.shape.gradient_fill
                    && self.shape.fill.is_none()
                {
                    self.shape.fill = gradient_fill.stops.first().map(|stop| stop.color);
                }
            }
            b"pattFill" if self.shape.in_sp_pr && !self.shape.in_ln && !self.in_rpr => {
                self.shape.pattern_fill = get_attr_str(e, b"prst")
                    .as_deref()
                    .and_then(PatternPreset::from_ooxml)
                    .map(|preset| {
                        parse_shape_pattern_fill(reader, preset, self.ctx.theme, self.ctx.color_map)
                    });
            }
            b"effectLst" if self.shape.in_sp_pr && !self.shape.in_ln => {
                self.shape.has_direct_effect_lst = true;
                self.shape.shadow = parse_effect_list(reader, self.ctx.theme, self.ctx.color_map);
            }
            b"extLst" if self.shape.in_sp_pr && !self.in_txbody => {
                // Office extension payloads such as a16:hiddenLine are not visible shape
                // styling. If we parse nested fills here, they can overwrite the actual
                // shape fill, as seen on grouped icon ellipses that should stay white.
                crate::parser::xml_util::skip_element(reader, b"extLst");
            }
            b"ln" if self.shape.in_sp_pr => {
                self.shape.in_ln = true;
                self.shape.ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                self.shape.ln_dash_style = BorderLineStyle::Solid;
                self.shape.ln_join = None;
            }
            b"prstDash" if self.shape.in_ln => {
                self.shape.ln_dash_style = get_attr_str(e, b"val")
                    .as_deref()
                    .map(pptx_dash_to_border_style)
                    .unwrap_or(BorderLineStyle::Solid);
            }
            b"round" | b"bevel" | b"miter" if self.shape.in_ln => {
                self.shape.ln_join = drawingml_line_join(local.as_ref());
            }
            b"tailEnd" if self.shape.in_ln => {
                self.shape.tail_end = parse_arrow_head(get_attr_str(e, b"type").as_deref());
            }
            b"headEnd" if self.shape.in_ln => {
                self.shape.head_end = parse_arrow_head(get_attr_str(e, b"type").as_deref());
            }
            b"solidFill" if self.shape.in_ln => {
                self.solid_fill_ctx = SolidFillCtx::LineFill;
            }
            b"ph" if self.in_shape => {
                self.shape.has_placeholder = true;
                self.shape.ph_type = get_attr_str(e, b"type");
                self.shape.ph_idx = get_attr_str(e, b"idx");
            }
            b"ph" if self.in_pic => {
                self.pic.has_placeholder = true;
                self.pic.ph_type = get_attr_str(e, b"type");
                self.pic.ph_idx = get_attr_str(e, b"idx");
            }
            _ => return false,
        }
        true
    }

    /// Text body: paragraphs, spacing, bullets, runs, and run properties.
    ///
    /// Returns `true` when the element was dispatched here. The sub-handlers
    /// preserve the original single-match arm order: each owns a contiguous
    /// slice of it, so guard overlap between slices keeps its old priority.
    fn handle_start_text_body(&mut self, reader: &mut Reader<&[u8]>, e: &BytesStart<'_>) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            b"txBody" if self.in_shape => {
                self.in_txbody = true;
                self.text_body_style_defaults = if self.shape.has_placeholder {
                    // Placeholder text stacks the master txStyles bucket and
                    // the matching master/layout placeholder list styles.
                    self.placeholder_geometry
                        .map(|map| {
                            map.text_defaults(
                                self.shape.ph_type.as_deref(),
                                self.shape.ph_idx.as_deref(),
                            )
                        })
                        .unwrap_or_default()
                } else {
                    self.ctx.inherited_text_body_defaults.clone()
                };
                // The layout/master placeholder's `<a:bodyPr>` is the base the
                // slide's own then overrides attribute by attribute, so it has
                // to land before `<a:bodyPr>` is read a few events later.
                if self.shape.has_placeholder
                    && let Some(map) = self.placeholder_geometry
                {
                    map.body_props(self.shape.ph_type.as_deref(), self.shape.ph_idx.as_deref())
                        .apply_to(&mut self.text_box);
                }
                // Apply fontRef default text color from <p:style> to all text levels,
                // overriding inherited layout/master defaults.
                if let Some(color) = self.shape.style_font_color {
                    self.text_body_style_defaults.apply_default_color(color);
                }
                if let Some(ref font_family) = self.shape.style_font_family {
                    self.text_body_style_defaults
                        .apply_default_font_family(font_family);
                }
            }
            b"bodyPr" if self.in_shape && self.in_txbody => {
                extract_pptx_text_box_body_props(e, &mut self.text_box);
            }
            // Only `<a:normAutofit/>` shrinks text. `<a:spAutoFit/>` grows the
            // shape to the text and leaves the run's declared size alone
            // (ECMA-376 §21.1.2.1.2 / §21.1.2.1.3, issue #898).
            b"normAutofit" if self.in_shape && self.in_txbody => {
                extract_pptx_normal_autofit(e, &mut self.text_box);
            }
            b"lstStyle" if self.in_shape && self.in_txbody => {
                let local_defaults =
                    parse_pptx_list_style(reader, self.ctx.theme, self.ctx.color_map);
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
                self.para_declares_end_para_rpr = false;
                self.para_bullet_definition = self
                    .text_body_style_defaults
                    .bullet_for_level(self.para_level);
                self.in_ln_spc = false;
                self.runs.clear();
                self.first_run_marker_style_override = None;
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
            b"tabLst" if self.in_para && !self.in_run => {
                self.para_style.tab_stops = Some(Vec::new());
            }
            b"tab" if self.in_para && !self.in_run => {
                extract_pptx_tab_stop(e, &mut self.para_style);
            }
            b"spcBef" if self.in_para && !self.in_run => {
                self.in_spc_bef = true;
            }
            b"spcAft" if self.in_para && !self.in_run => {
                self.in_spc_aft = true;
            }
            b"spcPct" if self.in_ln_spc => {
                extract_pptx_line_spacing_pct(e, &mut self.para_style);
            }
            b"spcPts" if self.in_ln_spc => {
                extract_pptx_line_spacing_pts(e, &mut self.para_style);
            }
            b"spcPts" if self.in_spc_bef => {
                extract_pptx_space_points(e, &mut self.para_style.space_before);
                self.para_style.space_before_percent = None;
            }
            b"spcPts" if self.in_spc_aft => {
                extract_pptx_space_points(e, &mut self.para_style.space_after);
                self.para_style.space_after_percent = None;
            }
            b"spcPct" if self.in_spc_bef => {
                extract_pptx_space_percent(e, &mut self.para_style.space_before_percent);
                self.para_style.space_before = None;
            }
            b"spcPct" if self.in_spc_aft => {
                extract_pptx_space_percent(e, &mut self.para_style.space_after_percent);
                self.para_style.space_after = None;
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
                        resolve_theme_font(&typeface, self.ctx.theme),
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
                self.run_field_type = None;
                self.run_style = self.para_default_run_style.clone();
                self.run_text.clear();
                self.run_has_explicit_underline = false;
                self.run_marker_style_before_hyperlink = None;
            }
            // `<a:fld>` carries an optional `<a:rPr>` and `<a:t>` the way
            // `<a:r>` does, so it is read as a run whose text may be
            // substituted on close. `CT_TextField` also permits an `<a:pPr>`
            // that `CT_RegularTextRun` does not, and makes `<a:t>` optional;
            // neither matters here, since paragraph properties are read from
            // the enclosing `<a:p>` and an empty run is dropped.
            //
            // Leaving the element unparsed dropped the text and left the shape
            // empty, which then printed as a bare outlined rectangle (#540).
            b"fld" if self.in_para => {
                self.in_run = true;
                self.run_field_type = get_attr_str(e, b"type");
                self.run_style = self.para_default_run_style.clone();
                self.run_text.clear();
                self.run_has_explicit_underline = false;
                self.run_marker_style_before_hyperlink = None;
            }
            b"rPr" if self.in_run => {
                self.in_rpr = true;
                self.rpr_applied_latin_typeface = false;
                self.rpr_applied_east_asian_typeface = false;
                self.run_has_explicit_underline = get_attr_str(e, b"u").is_some();
                extract_rpr_attributes(e, &mut self.run_style);
            }
            b"endParaRPr" if self.in_para && !self.in_run => {
                self.in_end_para_rpr = true;
                self.rpr_applied_latin_typeface = false;
                self.rpr_applied_east_asian_typeface = false;
                self.para_end_run_style = self.para_default_run_style.clone();
                self.para_declares_end_para_rpr = true;
                extract_rpr_attributes(e, &mut self.para_end_run_style);
            }
            b"ln" if self.in_rpr || self.in_end_para_rpr => {
                self.in_text_line = true;
            }
            b"solidFill" if self.in_rpr && !self.in_text_line => {
                self.solid_fill_ctx = SolidFillCtx::RunFill;
            }
            b"solidFill" if self.in_end_para_rpr && !self.in_text_line => {
                self.solid_fill_ctx = SolidFillCtx::EndParaFill;
            }
            b"hlinkClick" if self.in_rpr => {
                if self.run_marker_style_before_hyperlink.is_none() {
                    self.run_marker_style_before_hyperlink = Some(self.run_style.clone());
                }
                apply_pptx_hyperlink_style(
                    &mut self.run_style,
                    self.run_has_explicit_underline,
                    self.ctx.theme,
                    self.ctx.color_map,
                );
            }
            _ => return false,
        }
        true
    }

    /// Solid-fill color elements and `<p:style>` lnRef/fillRef/fontRef fallbacks.
    ///
    /// Returns `true` when the element was dispatched here. The sub-handlers
    /// preserve the original single-match arm order: each owns a contiguous
    /// slice of it, so guard overlap between slices keeps its old priority.
    fn handle_start_fill_colors_and_style_refs(
        &mut self,
        reader: &mut Reader<&[u8]>,
        e: &BytesStart<'_>,
    ) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            b"blipFill" if self.shape.in_sp_pr && !self.shape.in_ln => {
                self.shape.in_blip_fill = true;
            }
            b"blip" if self.shape.in_blip_fill => {
                self.shape.blip_embed = get_attr_str(e, b"r:embed");
            }
            b"alphaModFix" if self.shape.in_blip_fill => {
                if let Some(alpha) = crate::parser::drawingml::parse_alpha_mod_fix(e) {
                    self.shape.blip_alpha = Some(alpha);
                }
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.solid_fill_ctx != SolidFillCtx::None => {
                let parsed = parse_color_from_start(reader, e, self.ctx.theme, self.ctx.color_map);
                apply_solid_fill_color(
                    self.solid_fill_ctx,
                    &parsed,
                    &mut self.shape,
                    &mut self.run_style,
                    &mut self.para_end_run_style,
                    &mut self.para_bullet_definition,
                    &mut self.pic,
                );
            }
            // Style-matrix ref colors (`<a:lnRef>`/`<a:fillRef>`/`<a:fontRef>`)
            // can carry shade/tint transforms, which arrive as Start events;
            // the Empty-event arms below would miss them.
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.in_style_ln_ref => {
                let parsed = parse_color_from_start(reader, e, self.ctx.theme, self.ctx.color_map);
                self.shape.style_ln_color = parsed.color;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.in_style_fill_ref => {
                let parsed = parse_color_from_start(reader, e, self.ctx.theme, self.ctx.color_map);
                self.shape.style_fill_color = parsed.color;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.in_style_font_ref => {
                let parsed = parse_color_from_start(reader, e, self.ctx.theme, self.ctx.color_map);
                self.shape.style_font_color = parsed.color;
            }
            // `<a:lnRef>` inside `<p:style>` provides fallback line color.
            b"lnRef" if self.in_shape && !self.shape.in_sp_pr && !self.in_txbody => {
                self.in_style_ln_ref = true;
                self.shape.style_ln_idx = get_attr_str(e, b"idx")
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|idx| *idx > 0);
            }
            // `<a:fillRef>` inside `<p:style>` provides fallback fill color.
            b"fillRef" if self.in_shape && !self.shape.in_sp_pr && !self.in_txbody => {
                self.in_style_fill_ref = true;
                self.shape.style_fill_idx = get_attr_str(e, b"idx")
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|idx| *idx > 0);
            }
            // `<a:effectRef>` inside `<p:style>` names a theme effect style.
            b"effectRef" if self.in_shape && !self.shape.in_sp_pr && !self.in_txbody => {
                self.in_style_effect_ref = true;
                self.shape.style_effect_idx =
                    get_attr_str(e, b"idx").and_then(|value| value.parse::<i64>().ok());
            }
            // `<a:fontRef>` inside `<p:style>` provides fallback text color.
            b"fontRef" if self.in_shape && !self.shape.in_sp_pr && !self.in_txbody => {
                self.in_style_font_ref = true;
                self.shape.style_font_family = match get_attr_str(e, b"idx").as_deref() {
                    Some("major") => self.ctx.theme.major_font.clone(),
                    Some("minor") => self.ctx.theme.minor_font.clone(),
                    _ => None,
                };
            }
            b"t" if self.in_run => {
                self.in_text = true;
            }
            _ => return false,
        }
        true
    }

    /// Picture (`pic`) tree: blip fills, crops, outline, and image layers.
    ///
    /// Returns `true` when the element was dispatched here. The sub-handlers
    /// preserve the original single-match arm order: each owns a contiguous
    /// slice of it, so guard overlap between slices keeps its old priority.
    fn handle_start_picture(&mut self, e: &BytesStart<'_>) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            b"pic" if !self.in_shape && !self.in_pic => {
                self.in_pic = true;
                self.pic.reset();
            }
            b"spPr" if self.in_pic => {
                self.pic.in_sp_pr = true;
            }
            b"xfrm" if self.in_pic && self.pic.in_sp_pr => {
                self.pic.in_xfrm = true;
                self.pic.has_explicit_xfrm = true;
                // The shape path reads this a few arms up; the picture path
                // never did, so a rotated picture drew upright (issue #682).
                if let Some(rot) = get_attr_i64(e, b"rot") {
                    self.pic.rotation_deg = Some(rot as f64 / 60_000.0);
                }
                self.pic.flip_h =
                    get_attr_str(e, b"flipH").is_some_and(|v| v == "1" || v == "true");
                self.pic.flip_v =
                    get_attr_str(e, b"flipV").is_some_and(|v| v == "1" || v == "true");
            }
            b"ln" if self.in_pic && self.pic.in_sp_pr => {
                self.pic.in_ln = true;
                self.pic.ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                self.pic.ln_dash_style = BorderLineStyle::Solid;
                self.pic.ln_join = None;
            }
            b"solidFill" if self.in_pic && self.pic.in_ln => {
                self.solid_fill_ctx = SolidFillCtx::PicLineFill;
            }
            b"prstDash" if self.in_pic && self.pic.in_ln => {
                self.pic.ln_dash_style = get_attr_str(e, b"val")
                    .as_deref()
                    .map(pptx_dash_to_border_style)
                    .unwrap_or(BorderLineStyle::Solid);
            }
            b"round" | b"bevel" | b"miter" if self.in_pic && self.pic.in_ln => {
                self.pic.ln_join = drawingml_line_join(local.as_ref());
            }
            b"blipFill" if self.in_pic => {}
            b"blip" if self.in_pic => {
                self.pic.blip_embed = get_attr_str(e, b"r:embed");
            }
            b"alphaModFix" if self.in_pic => {
                if let Some(alpha) = crate::parser::drawingml::parse_alpha_mod_fix(e) {
                    self.pic.blip_alpha = Some(alpha);
                }
            }
            b"svgBlip" if self.in_pic => {
                self.pic.svg_blip_embed = get_attr_str(e, b"r:embed");
            }
            b"liveFeedProps" if self.in_pic => {
                self.pic.has_live_feed_properties = true;
            }
            b"imgLayer" if self.in_pic => {
                if let Some(rid) = get_attr_str(e, b"r:embed") {
                    self.pic.img_layer_embeds.push(rid);
                }
            }
            b"srcRect" if self.in_pic => {
                self.pic.crop = parse_src_rect(e);
            }
            _ => return false,
        }
        true
    }

    /// Handle an `Event::Empty` element by trying each domain sub-handler in
    /// the original dispatch order.
    fn handle_empty(&mut self, e: &BytesStart<'_>) {
        let _ = self.handle_empty_geometry_and_picture(e)
            || self.handle_empty_shape_props(e)
            || self.handle_empty_fill_colors_and_style_refs(e)
            || self.handle_empty_text_body(e);
    }

    /// Transform offsets/extents and self-closing picture attributes.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_empty_geometry_and_picture(&mut self, e: &BytesStart<'_>) -> bool {
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
            b"alphaModFix" if self.in_pic => {
                if let Some(alpha) = crate::parser::drawingml::parse_alpha_mod_fix(e) {
                    self.pic.blip_alpha = Some(alpha);
                }
            }
            b"svgBlip" if self.in_pic => {
                self.pic.svg_blip_embed = get_attr_str(e, b"r:embed");
            }
            b"liveFeedProps" if self.in_pic => {
                self.pic.has_live_feed_properties = true;
            }
            b"imgLayer" if self.in_pic => {
                if let Some(rid) = get_attr_str(e, b"r:embed") {
                    self.pic.img_layer_embeds.push(rid);
                }
            }
            b"srcRect" if self.in_pic => {
                self.pic.crop = parse_src_rect(e);
            }
            b"prstDash" if self.in_pic && self.pic.in_ln => {
                self.pic.ln_dash_style = get_attr_str(e, b"val")
                    .as_deref()
                    .map(pptx_dash_to_border_style)
                    .unwrap_or(BorderLineStyle::Solid);
            }
            b"round" | b"bevel" | b"miter" if self.in_pic && self.pic.in_ln => {
                self.pic.ln_join = drawingml_line_join(local.as_ref());
            }
            _ => return false,
        }
        true
    }

    /// Self-closing placeholder, body, geometry, and outline properties.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_empty_shape_props(&mut self, e: &BytesStart<'_>) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            // Handle self-closing <p:ph type="..."/> (placeholder marker).
            b"ph" if self.in_shape => {
                self.shape.has_placeholder = true;
                self.shape.ph_type = get_attr_str(e, b"type");
                self.shape.ph_idx = get_attr_str(e, b"idx");
            }
            b"ph" if self.in_pic => {
                self.pic.has_placeholder = true;
                self.pic.ph_type = get_attr_str(e, b"type");
                self.pic.ph_idx = get_attr_str(e, b"idx");
            }
            // Handle self-closing <a:bodyPr anchor="ctr"/> (no child elements).
            b"bodyPr" if self.in_shape && self.in_txbody => {
                extract_pptx_text_box_body_props(e, &mut self.text_box);
            }
            // Only `<a:normAutofit/>` shrinks text. `<a:spAutoFit/>` grows the
            // shape to the text and leaves the run's declared size alone
            // (ECMA-376 §21.1.2.1.2 / §21.1.2.1.3, issue #898).
            b"normAutofit" if self.in_shape && self.in_txbody => {
                extract_pptx_normal_autofit(e, &mut self.text_box);
            }
            b"prstGeom" if self.in_pic && self.pic.in_sp_pr => {
                self.pic.prst_geom = get_attr_str(e, b"prst");
            }
            b"gd" if self.in_pic && self.pic.in_prst_geom => {
                if self.pic.prst_adj.is_none()
                    && let Some(formula) = get_attr_str(e, b"fmla")
                    && let Some(value) = formula.strip_prefix("val ")
                    && let Ok(value) = value.trim().parse::<f64>()
                {
                    self.pic.prst_adj = Some(value / 100_000.0);
                }
            }
            // `<a:effectLst/>` carries no effect, but stating it is how a
            // shape switches its theme effect off, so it still counts as the
            // shape having spoken for itself (issue #740).
            b"effectLst" if self.shape.in_sp_pr && !self.shape.in_ln => {
                self.shape.has_direct_effect_lst = true;
            }
            b"prstGeom" if self.shape.in_sp_pr => {
                if let Some(prst) = get_attr_str(e, b"prst") {
                    self.shape.prst_geom = Some(prst);
                }
            }
            // Self-closing `<a:custGeom/>`: there is no `a:pathLst` to flatten,
            // so only the rectangle fallback applies. The start-tag handler is
            // the one that reads a path (issue #855).
            b"custGeom" if self.shape.in_sp_pr && self.shape.prst_geom.is_none() => {
                self.shape.prst_geom = Some("rect".to_string());
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
            b"round" | b"bevel" | b"miter" if self.shape.in_ln => {
                self.shape.ln_join = drawingml_line_join(local.as_ref());
            }
            b"tailEnd" if self.shape.in_ln => {
                self.shape.tail_end = parse_arrow_head(get_attr_str(e, b"type").as_deref());
            }
            b"headEnd" if self.shape.in_ln => {
                self.shape.head_end = parse_arrow_head(get_attr_str(e, b"type").as_deref());
            }
            // Preset adjustment values, in `<a:avLst>` document order.
            b"gd" if self.in_shape && self.shape.in_sp_pr => {
                if let Some(val) = get_attr_str(e, b"fmla")
                    .as_deref()
                    .and_then(|f| f.strip_prefix("val "))
                    .and_then(|s| s.parse::<f64>().ok())
                {
                    self.shape.adj_values.push(val);
                }
            }
            // `<a:noFill/>` inside `<p:spPr>` (not inside `<a:ln>`) explicitly disables fill.
            b"noFill" if self.shape.in_sp_pr && !self.shape.in_ln => {
                self.shape.explicit_no_fill = true;
            }
            b"noFill" if self.shape.in_ln => {
                self.shape.explicit_no_line = true;
            }
            _ => return false,
        }
        true
    }

    /// Self-closing solid-fill color elements and style-ref colors.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_empty_fill_colors_and_style_refs(&mut self, e: &BytesStart<'_>) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            b"blip" if self.shape.in_blip_fill => {
                self.shape.blip_embed = get_attr_str(e, b"r:embed");
            }
            b"alphaModFix" if self.shape.in_blip_fill => {
                if let Some(alpha) = crate::parser::drawingml::parse_alpha_mod_fix(e) {
                    self.shape.blip_alpha = Some(alpha);
                }
            }
            b"srcRect" if self.shape.in_blip_fill => {
                self.shape.blip_crop = parse_src_rect(e);
            }
            b"fontRef" if self.in_shape && !self.shape.in_sp_pr && !self.in_txbody => {
                self.shape.style_font_family = match get_attr_str(e, b"idx").as_deref() {
                    Some("major") => self.ctx.theme.major_font.clone(),
                    Some("minor") => self.ctx.theme.minor_font.clone(),
                    _ => None,
                };
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.in_style_font_ref => {
                let parsed = parse_color_from_empty(e, self.ctx.theme, self.ctx.color_map);
                self.shape.style_font_color = parsed.color;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.in_style_fill_ref => {
                let parsed = parse_color_from_empty(e, self.ctx.theme, self.ctx.color_map);
                self.shape.style_fill_color = parsed.color;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.in_style_ln_ref => {
                let parsed = parse_color_from_empty(e, self.ctx.theme, self.ctx.color_map);
                self.shape.style_ln_color = parsed.color;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.in_style_effect_ref => {
                let parsed = parse_color_from_empty(e, self.ctx.theme, self.ctx.color_map);
                self.shape.style_effect_color = parsed.color;
            }
            b"srgbClr" | b"schemeClr" | b"sysClr" if self.solid_fill_ctx != SolidFillCtx::None => {
                let parsed = parse_color_from_empty(e, self.ctx.theme, self.ctx.color_map);
                apply_solid_fill_color(
                    self.solid_fill_ctx,
                    &parsed,
                    &mut self.shape,
                    &mut self.run_style,
                    &mut self.para_end_run_style,
                    &mut self.para_bullet_definition,
                    &mut self.pic,
                );
            }
            _ => return false,
        }
        true
    }

    /// Self-closing paragraph, bullet, run-property, and typeface elements.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_empty_text_body(&mut self, e: &BytesStart<'_>) -> bool {
        let local = e.local_name();
        match local.as_ref() {
            b"rPr" if self.in_run => {
                self.run_has_explicit_underline = get_attr_str(e, b"u").is_some();
                extract_rpr_attributes(e, &mut self.run_style);
            }
            b"endParaRPr" if self.in_para && !self.in_run => {
                self.para_end_run_style = self.para_default_run_style.clone();
                self.para_declares_end_para_rpr = true;
                extract_rpr_attributes(e, &mut self.para_end_run_style);
            }
            b"ln" if self.in_rpr || self.in_end_para_rpr => {
                self.in_text_line = true;
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
            b"tabLst" if self.in_para && !self.in_run => {
                self.para_style.tab_stops = Some(Vec::new());
            }
            b"tab" if self.in_para && !self.in_run => {
                extract_pptx_tab_stop(e, &mut self.para_style);
            }
            b"spcBef" if self.in_para && !self.in_run => {
                self.in_spc_bef = true;
            }
            b"spcAft" if self.in_para && !self.in_run => {
                self.in_spc_aft = true;
            }
            b"spcPct" if self.in_ln_spc => {
                extract_pptx_line_spacing_pct(e, &mut self.para_style);
            }
            b"spcPts" if self.in_ln_spc => {
                extract_pptx_line_spacing_pts(e, &mut self.para_style);
            }
            b"spcPts" if self.in_spc_bef => {
                extract_pptx_space_points(e, &mut self.para_style.space_before);
                self.para_style.space_before_percent = None;
            }
            b"spcPts" if self.in_spc_aft => {
                extract_pptx_space_points(e, &mut self.para_style.space_after);
                self.para_style.space_after_percent = None;
            }
            b"spcPct" if self.in_spc_bef => {
                extract_pptx_space_percent(e, &mut self.para_style.space_before_percent);
                self.para_style.space_before = None;
            }
            b"spcPct" if self.in_spc_aft => {
                extract_pptx_space_percent(e, &mut self.para_style.space_after_percent);
                self.para_style.space_after = None;
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
                        resolve_theme_font(&typeface, self.ctx.theme),
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
            b"hlinkClick" if self.in_rpr => {
                if self.run_marker_style_before_hyperlink.is_none() {
                    self.run_marker_style_before_hyperlink = Some(self.run_style.clone());
                }
                apply_pptx_hyperlink_style(
                    &mut self.run_style,
                    self.run_has_explicit_underline,
                    self.ctx.theme,
                    self.ctx.color_map,
                );
            }
            b"latin" if self.in_rpr => {
                if !self.rpr_applied_latin_typeface {
                    self.run_style.font_family = None;
                }
                apply_typeface_to_style(e, &mut self.run_style, self.ctx.theme);
                self.rpr_applied_latin_typeface |= self.run_style.font_family.is_some();
            }
            b"ea" if self.in_rpr => {
                if !self.rpr_applied_east_asian_typeface {
                    self.run_style.east_asian_font_family = None;
                }
                apply_typeface_to_style(e, &mut self.run_style, self.ctx.theme);
                self.rpr_applied_east_asian_typeface |=
                    self.run_style.east_asian_font_family.is_some();
            }
            b"cs" if self.in_rpr => {}
            b"latin" if self.in_end_para_rpr => {
                if !self.rpr_applied_latin_typeface {
                    self.para_end_run_style.font_family = None;
                }
                apply_typeface_to_style(e, &mut self.para_end_run_style, self.ctx.theme);
                self.rpr_applied_latin_typeface |= self.para_end_run_style.font_family.is_some();
            }
            b"ea" if self.in_end_para_rpr => {
                if !self.rpr_applied_east_asian_typeface {
                    self.para_end_run_style.east_asian_font_family = None;
                }
                apply_typeface_to_style(e, &mut self.para_end_run_style, self.ctx.theme);
                self.rpr_applied_east_asian_typeface |=
                    self.para_end_run_style.east_asian_font_family.is_some();
            }
            b"cs" if self.in_end_para_rpr => {}
            _ => return false,
        }
        true
    }

    /// Handle an `Event::Text` element.
    fn handle_text(&mut self, text: &str) {
        if self.in_text {
            self.run_text.push_str(text);
        }
    }

    /// Handle an `Event::End` element by trying each domain sub-handler in
    /// the original dispatch order.
    fn handle_end(&mut self, local_name: &[u8]) {
        let _ = self.handle_end_shape(local_name)
            || self.handle_end_text_body(local_name)
            || self.handle_end_fill_and_style_refs(local_name)
            || self.handle_end_picture_and_frame(local_name);
    }

    /// Shape close (finalizes the accumulated shape) and shape scope flags.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_end_shape(&mut self, local_name: &[u8]) -> bool {
        match local_name {
            b"sp" | b"cxnSp" if self.in_shape => {
                self.shape.depth -= 1;
                if self.shape.depth == 0 {
                    // Skip placeholder shapes when parsing master/layout layers.
                    // Placeholder content is only visible when the slide itself
                    // overrides it; master/layout placeholder text (e.g.
                    // "마스터 제목 스타일 편집") should never be rendered.
                    if self.shape.has_placeholder
                        && !self.shape.has_explicit_xfrm
                        && let Some(geometry) = self.placeholder_geometry.and_then(|map| {
                            map.lookup(self.shape.ph_type.as_deref(), self.shape.ph_idx.as_deref())
                        })
                    {
                        self.shape.x = geometry.x;
                        self.shape.y = geometry.y;
                        self.shape.cx = geometry.cx;
                        self.shape.cy = geometry.cy;
                        self.shape.rotation_deg = geometry.rotation_deg;
                    }
                    // A `<p:style>` effect reference fills in only where the
                    // shape stated no effect of its own (issues #740, #1298).
                    // Resolved here rather than at the `<a:effectRef>` itself
                    // so it does not depend on spPr preceding p:style in the
                    // file.
                    if !self.shape.has_direct_effect_lst
                        && let Some(effect_idx) = self.shape.style_effect_idx.take()
                    {
                        if self.shape.shadow.is_none() {
                            self.shape.shadow = resolve_effect_ref(
                                effect_idx,
                                self.shape.style_effect_color.take(),
                                self.ctx.theme,
                                self.ctx.color_map,
                            );
                        }
                        self.shape.top_bevel =
                            resolve_effect_ref_top_bevel(effect_idx, self.ctx.theme);
                    }
                    // A fill style can be a gradient rather than the flat
                    // child color. Resolve it only when spPr supplied no fill
                    // of its own, regardless of whether style precedes spPr.
                    if self.shape.blip_embed.is_none()
                        && self.shape.fill.is_none()
                        && self.shape.gradient_fill.is_none()
                        && self.shape.pattern_fill.is_none()
                        && !self.shape.explicit_no_fill
                        && let Some(fill_idx) = self.shape.style_fill_idx.take()
                        && let Some((color, gradient)) = resolve_fill_ref(
                            fill_idx,
                            self.shape.style_fill_color,
                            self.ctx.theme,
                            self.ctx.color_map,
                        )
                    {
                        self.shape.style_fill_color = color.or(self.shape.style_fill_color);
                        self.shape.gradient_fill = gradient;
                    }
                    // A slide placeholder inherits its shape properties from
                    // the layout's matching copy, so a template's colour band
                    // behind a title is declared there while the slide carries
                    // only the text. The layout copy is not drawn — its prompt
                    // text would come with it — so the fill has to reach the
                    // slide's own shape, where it also lands at the right
                    // place in z-order (issue #856).
                    if self.shape.has_placeholder
                        && self.shape.blip_embed.is_none()
                        && self.shape.fill.is_none()
                        && self.shape.gradient_fill.is_none()
                        && self.shape.pattern_fill.is_none()
                        && !self.shape.explicit_no_fill
                        && let Some(map) = self.placeholder_geometry
                        && let Some(inherited) = map.lookup_fill(
                            self.shape.ph_type.as_deref(),
                            self.shape.ph_idx.as_deref(),
                        )
                    {
                        self.shape.fill = Some(inherited.color);
                        if self.shape.opacity.is_none() {
                            self.shape.opacity = inherited.opacity;
                        }
                    }
                    // The band inherited above is not a rectangle unless the
                    // layout says so: a template panel's `<a:custGeom>` (or
                    // preset) travels the same chain as its fill, so the
                    // colour paints the declared path rather than the
                    // placeholder's bounding box (issue #1029). Gated on the
                    // shape painting something, so a text-only placeholder
                    // does not grow an invisible background element.
                    let paints_something: bool = self.shape.fill.is_some()
                        || self.shape.gradient_fill.is_some()
                        || self.shape.pattern_fill.is_some()
                        || self.shape.blip_embed.is_some()
                        || (!self.shape.explicit_no_fill && self.shape.style_fill_color.is_some())
                        || (!self.shape.explicit_no_line
                            && (self.shape.ln_color.is_some()
                                || self.shape.style_ln_color.is_some()));
                    if self.shape.has_placeholder
                        && self.shape.prst_geom.is_none()
                        && self.shape.custom_geometry.is_empty()
                        && paints_something
                        && let Some(map) = self.placeholder_geometry
                        && let Some(inherited) = map.lookup_shape_geometry(
                            self.shape.ph_type.as_deref(),
                            self.shape.ph_idx.as_deref(),
                        )
                    {
                        self.shape.custom_geometry = inherited.subpaths.to_vec();
                        // Mirror the slide-side custGeom convention: a preset
                        // name when the layer stated one, otherwise the
                        // rectangle fallback that guards a degenerate path.
                        self.shape.prst_geom = Some(inherited.preset.unwrap_or("rect").to_string());
                    }
                    if !(self.skip_placeholders && self.shape.has_placeholder) {
                        self.elements.extend(finalize_shape(
                            &mut self.shape,
                            &mut self.paragraphs,
                            self.text_box,
                            &self.ctx.theme.line_styles,
                            self.ctx.images,
                            self.ctx.warning_context,
                            &mut self.warnings,
                        ));
                    }
                    self.in_shape = false;
                }
            }
            b"spPr" if self.shape.in_sp_pr => {
                self.shape.in_sp_pr = false;
            }
            b"blipFill" if self.shape.in_blip_fill => {
                self.shape.in_blip_fill = false;
            }
            b"xfrm" if self.shape.in_xfrm => {
                self.shape.in_xfrm = false;
            }
            b"ln" if self.shape.in_ln => {
                self.shape.in_ln = false;
            }
            _ => return false,
        }
        true
    }

    /// Text-body close: paragraph/run assembly and paragraph scope flags.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_end_text_body(&mut self, local_name: &[u8]) -> bool {
        match local_name {
            b"txBody" if self.in_txbody => {
                self.in_txbody = false;
            }
            b"p" if self.in_para => {
                let resolved_list_marker = resolve_pptx_list_marker(
                    &self.para_bullet_definition,
                    self.para_level,
                    &self.runs,
                    self.first_run_marker_style_override.as_ref(),
                    &self.para_end_run_style,
                    &self.para_default_run_style,
                );
                let mut paragraph_runs = std::mem::take(&mut self.runs);
                let mut paragraph_style: ParagraphStyle = self.para_style.clone();
                paragraph_style.paragraph_mark_font_family = pptx_paragraph_mark_font_family(
                    &self.para_end_run_style,
                    self.ctx.theme,
                    PptxParagraphMark::for_paragraph(
                        self.para_declares_end_para_rpr,
                        &paragraph_runs,
                    ),
                );
                insert_hangul_kinsoku_break_markers(&mut paragraph_runs);
                // `a:tab pos` is measured from the text origin — the box edge
                // plus `lIns` — not from the box edge itself: the native
                // export of customGeo.pptx page 46 lands its value run at
                // exactly text_origin + pos, 7.2pt (one default lIns) right
                // of where the former inset subtraction put it (issue #785).
                self.paragraphs.push(PptxParagraphEntry {
                    paragraph: Paragraph {
                        style: paragraph_style,
                        runs: paragraph_runs,
                    },
                    list_marker: resolved_list_marker,
                    paragraph_mark_font_size_pt: self.para_end_run_style.font_size,
                });
                self.in_para = false;
            }
            b"r" | b"fld" if self.in_run => {
                if self.run_field_type.as_deref() == Some("slidenum") {
                    // The cached `<a:t>` is whatever PowerPoint last drew, so
                    // it goes stale when slides are reordered. Every other
                    // field type keeps its cache, which is both PowerPoint's
                    // own fallback and the only deterministic reading of a
                    // date field.
                    self.run_text = self.ctx.slide_number.to_string();
                }
                if !self.run_text.is_empty() {
                    if self.runs.is_empty() {
                        self.first_run_marker_style_override =
                            self.run_marker_style_before_hyperlink.clone();
                    }
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
                self.run_field_type = None;
            }
            b"rPr" if self.in_rpr => {
                self.in_rpr = false;
            }
            b"endParaRPr" if self.in_end_para_rpr => {
                self.in_end_para_rpr = false;
            }
            b"ln" if self.in_text_line => {
                self.in_text_line = false;
            }
            b"lnSpc" if self.in_ln_spc => {
                self.in_ln_spc = false;
            }
            b"spcBef" if self.in_spc_bef => {
                self.in_spc_bef = false;
            }
            b"spcAft" if self.in_spc_aft => {
                self.in_spc_aft = false;
            }
            _ => return false,
        }
        true
    }

    /// Solid-fill and style-ref scope closes.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_end_fill_and_style_refs(&mut self, local_name: &[u8]) -> bool {
        match local_name {
            b"solidFill" if self.solid_fill_ctx != SolidFillCtx::None => {
                self.solid_fill_ctx = SolidFillCtx::None;
            }
            b"lnRef" if self.in_style_ln_ref => {
                self.in_style_ln_ref = false;
            }
            b"effectRef" if self.in_style_effect_ref => {
                self.in_style_effect_ref = false;
            }
            b"fillRef" if self.in_style_fill_ref => {
                self.in_style_fill_ref = false;
            }
            b"fontRef" if self.in_style_font_ref => {
                self.in_style_font_ref = false;
            }
            b"t" if self.in_text => {
                self.in_text = false;
            }
            _ => return false,
        }
        true
    }

    /// Picture close (finalizes the image element) and graphic-frame flags.
    ///
    /// Returns `true` when the element was dispatched here (same contiguous
    /// arm-order preservation as the `handle_start_*` sub-handlers).
    fn handle_end_picture_and_frame(&mut self, local_name: &[u8]) -> bool {
        match local_name {
            b"pic" if self.in_pic => {
                if self.pic.has_placeholder
                    && !self.pic.has_explicit_xfrm
                    && let Some(geometry) = self.placeholder_geometry.and_then(|map| {
                        map.lookup(self.pic.ph_type.as_deref(), self.pic.ph_idx.as_deref())
                    })
                {
                    self.pic.x = geometry.x;
                    self.pic.y = geometry.y;
                    self.pic.cx = geometry.cx;
                    self.pic.cy = geometry.cy;
                    self.pic.rotation_deg = geometry.rotation_deg;
                }
                // A template's picture placeholder is often an `ellipse` or a
                // custom path: PowerPoint crops every photo dropped into it to
                // that outline. The outline inherits on its own, independent of
                // the frame above, so a picture that restates its `<a:xfrm>`
                // still takes the layout's clip. Without it the photo's full
                // rectangular canvas paints around the circle (issue #1373).
                if self.pic.has_placeholder
                    && self.pic.prst_geom.is_none()
                    && self.pic.custom_geometry.is_empty()
                    && let Some(map) = self.placeholder_geometry
                    && let Some(inherited) = map.lookup_shape_geometry(
                        self.pic.ph_type.as_deref(),
                        self.pic.ph_idx.as_deref(),
                    )
                {
                    self.pic.custom_geometry = inherited.subpaths.to_vec();
                    self.pic.prst_geom = Some(inherited.preset.unwrap_or("rect").to_string());
                }
                let (element, picture_warnings) =
                    finalize_picture(&self.pic, self.ctx.images, self.ctx.warning_context);
                self.warnings.extend(picture_warnings);
                if let Some(element) = element {
                    self.elements.push(element);
                }
                self.in_pic = false;
            }
            b"spPr" if self.in_pic && self.pic.in_sp_pr => {
                self.pic.in_sp_pr = false;
            }
            b"prstGeom" if self.in_pic && self.pic.in_prst_geom => {
                self.pic.in_prst_geom = false;
            }
            b"ln" if self.in_pic && self.pic.in_ln => {
                self.pic.in_ln = false;
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
            _ => return false,
        }
        true
    }

    /// Consume the parser and return the accumulated results.
    fn finish(self) -> (Vec<FixedElement>, Vec<ConvertWarning>) {
        (self.elements, self.warnings)
    }
}

// ── Main parse function ─────────────────────────────────────────────────

/// Parse a slide XML to extract positioned elements (text boxes, shapes, images).
pub(super) fn parse_slide_xml<'a>(
    xml: &'a str,
    ctx: &SlideParseContext<'a>,
    placeholder_geometry: Option<&'a PlaceholderGeometryMap>,
) -> Result<(Vec<FixedElement>, Vec<ConvertWarning>), ConvertError> {
    parse_slide_xml_inner(xml, ctx, false, placeholder_geometry)
}

fn parse_slide_xml_inner<'a>(
    xml: &'a str,
    ctx: &SlideParseContext<'a>,
    skip_placeholders: bool,
    placeholder_geometry: Option<&'a PlaceholderGeometryMap>,
) -> Result<(Vec<FixedElement>, Vec<ConvertWarning>), ConvertError> {
    let mut reader = Reader::from_str(xml);
    let mut parser = SlideXmlParser::new(xml, *ctx);
    parser.skip_placeholders = skip_placeholders;
    parser.placeholder_geometry = placeholder_geometry;

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
                if let Some(text) = crate::parser::xml_util::decode_general_ref(reference) {
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

#[cfg(test)]
mod picture_mask_tests {
    use super::{apply_path_mask, clip_svg_to_path, point_is_inside_even_odd};
    use crate::config::ConvertOptions;
    use crate::ir::{FixedElementKind, ImageCrop, ImageFormat, Page};
    use crate::parser::Parser as _;

    const SVG: &str = r#"<svg width="100" height="50" viewBox="0 0 100 50" xmlns="http://www.w3.org/2000/svg"><rect width="100" height="50"/></svg>"#;

    /// An SVG takes its `a:custGeom` as a `<clipPath>` rather than an alpha
    /// mask, since there is no raster to mask (issue #897).
    ///
    /// The subpaths arrive normalised to 0..1 of the picture box, so they are
    /// scaled onto the root's own viewBox to become user-space coordinates.
    #[test]
    fn an_svg_takes_its_custom_geometry_as_a_clip_path() {
        let triangle = vec![crate::ir::Subpath::closed_outline(vec![
            (0.0, 0.0),
            (1.0, 0.0),
            (0.0, 1.0),
        ])];
        let out = clip_svg_to_path(SVG.as_bytes(), &triangle, None).expect("the SVG is clipped");
        let text = String::from_utf8(out).expect("still an SVG");

        assert!(text.contains("<clipPath"), "a clipPath is added: {text}");
        assert!(
            text.contains(r#"clipPathUnits="userSpaceOnUse""#),
            "the path is in the root's own units: {text}"
        );
        // 0..1 across a 100 x 50 viewBox.
        assert!(
            text.contains("M 0 0 L 100 0 L 0 50 Z"),
            "the triangle is scaled onto the viewBox: {text}"
        );
        assert!(
            text.contains(r#"clip-rule="evenodd""#),
            "an inner boundary must carve a hole, as the raster mask does: {text}"
        );
        assert!(
            text.contains(r#"<rect width="100" height="50"/>"#),
            "the drawing itself is untouched: {text}"
        );
        assert!(
            text.trim_end().ends_with("</svg>"),
            "the document stays well formed: {text}"
        );
    }

    /// A root that states no usable viewBox is left alone rather than clipped
    /// against guessed units.
    #[test]
    fn an_svg_without_a_view_box_is_not_clipped() {
        const NO_BOX: &str = r#"<svg width="10" height="10"><rect width="10" height="10"/></svg>"#;
        let triangle = vec![crate::ir::Subpath::closed_outline(vec![
            (0.0, 0.0),
            (1.0, 0.0),
            (0.0, 1.0),
        ])];
        assert!(clip_svg_to_path(NO_BOX.as_bytes(), &triangle, None).is_none());
    }

    /// The unit square, and a smaller square inside it.
    fn frame() -> Vec<crate::ir::Subpath> {
        vec![
            crate::ir::Subpath::closed_outline(vec![
                (0.0, 0.0),
                (1.0, 0.0),
                (1.0, 1.0),
                (0.0, 1.0),
            ]),
            crate::ir::Subpath::closed_outline(vec![
                (0.25, 0.25),
                (0.75, 0.25),
                (0.75, 0.75),
                (0.25, 0.75),
            ]),
        ]
    }

    /// A picture's custom crop keeps what the path encloses and clears the
    /// rest, so a point outside every subpath is masked away (issue #872).
    #[test]
    fn a_point_outside_the_only_subpath_is_masked() {
        let triangle = vec![crate::ir::Subpath::closed_outline(vec![
            (0.0, 0.0),
            (1.0, 0.0),
            (0.0, 1.0),
        ])];
        assert!(point_is_inside_even_odd(&triangle, 0.1, 0.1));
        assert!(!point_is_inside_even_odd(&triangle, 0.9, 0.9));
    }

    /// An inner boundary carves a hole rather than adding to the fill, so the
    /// middle of a frame is masked away too (issues #870, #872).
    #[test]
    fn the_inside_of_a_frame_is_a_hole() {
        assert!(
            point_is_inside_even_odd(&frame(), 0.1, 0.5),
            "the band between the two rings is kept"
        );
        assert!(
            !point_is_inside_even_odd(&frame(), 0.5, 0.5),
            "the centre falls inside both rings, so it is a hole"
        );
        assert!(
            !point_is_inside_even_odd(&frame(), 1.5, 0.5),
            "a point outside both is masked"
        );
    }

    /// A degenerate subpath encloses nothing and must not swallow the picture.
    #[test]
    fn a_subpath_with_too_few_points_is_ignored() {
        let degenerate = vec![crate::ir::Subpath::closed_outline(vec![
            (0.0, 0.0),
            (1.0, 1.0),
        ])];
        assert!(!point_is_inside_even_odd(&degenerate, 0.5, 0.5));
    }

    /// A vertex exactly on the ray must not be counted twice, or a whole row
    /// of the mask inverts.
    #[test]
    fn a_vertex_on_the_ray_is_counted_once() {
        let diamond = vec![crate::ir::Subpath::closed_outline(vec![
            (0.5, 0.0),
            (1.0, 0.5),
            (0.5, 1.0),
            (0.0, 0.5),
        ])];
        assert!(point_is_inside_even_odd(&diamond, 0.5, 0.5));
        assert!(!point_is_inside_even_odd(&diamond, 0.05, 0.05));
        // The ray at y = 0.5 passes exactly through the left and right vertices.
        assert!(!point_is_inside_even_odd(&diamond, 1.2, 0.5));
    }

    /// PowerPoint crops the source with `a:srcRect` first and clips the
    /// resulting frame with the custom geometry second. The renderer applies
    /// the crop later by narrowing the viewBox, so the clip has to be written
    /// into the region that narrowing will keep — expressed against the full
    /// box it lands on the pre-crop artwork instead (issue #1018).
    #[test]
    fn a_src_rect_crop_moves_the_svg_clip_into_the_kept_region() {
        let triangle = vec![crate::ir::Subpath::closed_outline(vec![
            (0.0, 0.0),
            (1.0, 0.0),
            (0.0, 1.0),
        ])];
        let crop = ImageCrop {
            left: 0.25,
            top: 0.0,
            right: 0.25,
            bottom: 0.0,
        };
        let out =
            clip_svg_to_path(SVG.as_bytes(), &triangle, Some(crop)).expect("the SVG is clipped");
        let text = String::from_utf8(out).expect("still an SVG");
        // The kept region is x 25..75 of the 100 x 50 viewBox.
        assert!(
            text.contains("M 25 0 L 75 0 L 25 50 Z"),
            "the clip spans the cropped frame, not the full box: {text}"
        );
    }

    /// A crop the renderer will refuse to apply (nothing kept) must leave the
    /// clip on the full box, or clip and crop disagree about the frame again.
    #[test]
    fn a_degenerate_crop_leaves_the_svg_clip_on_the_full_box() {
        let triangle = vec![crate::ir::Subpath::closed_outline(vec![
            (0.0, 0.0),
            (1.0, 0.0),
            (0.0, 1.0),
        ])];
        let crop = ImageCrop {
            left: 0.7,
            top: 0.0,
            right: 0.7,
            bottom: 0.0,
        };
        let out =
            clip_svg_to_path(SVG.as_bytes(), &triangle, Some(crop)).expect("the SVG is clipped");
        let text = String::from_utf8(out).expect("still an SVG");
        assert!(
            text.contains("M 0 0 L 100 0 L 0 50 Z"),
            "an unapplied crop keeps the clip on the full box: {text}"
        );
    }

    /// The raster mask has the same two coordinate systems: the alpha is
    /// zeroed on the full bitmap while the crop happens later in the
    /// renderer, so the path has to be evaluated in cropped-frame
    /// coordinates (issue #1018).
    #[test]
    fn a_src_rect_crop_moves_the_raster_mask_too() {
        let opaque = image::RgbaImage::from_pixel(8, 8, image::Rgba([255, 0, 0, 255]));
        let mut png = std::io::Cursor::new(Vec::new());
        image::DynamicImage::ImageRgba8(opaque)
            .write_to(&mut png, image::ImageFormat::Png)
            .expect("encode the fixture");
        // Keep the right half of the source; the path keeps the frame's
        // left half — pixels 4..6 of the original bitmap.
        let crop = ImageCrop {
            left: 0.5,
            top: 0.0,
            right: 0.0,
            bottom: 0.0,
        };
        let half = vec![crate::ir::Subpath::closed_outline(vec![
            (0.0, 0.0),
            (0.5, 0.0),
            (0.5, 1.0),
            (0.0, 1.0),
        ])];
        let (data, _format) =
            apply_path_mask(&png.into_inner(), &half, Some(crop)).expect("the raster is masked");
        let masked = image::load_from_memory(&data).expect("decode").into_rgba8();
        assert_eq!(
            masked.get_pixel(4, 4)[3],
            255,
            "the frame's left half survives"
        );
        assert_eq!(
            masked.get_pixel(7, 4)[3],
            0,
            "the frame's right half is masked away"
        );
    }

    /// End to end through the parser: the probe deck's slide 2 combines
    /// `srcRect l="20000" t="10000"` with a pentagon custGeom on an SVG
    /// picture (the issue #1018 shape). The clip must land in the kept
    /// region of the 800x600 viewBox — x 160..800, y 60..600 — not the
    /// full box.
    #[test]
    fn a_cropped_svg_pictures_clip_lands_in_the_kept_region() {
        let data = include_bytes!("../../../../tests/fixtures/pptx/svg-srcrect-clip-probe.pptx");
        let (doc, _warnings) = crate::parser::pptx::PptxParser
            .parse(data, &ConvertOptions::default())
            .expect("the probe deck parses");
        let Page::Fixed(page) = &doc.pages[1] else {
            panic!("slide 2 is a fixed page");
        };
        let svg = page
            .elements
            .iter()
            .find_map(|element| match &element.kind {
                FixedElementKind::Image(image) if image.format == ImageFormat::Svg => {
                    Some(std::str::from_utf8(&image.data).expect("SVG stays UTF-8"))
                }
                _ => None,
            })
            .expect("slide 2 carries the clipped SVG");
        // Pentagon vertices mapped into the kept region: apex (480, 60) and
        // the mid-left anchor (160, 330). The full-box mapping would put
        // them at (400, 0) and (0, 300).
        assert!(
            svg.contains("M 480 60") && svg.contains("L 160 330"),
            "the clip is written in cropped-frame coordinates: {}",
            &svg[..svg.len().min(600)]
        );
    }
}
