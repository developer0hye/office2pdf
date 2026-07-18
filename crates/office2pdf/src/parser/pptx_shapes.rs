use super::*;

/// Map OOXML preset dash values to `BorderLineStyle`.
pub(super) fn pptx_dash_to_border_style(val: &str) -> BorderLineStyle {
    match val {
        "dash" | "lgDash" | "sysDash" => BorderLineStyle::Dashed,
        "dot" | "sysDot" | "lgDashDot" => BorderLineStyle::Dotted,
        "dashDot" | "sysDashDot" => BorderLineStyle::DashDot,
        "lgDashDotDot" | "sysDashDotDot" => BorderLineStyle::DashDotDot,
        "solid" => BorderLineStyle::Solid,
        _ => BorderLineStyle::Solid,
    }
}

/// Group shape coordinate transform.
///
/// Maps child coordinates from the group's internal coordinate space
/// to the parent (slide or outer group) coordinate space.
#[derive(Debug, Default)]
struct GroupTransform {
    /// Group position on parent, in EMU.
    off_x: i64,
    off_y: i64,
    /// Group extent (size) on parent, in EMU.
    ext_cx: i64,
    ext_cy: i64,
    /// Child coordinate space origin, in EMU.
    ch_off_x: i64,
    ch_off_y: i64,
    /// Child coordinate space extent, in EMU.
    ch_ext_cx: i64,
    ch_ext_cy: i64,
    /// Group rotation in degrees (clockwise), from the group xfrm `rot`.
    rot_deg: f64,
}

impl GroupTransform {
    /// Apply the transform to a `FixedElement` whose coordinates are already in points.
    fn apply(&self, elem: &mut FixedElement) {
        let scale_x = if self.ch_ext_cx != 0 {
            self.ext_cx as f64 / self.ch_ext_cx as f64
        } else {
            1.0
        };
        let scale_y = if self.ch_ext_cy != 0 {
            self.ext_cy as f64 / self.ch_ext_cy as f64
        } else {
            1.0
        };

        let off_x_pt = emu_to_pt(self.off_x);
        let off_y_pt = emu_to_pt(self.off_y);
        let ch_off_x_pt = emu_to_pt(self.ch_off_x);
        let ch_off_y_pt = emu_to_pt(self.ch_off_y);

        elem.x = off_x_pt + (elem.x - ch_off_x_pt) * scale_x;
        elem.y = off_y_pt + (elem.y - ch_off_y_pt) * scale_y;
        elem.width *= scale_x;
        elem.height *= scale_y;

        // Scale inner ImageData dimensions so the rendered image matches
        // the group-transformed size, not the raw child-space size.
        if let FixedElementKind::Image(ref mut img) = elem.kind {
            if let Some(ref mut w) = img.width {
                *w *= scale_x;
            }
            if let Some(ref mut h) = img.height {
                *h *= scale_y;
            }
            if let Some(ref mut stroke) = img.stroke {
                stroke.width *= (scale_x + scale_y) / 2.0;
            }
        }

        // Line and polyline geometry is baked in child-space points; scale
        // them with the group, or hairline axes collapse to sub-pixel stubs.
        if let FixedElementKind::Shape(ref mut shape) = elem.kind {
            match &mut shape.kind {
                ShapeKind::Line { x1, y1, x2, y2, .. } => {
                    *x1 *= scale_x;
                    *x2 *= scale_x;
                    *y1 *= scale_y;
                    *y2 *= scale_y;
                }
                ShapeKind::Polyline { points, .. } => {
                    for (x, y) in points.iter_mut() {
                        *x *= scale_x;
                        *y *= scale_y;
                    }
                }
                _ => {}
            }
        }

        // Compose the group's own rotation: orbit the child's center around
        // the group center and add the angle to the child's own rotation
        // (shape geometry only — images and text boxes carry no rotation).
        if self.rot_deg != 0.0 {
            let group_center_x = off_x_pt + emu_to_pt(self.ext_cx) / 2.0;
            let group_center_y = off_y_pt + emu_to_pt(self.ext_cy) / 2.0;
            let element_center_x = elem.x + elem.width / 2.0;
            let element_center_y = elem.y + elem.height / 2.0;
            let radians = self.rot_deg.to_radians();
            let (sin, cos) = radians.sin_cos();
            let dx = element_center_x - group_center_x;
            let dy = element_center_y - group_center_y;
            let rotated_x = group_center_x + dx * cos - dy * sin;
            let rotated_y = group_center_y + dx * sin + dy * cos;
            elem.x = rotated_x - elem.width / 2.0;
            elem.y = rotated_y - elem.height / 2.0;
            if let FixedElementKind::Shape(ref mut shape) = elem.kind {
                shape.rotation_deg = Some(shape.rotation_deg.unwrap_or(0.0) + self.rot_deg);
            }
        }
    }
}

/// Parse a `<p:grpSp>` group shape from the reader.
///
/// Called right after the `<p:grpSp>` start tag has been consumed.
/// Reads through the group's header sections (`nvGrpSpPr`, `grpSpPr`),
/// extracts the coordinate transform, then slices the original XML to
/// get the child shapes, and recursively parses them via `parse_slide_xml`.
#[allow(clippy::too_many_arguments)]
pub(super) fn parse_group_shape(
    reader: &mut Reader<&[u8]>,
    xml: &str,
    images: &SlideImageMap,
    theme: &ThemeData,
    color_map: &ColorMapData,
    warning_context: &str,
    inherited_text_body_defaults: &PptxTextBodyStyleDefaults,
    table_styles: &table_styles::TableStyleMap,
) -> Result<(Vec<FixedElement>, Vec<ConvertWarning>), ConvertError> {
    let mut transform = GroupTransform::default();
    let mut in_xfrm = false;
    let mut header_depth: usize = 0;
    let mut children_start = reader.buffer_position() as usize;

    // Phase 1: Read nvGrpSpPr and grpSpPr sections, extracting the transform.
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"nvGrpSpPr" if header_depth == 0 => header_depth = 1,
                b"grpSpPr" if header_depth == 0 => header_depth = 1,
                b"xfrm" if header_depth == 1 => {
                    in_xfrm = true;
                    if let Some(rot) = get_attr_i64(e, b"rot") {
                        transform.rot_deg = rot as f64 / 60_000.0;
                    }
                }
                _ if header_depth > 0 => header_depth += 1,
                _ => break,
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"grpSpPr" if header_depth == 0 => {
                    children_start = reader.buffer_position() as usize;
                    break;
                }
                b"off" if in_xfrm => {
                    transform.off_x = get_attr_i64(e, b"x").unwrap_or(0);
                    transform.off_y = get_attr_i64(e, b"y").unwrap_or(0);
                }
                b"ext" if in_xfrm => {
                    transform.ext_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                    transform.ext_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                }
                b"chOff" if in_xfrm => {
                    transform.ch_off_x = get_attr_i64(e, b"x").unwrap_or(0);
                    transform.ch_off_y = get_attr_i64(e, b"y").unwrap_or(0);
                }
                b"chExt" if in_xfrm => {
                    transform.ch_ext_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                    transform.ch_ext_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"xfrm" if in_xfrm => in_xfrm = false,
                b"grpSpPr" if header_depth == 1 => {
                    children_start = reader.buffer_position() as usize;
                    break;
                }
                b"nvGrpSpPr" if header_depth == 1 => header_depth = 0,
                _ if header_depth > 1 => header_depth -= 1,
                b"grpSp" => return Ok((Vec::new(), Vec::new())),
                _ => {}
            },
            Ok(Event::Eof) => return Ok((Vec::new(), Vec::new())),
            Err(error) => {
                return Err(crate::parser::parse_err(format!(
                    "XML error in group shape: {error}"
                )));
            }
            _ => {}
        }
    }

    // Phase 2: Skip to </p:grpSp>, recording where the children end.
    let mut group_depth: usize = 1;
    loop {
        let position = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpSp" => {
                group_depth += 1;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"grpSp" => {
                group_depth -= 1;
                if group_depth == 0 {
                    let children_xml = &xml[children_start..position];
                    if children_xml.trim().is_empty() {
                        return Ok((Vec::new(), Vec::new()));
                    }

                    let wrapped = format!(
                        r#"<r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{children_xml}</r>"#
                    );

                    let (mut child_elements, warnings) = parse_slide_xml(
                        &wrapped,
                        images,
                        theme,
                        color_map,
                        warning_context,
                        inherited_text_body_defaults,
                        table_styles,
                        None,
                    )?;
                    for element in &mut child_elements {
                        transform.apply(element);
                    }
                    return Ok((child_elements, warnings));
                }
            }
            Ok(Event::Eof) => return Ok((Vec::new(), Vec::new())),
            Err(error) => {
                return Err(crate::parser::parse_err(format!(
                    "XML error in group shape: {error}"
                )));
            }
            _ => {}
        }
    }
}

fn parse_crop_fraction(e: &quick_xml::events::BytesStart, key: &[u8]) -> f64 {
    get_attr_i64(e, key)
        .map(|value| (value as f64 / 100_000.0).clamp(0.0, 1.0))
        .unwrap_or(0.0)
}

pub(super) fn parse_src_rect(e: &quick_xml::events::BytesStart) -> Option<ImageCrop> {
    let crop = ImageCrop {
        left: parse_crop_fraction(e, b"l"),
        top: parse_crop_fraction(e, b"t"),
        right: parse_crop_fraction(e, b"r"),
        bottom: parse_crop_fraction(e, b"b"),
    };
    (!crop.is_empty()).then_some(crop)
}

/// Map a PPTX preset geometry name to an IR ShapeKind.
///
/// `flip_h`/`flip_v` from `<a:xfrm>` reverse the line endpoint direction,
/// which matters for connectors drawn right-to-left or bottom-to-top.
#[allow(clippy::too_many_arguments)]
pub(super) fn prst_to_shape_kind(
    prst: &str,
    width: f64,
    height: f64,
    flip_h: bool,
    flip_v: bool,
    head_end: ArrowHead,
    tail_end: ArrowHead,
    adj_values: &[f64],
) -> ShapeKind {
    match prst {
        "ellipse" => ShapeKind::Ellipse,
        "line" | "straightConnector1" => {
            let (x1, y1, x2, y2) = line_endpoints(width, height, flip_h, flip_v);
            ShapeKind::Line {
                x1,
                y1,
                x2,
                y2,
                head_end,
                tail_end,
            }
        }
        // Bent connectors: L-shaped or Z-shaped paths
        "bentConnector2" => {
            let points: Vec<(f64, f64)> = bent_connector2_points(width, height, flip_h, flip_v);
            ShapeKind::Polyline {
                points,
                head_end,
                tail_end,
            }
        }
        "bentConnector3" => {
            let adj: f64 = adj_values.first().copied().unwrap_or(50_000.0) / 100_000.0;
            let points: Vec<(f64, f64)> =
                bent_connector3_points(width, height, flip_h, flip_v, adj);
            ShapeKind::Polyline {
                points,
                head_end,
                tail_end,
            }
        }
        "bentConnector4" | "bentConnector5" => {
            let adj1: f64 = adj_values.first().copied().unwrap_or(50_000.0) / 100_000.0;
            let adj2: f64 = adj_values.get(1).copied().unwrap_or(50_000.0) / 100_000.0;
            let points: Vec<(f64, f64)> =
                bent_connector4_points(width, height, flip_h, flip_v, adj1, adj2);
            ShapeKind::Polyline {
                points,
                head_end,
                tail_end,
            }
        }
        // Curved connectors: approximated as bent for now
        "curvedConnector2" | "curvedConnector3" | "curvedConnector4" | "curvedConnector5" => {
            let (x1, y1, x2, y2) = line_endpoints(width, height, flip_h, flip_v);
            ShapeKind::Line {
                x1,
                y1,
                x2,
                y2,
                head_end,
                tail_end,
            }
        }
        "roundRect" => ShapeKind::RoundedRectangle {
            radius_fraction: 0.1,
        },
        // homePlate: pentagon arrow tab (rect with pointed right edge)
        "homePlate" => {
            let adj: f64 = adj_values.first().copied().unwrap_or(50_000.0);
            let ss: f64 = width.min(height);
            let dx: f64 = (adj / 100_000.0 * ss).min(width);
            let notch_x: f64 = (width - dx) / width;
            ShapeKind::Polygon {
                vertices: vec![
                    (0.0, 0.0),
                    (notch_x, 0.0),
                    (1.0, 0.5),
                    (notch_x, 1.0),
                    (0.0, 1.0),
                ],
            }
        }
        "triangle" => ShapeKind::Polygon {
            vertices: vec![(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)],
        },
        "rtTriangle" => ShapeKind::Polygon {
            vertices: vec![(0.0, 0.0), (1.0, 1.0), (0.0, 1.0)],
        },
        "diamond" => ShapeKind::Polygon {
            vertices: vec![(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)],
        },
        "pentagon" => ShapeKind::Polygon {
            vertices: regular_polygon_vertices(5),
        },
        "hexagon" => ShapeKind::Polygon {
            vertices: regular_polygon_vertices(6),
        },
        "octagon" => ShapeKind::Polygon {
            vertices: regular_polygon_vertices(8),
        },
        "rightArrow" | "arrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Right),
        },
        "leftArrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Left),
        },
        "upArrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Up),
        },
        "downArrow" => ShapeKind::Polygon {
            vertices: arrow_vertices(ArrowDir::Down),
        },
        "star4" => ShapeKind::Polygon {
            vertices: star_vertices(4),
        },
        "star5" => ShapeKind::Polygon {
            vertices: star_vertices(5),
        },
        "star6" => ShapeKind::Polygon {
            vertices: star_vertices(6),
        },
        _ => ShapeKind::Rectangle,
    }
}

enum ArrowDir {
    Right,
    Left,
    Up,
    Down,
}

/// Generate vertices for a regular polygon inscribed in the unit square (0–1).
fn regular_polygon_vertices(n: usize) -> Vec<(f64, f64)> {
    let mut vertices = Vec::with_capacity(n);
    for i in 0..n {
        let angle = -std::f64::consts::FRAC_PI_2 + 2.0 * std::f64::consts::PI * i as f64 / n as f64;
        let x = 0.5 + 0.5 * angle.cos();
        let y = 0.5 + 0.5 * angle.sin();
        vertices.push((x, y));
    }
    // Preset geometries fill the whole shape box; the inscribed-circle
    // vertices leave slack on the flat sides (a pentagon only reaches ~90%
    // height), printing the shape smaller than PowerPoint (issue #319).
    let min_x = vertices.iter().map(|v| v.0).fold(f64::MAX, f64::min);
    let max_x = vertices.iter().map(|v| v.0).fold(f64::MIN, f64::max);
    let min_y = vertices.iter().map(|v| v.1).fold(f64::MAX, f64::min);
    let max_y = vertices.iter().map(|v| v.1).fold(f64::MIN, f64::max);
    let span_x = (max_x - min_x).max(f64::EPSILON);
    let span_y = (max_y - min_y).max(f64::EPSILON);
    vertices
        .into_iter()
        .map(|(x, y)| ((x - min_x) / span_x, (y - min_y) / span_y))
        .collect()
}

/// Generate arrow polygon vertices (7-point arrow) in normalized coordinates.
fn arrow_vertices(dir: ArrowDir) -> Vec<(f64, f64)> {
    let right: Vec<(f64, f64)> = vec![
        (0.0, 0.25),
        (0.6, 0.25),
        (0.6, 0.0),
        (1.0, 0.5),
        (0.6, 1.0),
        (0.6, 0.75),
        (0.0, 0.75),
    ];
    match dir {
        ArrowDir::Right => right,
        ArrowDir::Left => right.into_iter().map(|(x, y)| (1.0 - x, y)).collect(),
        ArrowDir::Up => right.into_iter().map(|(x, y)| (y, 1.0 - x)).collect(),
        ArrowDir::Down => right.into_iter().map(|(x, y)| (1.0 - y, x)).collect(),
    }
}

/// Generate star polygon vertices with `n` points inscribed in the unit square.
fn star_vertices(n: usize) -> Vec<(f64, f64)> {
    let mut vertices = Vec::with_capacity(n * 2);
    let inner_radius = 0.4;
    for i in 0..(n * 2) {
        let angle = -std::f64::consts::FRAC_PI_2 + std::f64::consts::PI * i as f64 / n as f64;
        let radius = if i % 2 == 0 { 0.5 } else { 0.5 * inner_radius };
        let x = 0.5 + radius * angle.cos();
        let y = 0.5 + radius * angle.sin();
        vertices.push((x, y));
    }
    vertices
}

// ── Connector geometry helpers ──────────────────────────────────────

/// Compute line start/end points within the bounding box, accounting for flips.
///
/// Without flip: (0,0) → (w,h).  With flipH: (w,0) → (0,h).
/// With flipV: (0,h) → (w,0).  Both: (w,h) → (0,0).
fn line_endpoints(width: f64, height: f64, flip_h: bool, flip_v: bool) -> (f64, f64, f64, f64) {
    let (x1, x2): (f64, f64) = if flip_h { (width, 0.0) } else { (0.0, width) };
    let (y1, y2): (f64, f64) = if flip_v { (height, 0.0) } else { (0.0, height) };
    (x1, y1, x2, y2)
}

/// bentConnector2: simple L-shape (one bend).
///
/// Without flip: right then down → (0,0) → (w,0) → (w,h).
fn bent_connector2_points(width: f64, height: f64, flip_h: bool, flip_v: bool) -> Vec<(f64, f64)> {
    let (x1, y1, x2, y2) = line_endpoints(width, height, flip_h, flip_v);
    vec![(x1, y1), (x2, y1), (x2, y2)]
}

/// bentConnector3: Z-shape with one adjustable midpoint.
///
/// `adj` is the fraction (0.0–1.0) along the primary axis where the bend occurs.
/// Without flip: right to adj%, then vertical, then right to end.
fn bent_connector3_points(
    width: f64,
    height: f64,
    flip_h: bool,
    flip_v: bool,
    adj: f64,
) -> Vec<(f64, f64)> {
    let (x1, y1, x2, y2) = line_endpoints(width, height, flip_h, flip_v);
    let mid_x: f64 = x1 + (x2 - x1) * adj;
    vec![(x1, y1), (mid_x, y1), (mid_x, y2), (x2, y2)]
}

/// bentConnector4: S-shape with two adjustable midpoints.
fn bent_connector4_points(
    width: f64,
    height: f64,
    flip_h: bool,
    flip_v: bool,
    adj1: f64,
    adj2: f64,
) -> Vec<(f64, f64)> {
    let (x1, y1, x2, y2) = line_endpoints(width, height, flip_h, flip_v);
    let mid_x: f64 = x1 + (x2 - x1) * adj1;
    let mid_y: f64 = y1 + (y2 - y1) * adj2;
    vec![(x1, y1), (mid_x, y1), (mid_x, mid_y), (x2, mid_y), (x2, y2)]
}

/// Parse OOXML arrowhead type attribute to IR ArrowHead.
pub(super) fn parse_arrow_head(type_val: Option<&str>) -> ArrowHead {
    match type_val {
        Some("triangle" | "stealth" | "arrow" | "diamond" | "oval") => ArrowHead::Triangle,
        _ => ArrowHead::None,
    }
}
