use std::fmt::Write;

use super::shadow_outline::{CornerReach, OffsetCorner, arc_beziers, offset_ring};
use super::*;
use crate::ir::{Subpath, TopBevel};

pub(super) fn generate_shape(
    out: &mut String,
    shape: &Shape,
    width: f64,
    height: f64,
    ctx: &mut GenCtx,
) {
    // Render shadow as offset duplicate before main shape
    if let Some(shadow) = &shape.shadow {
        write_shadow_shape(out, shape, width, height, shadow, ctx);
    }

    let use_typst_rotation = shape.rotation_deg.is_some()
        && !matches!(
            shape.kind,
            ShapeKind::Line { .. } | ShapeKind::Polyline { .. }
        );
    if let Some(deg) = shape.rotation_deg.filter(|_| use_typst_rotation) {
        let (dx, dy): (f64, f64) = centre_pivot_shift(width, height, deg, false, false);
        let _ = write!(
            out,
            "#move(dx: {}pt, dy: {}pt)[#rotate({}deg, origin: top + left, reflow: false)[",
            format_f64(dx),
            format_f64(dy),
            format_f64(deg)
        );
    }

    match &shape.kind {
        ShapeKind::Rectangle => {
            out.push_str("#rect(");
            write_shape_params(out, shape, width, height);
            out.push_str(")\n");
        }
        ShapeKind::Ellipse => {
            out.push_str("#ellipse(");
            write_shape_params(out, shape, width, height);
            out.push_str(")\n");
        }
        ShapeKind::Line {
            x1,
            y1,
            x2,
            y2,
            head_end,
            tail_end,
        } => {
            let ((start_x, start_y), (end_x, end_y)) =
                rotated_line_points(*x1, *y1, *x2, *y2, width, height, shape.rotation_deg);
            let has_arrowheads: bool = *tail_end != ArrowHead::None || *head_end != ArrowHead::None;
            // When arrowheads follow the line, wrap everything in #place()
            // so that Typst overlays them at the same origin instead of
            // stacking sequentially.
            if has_arrowheads {
                out.push_str("#place(top + left)[");
            }
            out.push_str("#line(");
            let _ = write!(
                out,
                "start: ({}pt, {}pt), end: ({}pt, {}pt)",
                format_f64(start_x),
                format_f64(start_y),
                format_f64(end_x),
                format_f64(end_y),
            );
            write_shape_stroke(out, &shape.stroke);
            out.push_str(")\n");
            if has_arrowheads {
                out.push_str("]\n");
            }
            if *tail_end != ArrowHead::None {
                write_arrowhead_at(out, &shape.stroke, (start_x, start_y), (end_x, end_y));
            }
            if *head_end != ArrowHead::None {
                write_arrowhead_at(out, &shape.stroke, (end_x, end_y), (start_x, start_y));
            }
        }
        ShapeKind::Polyline {
            points,
            head_end,
            tail_end,
        } => {
            let rotated_points: Vec<(f64, f64)> =
                rotate_points(points, width, height, shape.rotation_deg);
            write_polyline(out, &shape.stroke, &rotated_points);
            if rotated_points.len() >= 2 {
                if *tail_end != ArrowHead::None {
                    let last = rotated_points[rotated_points.len() - 1];
                    let second_last = rotated_points[rotated_points.len() - 2];
                    write_arrowhead_at(out, &shape.stroke, second_last, last);
                }
                if *head_end != ArrowHead::None {
                    let first = rotated_points[0];
                    let second = rotated_points[1];
                    write_arrowhead_at(out, &shape.stroke, second, first);
                }
            }
        }
        ShapeKind::RoundedRectangle { radius_fraction } => {
            let radius = radius_fraction * width.min(height);
            out.push_str("#rect(");
            write_shape_params(out, shape, width, height);
            let _ = write!(out, ", radius: {}pt", format_f64(radius));
            out.push_str(")\n");
        }
        ShapeKind::Polygon { vertices } => {
            write_polygon(out, shape, width, height, vertices);
        }
        ShapeKind::Path { subpaths } => {
            write_subpath_curve(out, shape, width, height, subpaths);
        }
    }

    if matches!(shape.kind, ShapeKind::Rectangle)
        && let Some(bevel) = &shape.top_bevel
    {
        write_rectangle_top_bevel(out, shape, width, height, bevel);
    }

    if use_typst_rotation {
        out.push_str("]]\n");
    }
}

/// Approximate an orthographic circular top bevel with four inset faces.
///
/// Typst has no 3-D shape primitive, so the bevel is painted over the normal
/// face. A subtle cool front-light pass matches PowerPoint's three-point rig;
/// the top face carries the circular profile's narrow specular band, while
/// the remaining faces darken toward their outer edges. Each overlay is
/// absolutely placed so it cannot change the surrounding fixed layout
/// (issue #1298).
fn write_rectangle_top_bevel(
    out: &mut String,
    shape: &Shape,
    width: f64,
    height: f64,
    bevel: &TopBevel,
) {
    if shape.fill.is_none() && shape.gradient_fill.is_none() && shape.pattern_fill.is_none() {
        return;
    }

    let inset = bevel.width.min(width / 2.0).min(height / 2.0).max(0.0);
    if inset <= 0.0 {
        return;
    }

    let strength = (bevel.height / bevel.width.max(f64::EPSILON) / 0.4).clamp(0.25, 2.0);
    let shape_opacity = shape.opacity.unwrap_or(1.0).clamp(0.0, 1.0);
    let alpha =
        |base: f64| -> u8 { (base * strength * shape_opacity).clamp(0.0, 255.0).round() as u8 };
    let rotation_bias = (bevel.light_rig_rotation_deg.to_radians().sin() * 16.0).clamp(-16.0, 16.0);

    // Three-point lighting lifts the front face slightly toward the cool
    // highlight. Keeping this translucent preserves solid, gradient, and
    // pattern fills underneath.
    let _ = writeln!(
        out,
        "#place(top + left)[#rect(width: {}pt, height: {}pt, fill: gradient.linear((rgb(90, 165, 255, {}), 0%), (rgb(90, 165, 255, {}), 100%), angle: 270deg, space: rgb), stroke: none)]",
        format_f64(width),
        format_f64(height),
        alpha(40.0),
        alpha(64.0),
    );

    let right = width - inset;
    let bottom = height - inset;
    let top_fill = format!(
        "gradient.linear((rgb(90, 165, 255, 0), 0%), (rgb(150, 225, 255, {}), 25%), (rgb(90, 165, 255, 0), 100%), angle: 90deg, space: rgb)",
        alpha(255.0),
    );
    let left_fill = format!(
        "gradient.linear((rgb(0, 0, 0, {}), 0%), (rgb(0, 0, 0, {}), 20%), (rgb(0, 0, 0, {}), 50%), (rgb(0, 0, 0, 0), 100%), space: rgb)",
        alpha(80.0 - rotation_bias),
        alpha(38.0 - rotation_bias / 2.0),
        alpha(13.0),
    );

    write_bevel_face(
        out,
        &[(0.0, 0.0), (inset, inset), (right, inset), (width, 0.0)],
        &top_fill,
    );
    write_bevel_face(
        out,
        &[(0.0, 0.0), (inset, inset), (inset, bottom), (0.0, height)],
        &left_fill,
    );

    // Typst clamps a gradient on the oppositely wound right and bottom
    // trapezoids to their dark stop. Thin, non-overlapping slices preserve
    // the circular falloff without introducing layout-visible boxes.
    const SLICES: usize = 10;
    for index in 0..SLICES {
        let start = index as f64 / SLICES as f64;
        let end = (index + 1) as f64 / SLICES as f64;
        let midpoint = (start + end) / 2.0;

        let x0 = right + start * inset;
        let x1 = right + end * inset;
        let top0 = inset * (1.0 - start);
        let top1 = inset * (1.0 - end);
        let bottom0 = height - top0;
        let bottom1 = height - top1;
        let side_alpha = alpha((80.0 + rotation_bias) * midpoint.powf(2.0));
        write_bevel_face(
            out,
            &[(x0, top0), (x1, top1), (x1, bottom1), (x0, bottom0)],
            &format!("rgb(0, 0, 0, {side_alpha})"),
        );

        let y0 = bottom + start * inset;
        let y1 = bottom + end * inset;
        let left0 = inset * (1.0 - start);
        let left1 = inset * (1.0 - end);
        let bottom_color = if midpoint < 0.8 {
            format!("rgb(0, 0, 0, {})", alpha(28.0 * midpoint))
        } else {
            let contact = (midpoint - 0.8) / 0.2;
            format!("rgb(30, 15, 0, {})", alpha(20.0 + 60.0 * contact),)
        };
        write_bevel_face(
            out,
            &[
                (left0, y0),
                (width - left0, y0),
                (width - left1, y1),
                (left1, y1),
            ],
            &bottom_color,
        );
    }
}

fn write_bevel_face(out: &mut String, vertices: &[(f64, f64)], fill: &str) {
    out.push_str("#place(top + left)[#polygon(");
    for (index, (x, y)) in vertices.iter().enumerate() {
        if index > 0 {
            out.push_str(", ");
        }
        let _ = write!(out, "({}pt, {}pt)", format_f64(*x), format_f64(*y));
    }
    let _ = writeln!(out, ", fill: {fill}, stroke: none)]");
}

fn rotated_line_points(
    x1: f64,
    y1: f64,
    x2: f64,
    y2: f64,
    width: f64,
    height: f64,
    rotation_deg: Option<f64>,
) -> ((f64, f64), (f64, f64)) {
    (
        rotate_point((x1, y1), width, height, rotation_deg),
        rotate_point((x2, y2), width, height, rotation_deg),
    )
}

fn rotate_points(
    points: &[(f64, f64)],
    width: f64,
    height: f64,
    rotation_deg: Option<f64>,
) -> Vec<(f64, f64)> {
    points
        .iter()
        .copied()
        .map(|point| rotate_point(point, width, height, rotation_deg))
        .collect()
}

fn rotate_point(
    point: (f64, f64),
    width: f64,
    height: f64,
    rotation_deg: Option<f64>,
) -> (f64, f64) {
    let Some(rotation_deg) = rotation_deg else {
        return point;
    };

    if rotation_deg.abs() < 0.001 {
        return point;
    }

    let angle_rad = rotation_deg.to_radians();
    let cos_theta = angle_rad.cos();
    let sin_theta = angle_rad.sin();
    let center_x = width / 2.0;
    let center_y = height / 2.0;
    let delta_x = point.0 - center_x;
    let delta_y = point.1 - center_y;

    (
        center_x + delta_x * cos_theta - delta_y * sin_theta,
        center_y + delta_x * sin_theta + delta_y * cos_theta,
    )
}

/// The blur's standard deviation as a fraction of `blurRad`: PowerPoint
/// treats the declared radius as the blur's full 3-sigma extent. A
/// one-factor probe over native exports at blur 1/3.15/6.3/12.6/18.9pt
/// fit the flattened shadow bitmap's ramp — a Gaussian CDF centred on
/// the shadow silhouette — at sigma/blurRad 0.331–0.345 on every edge
/// (issue #784, tightening #390's coarser 0.23–0.35 reading whose 0.3
/// midpoint cut the ramp's reach and density about 10% short).
const SHADOW_BLUR_SIGMA_PER_RADIUS: f64 = 1.0 / 3.0;

/// How far the filtered SVG follows the Gaussian tail, in sigma units.
///
/// At 2.6 sigma the one-sided tail is below 0.5%, so clipping the generated
/// asset there preserves the visible reach PowerPoint exports without giving
/// each shadow an unbounded bitmap (issues #662, #1309).
pub(super) const SHADOW_BLUR_EXTENT_SIGMA: f64 = 2.6;

/// The blur's standard deviation in points, zero for a crisp shadow.
pub(super) fn shadow_blur_sigma(shadow: &Shadow) -> f64 {
    (SHADOW_BLUR_SIGMA_PER_RADIUS * shadow.blur_radius).max(0.0)
}

pub(super) fn shadow_alpha(shadow: &Shadow) -> u8 {
    (shadow.opacity.clamp(0.0, 1.0) * 255.0).round() as u8
}

/// How far an outline pushes the shadow's silhouette past the fill path.
///
/// PowerPoint casts an `a:outerShdw` from the *stroked* shape — the fill path
/// grown by half the line width, since a stroke straddles its path — and only
/// then offsets it by `dist`. Measured against a native macOS PowerPoint
/// export of `customGeo.pptx` page 46: the title banner's fill path is
/// x [60, 672], y [36, 113.75] under a 3pt outline, and the export's
/// flattened shadow bitmap puts its half-alpha silhouette at x [58.44,
/// 673.56], y [36.00, 116.88] — fill outset by 1.5pt, then `dist` 1.57pt
/// down, agreeing on every edge to within 0.09pt (issue #1057).
///
/// Casting from the fill path alone left the whole ramp a half-width short on
/// every side, which a light outline then covered at its densest — the same
/// defect that read as the banner's outer white edge going missing (#1044).
pub(super) fn shadow_outline_outset(stroke: &Option<BorderSide>) -> f64 {
    stroke
        .as_ref()
        .filter(|stroke| stroke.style != BorderLineStyle::None)
        .map_or(0.0, |stroke| (stroke.width / 2.0).max(0.0))
}

/// The corner radius a square-cornered fill path's shadow silhouette carries.
///
/// The silhouette is cast from the *stroked* shape (#1057), so it turns each
/// corner the way the outline's `a:ln` join does. DrawingML's default is
/// `a:round` (#1090), whose arc has the stroke's half-width as its radius and
/// is centred on the fill path's corner — exactly the arc a `#rect` grown by
/// that half-width and given that radius draws. `a:miter` instead runs the
/// stroke out to a point that already lies on the grown box's own corner, so
/// it contributes no arc.
///
/// `a:bevel` chamfers the corner. The chamfer's two ends are the round join's
/// two tangent points, so that arc is the closest circle `#rect`'s single
/// `radius` can spell: it sits outside the true chamfer by at most
/// `(2 - sqrt(2)) / 4` of the line width, against the `sqrt(2) / 4` a square
/// corner overshoots it by.
///
/// Distinct from [`shadow_outline_outset`], which is how far the outline
/// pushes every edge out: a mitred 3pt outline still grows the silhouette by
/// 1.5pt a side while leaving its corners square.
pub(super) fn shadow_silhouette_corner_radius(stroke: &Option<BorderSide>) -> f64 {
    stroke
        .as_ref()
        .filter(|stroke| stroke.style != BorderLineStyle::None)
        .filter(|stroke| stroke.join != LineJoin::Miter)
        .map_or(0.0, |stroke| (stroke.width / 2.0).max(0.0))
}

/// A corner arc can never bite deeper than half the box it turns; Typst would
/// clamp it anyway.
pub(super) fn clamp_ring_corner_radius(radius: f64, width: f64, height: f64) -> f64 {
    radius.clamp(0.0, 0.5 * width.min(height))
}

/// Render a shadow before the shape it sits under.
fn write_shadow_shape(
    out: &mut String,
    shape: &Shape,
    width: f64,
    height: f64,
    shadow: &Shadow,
    ctx: &mut GenCtx,
) {
    if matches!(
        shape.kind,
        ShapeKind::Line { .. } | ShapeKind::Polyline { .. }
    ) {
        // Lines don't have meaningful shadows; skip
        return;
    }
    let dir_rad = shadow.direction.to_radians();
    let dx = shadow.distance * dir_rad.cos();
    let dy = shadow.distance * dir_rad.sin();
    let outline_outset: f64 = shadow_outline_outset(&shape.stroke);
    if shadow.blur_radius > 0.0 {
        write_blurred_shadow_asset(
            out,
            ctx,
            &shape.kind,
            (width, height),
            &shape.stroke,
            shadow,
            (dx, dy),
        );
        return;
    }
    let silhouette_radius: f64 = shadow_silhouette_corner_radius(&shape.stroke);
    // A polygon and a custom geometry both cast the shadow of their own
    // outline, so both take the same silhouette: closed rings in the shape's
    // frame, each offset rather than scaled (issue #1206).
    let outline_rings: Option<Vec<OutlineRing>> = shadow_outline_rings(&shape.kind, width, height);

    let alpha: u8 = shadow_alpha(shadow);
    {
        let expansion = outline_outset;
        if let Some(rings) = &outline_rings {
            if !rings.is_empty() {
                write_offset_ring_layer(
                    out,
                    rings,
                    RingGeometry {
                        expansion,
                        silhouette_radius,
                    },
                    shadow,
                    alpha,
                    (dx, dy),
                );
            }
            if let ShapeKind::Path { subpaths } = &shape.kind {
                write_crisp_open_path_shadow(
                    out,
                    subpaths,
                    width,
                    height,
                    &shape.stroke,
                    shadow,
                    (dx, dy),
                );
            }
            return;
        }
        let layer_width = (width + 2.0 * expansion).max(0.0);
        let layer_height = (height + 2.0 * expansion).max(0.0);
        let _ = write!(
            out,
            "#place(top + left, dx: {}pt, dy: {}pt)[",
            format_f64(dx - expansion),
            format_f64(dy - expansion),
        );
        match &shape.kind {
            ShapeKind::RoundedRectangle { radius_fraction } => {
                // The outline's join cannot square a corner the fill path
                // already curves, so the silhouette's arc is the shape's own
                // plus the half-width the stroke adds (#1057).
                let silhouette_arc: f64 = radius_fraction * width.min(height) + outline_outset;
                let radius =
                    clamp_ring_corner_radius(silhouette_arc.max(0.0), layer_width, layer_height);
                let _ = write!(
                    out,
                    "#rect(width: {}pt, height: {}pt, radius: {}pt, fill: rgb({}, {}, {}, {}))",
                    format_f64(layer_width),
                    format_f64(layer_height),
                    format_f64(radius),
                    shadow.color.r,
                    shadow.color.g,
                    shadow.color.b,
                    alpha,
                );
            }
            ShapeKind::Rectangle => {
                let radius: f64 =
                    clamp_ring_corner_radius(silhouette_radius.max(0.0), layer_width, layer_height);
                let _ = write!(
                    out,
                    "#rect(width: {}pt, height: {}pt, radius: {}pt, fill: rgb({}, {}, {}, {}))",
                    format_f64(layer_width),
                    format_f64(layer_height),
                    format_f64(radius),
                    shadow.color.r,
                    shadow.color.g,
                    shadow.color.b,
                    alpha,
                );
            }
            // An ellipse has no corner for a join to turn, so its silhouette
            // is the shape scaled by the expansion on both axes.
            ShapeKind::Ellipse => {
                let _ = write!(
                    out,
                    "#ellipse(width: {}pt, height: {}pt, fill: rgb({}, {}, {}, {}))",
                    format_f64(layer_width),
                    format_f64(layer_height),
                    shadow.color.r,
                    shadow.color.g,
                    shadow.color.b,
                    alpha,
                );
            }
            // Polygon, Path and Line are handled above; any future variants
            // gracefully skip the shadow rather than panicking.
            _ => {}
        }
        out.push_str("]\n");
    }
}

/// Write the offset stroke cast by every unclosed custom-geometry subpath.
///
/// Typst closes an open curve for filling, so this copy deliberately has no
/// fill and retains the source line's width, dash, and join. That leaves an
/// elbow connector open while painting the grey band PowerPoint places under
/// each of its legs (issues #1205, #1305).
fn write_crisp_open_path_shadow(
    out: &mut String,
    subpaths: &[Subpath],
    width: f64,
    height: f64,
    stroke: &Option<BorderSide>,
    shadow: &Shadow,
    offset: (f64, f64),
) {
    let Some(stroke) = stroke
        .as_ref()
        .filter(|stroke| stroke.style != BorderLineStyle::None && stroke.width > 0.0)
    else {
        return;
    };
    if !subpaths
        .iter()
        .any(|subpath| !subpath.closed && subpath.vertices.len() >= 2)
    {
        return;
    }

    let paint: String = rgb_with_alpha(&shadow.color, shadow_alpha(shadow));
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[#curve(stroke: {}",
        format_f64(offset.0),
        format_f64(offset.1),
        drawingml_stroke_value_with_paint(stroke, &paint),
    );
    for subpath in subpaths
        .iter()
        .filter(|subpath| !subpath.closed && subpath.vertices.len() >= 2)
    {
        write_curve_subpath(out, width, height, subpath);
    }
    out.push_str(")]\n");
}

/// Write one filtered SVG whose alpha is the silhouette convolved with a
/// Gaussian. Typst (0.14 and 0.15 alike) rasterises SVG filters at four samples per output
/// point, so the PDF receives a continuous 288-DPI ramp instead of one flat
/// alpha plateau per duplicate (issue #1309).
pub(super) fn write_blurred_shadow_asset(
    out: &mut String,
    ctx: &mut GenCtx,
    kind: &ShapeKind,
    size: (f64, f64),
    stroke: &Option<BorderSide>,
    shadow: &Shadow,
    offset: (f64, f64),
) {
    let (width, height): (f64, f64) = size;
    let sigma: f64 = shadow_blur_sigma(shadow);
    if sigma <= 0.0 {
        return;
    }
    let source_stroke: Option<&BorderSide> = stroke
        .as_ref()
        .filter(|stroke| stroke.style != BorderLineStyle::None && stroke.width > 0.0);
    let source_outset: f64 = source_stroke.map_or(0.0, |stroke| {
        let half_width: f64 = stroke.width / 2.0;
        if stroke.join == LineJoin::Miter {
            // SVG's default miter limit can carry an acute join beyond
            // the half-width reached by round and bevel joins. Extra
            // transparent canvas prevents that source from being clipped
            // before the Gaussian filter sees it.
            4.0 * stroke.width
        } else {
            half_width
        }
    });
    let reach: f64 = SHADOW_BLUR_EXTENT_SIGMA * sigma;
    let padding: f64 = source_outset + reach;
    let asset_width: f64 = width + 2.0 * padding;
    let asset_height: f64 = height + 2.0 * padding;
    if !(asset_width.is_finite()
        && asset_height.is_finite()
        && asset_width > 0.0
        && asset_height > 0.0)
    {
        return;
    }

    let Some(body) = shadow_svg_body(kind, width, height, padding, source_stroke.is_some()) else {
        return;
    };
    let opacity: f64 = shadow.opacity.clamp(0.0, 1.0);
    let mut svg = format!(
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{}\" height=\"{}\" viewBox=\"0 0 {} {}\"><defs><filter id=\"blur\" filterUnits=\"userSpaceOnUse\" x=\"0\" y=\"0\" width=\"{}\" height=\"{}\" color-interpolation-filters=\"sRGB\"><feGaussianBlur stdDeviation=\"{}\"/></filter></defs><g filter=\"url(#blur)\" opacity=\"{}\" fill=\"rgb({}, {}, {})\"",
        format_f64(asset_width),
        format_f64(asset_height),
        format_f64(asset_width),
        format_f64(asset_height),
        format_f64(asset_width),
        format_f64(asset_height),
        format_f64(sigma),
        format_f64(opacity),
        shadow.color.r,
        shadow.color.g,
        shadow.color.b,
    );
    if let Some(stroke) = source_stroke {
        let join = match stroke.join {
            LineJoin::Round => "round",
            LineJoin::Bevel => "bevel",
            LineJoin::Miter => "miter",
        };
        let _ = write!(
            svg,
            " stroke=\"rgb({}, {}, {})\" stroke-width=\"{}\" stroke-linejoin=\"{}\"",
            shadow.color.r,
            shadow.color.g,
            shadow.color.b,
            format_f64(stroke.width.max(0.0)),
            join,
        );
    } else {
        svg.push_str(" stroke=\"none\"");
    }
    svg.push('>');
    svg.push_str(&body);
    svg.push_str("</g></svg>");

    let path: String = ctx.add_generated_svg(svg.into_bytes());
    let _ = writeln!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[#pdf.artifact(image(\"{}\", width: {}pt, height: {}pt))]",
        format_f64(offset.0 - padding),
        format_f64(offset.1 - padding),
        path,
        format_f64(asset_width),
        format_f64(asset_height),
    );
}

/// The unblurred silhouette inside a padded shadow SVG.
fn shadow_svg_body(
    kind: &ShapeKind,
    width: f64,
    height: f64,
    origin: f64,
    include_open_strokes: bool,
) -> Option<String> {
    match kind {
        ShapeKind::Rectangle => Some(format!(
            "<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\"/>",
            format_f64(origin),
            format_f64(origin),
            format_f64(width),
            format_f64(height),
        )),
        ShapeKind::RoundedRectangle { radius_fraction } => {
            let radius: f64 =
                (radius_fraction * width.min(height)).clamp(0.0, 0.5 * width.min(height));
            Some(format!(
                "<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" rx=\"{}\" ry=\"{}\"/>",
                format_f64(origin),
                format_f64(origin),
                format_f64(width),
                format_f64(height),
                format_f64(radius),
                format_f64(radius),
            ))
        }
        ShapeKind::Ellipse => Some(format!(
            "<ellipse cx=\"{}\" cy=\"{}\" rx=\"{}\" ry=\"{}\"/>",
            format_f64(origin + width / 2.0),
            format_f64(origin + height / 2.0),
            format_f64(width / 2.0),
            format_f64(height / 2.0),
        )),
        ShapeKind::Polygon { vertices } => {
            if vertices.len() < 3 {
                return None;
            }
            let points: String = vertices
                .iter()
                .map(|(x, y)| {
                    format!(
                        "{},{}",
                        format_f64(origin + x * width),
                        format_f64(origin + y * height),
                    )
                })
                .collect::<Vec<_>>()
                .join(" ");
            Some(format!("<polygon points=\"{points}\"/>"))
        }
        ShapeKind::Path { subpaths } => {
            let mut closed_data = String::new();
            let mut open_data = String::new();
            for subpath in subpaths
                .iter()
                .filter(|subpath| subpath.closed || include_open_strokes)
            {
                let Some((first, rest)) = subpath.vertices.split_first() else {
                    continue;
                };
                let data: &mut String = if subpath.closed {
                    &mut closed_data
                } else {
                    &mut open_data
                };
                let _ = write!(
                    data,
                    "M {} {}",
                    format_f64(origin + first.0 * width),
                    format_f64(origin + first.1 * height),
                );
                for (x, y) in rest {
                    let _ = write!(
                        data,
                        " L {} {}",
                        format_f64(origin + x * width),
                        format_f64(origin + y * height),
                    );
                }
                if subpath.closed {
                    data.push_str(" Z ");
                }
            }
            let mut body = String::new();
            if !closed_data.is_empty() {
                let _ = write!(body, "<path d=\"{closed_data}\" fill-rule=\"evenodd\"/>");
            }
            if !open_data.is_empty() {
                let _ = write!(body, "<path d=\"{open_data}\" fill=\"none\"/>");
            }
            (!body.is_empty()).then_some(body)
        }
        ShapeKind::Line { .. } | ShapeKind::Polyline { .. } => None,
    }
}

/// One closed ring of a shadow silhouette, in the shape's own frame.
struct OutlineRing {
    vertices: Vec<(f64, f64)>,
    /// A ring the fill rule leaves empty. Dilating the shape shrinks its
    /// holes, so a hole offsets the other way.
    hole: bool,
}

/// How far a crisp shadow silhouette sits from the fill path.
struct RingGeometry {
    /// The signed offset from the fill path contributed by the outline.
    expansion: f64,
    /// The arc the outline's join leaves on the silhouette itself.
    silhouette_radius: f64,
}

/// The closed rings a shape's shadow silhouette is made of, or `None` for a
/// kind whose silhouette is a box or an ellipse instead.
///
/// Only a closed subpath is a filled region: an unclosed polyline is stroked,
/// and filling it would paint the area an elbow connector merely brackets
/// (issues #1205, #1305).
fn shadow_outline_rings(kind: &ShapeKind, width: f64, height: f64) -> Option<Vec<OutlineRing>> {
    let rings: Vec<Vec<(f64, f64)>> = match kind {
        ShapeKind::Polygon { vertices } => vec![scale_vertices(vertices, width, height)],
        ShapeKind::Path { subpaths } => subpaths
            .iter()
            .filter(|subpath| subpath.closed)
            .map(|subpath| scale_vertices(&subpath.vertices, width, height))
            .collect(),
        _ => return None,
    };
    // Under the even-odd rule a ring is a hole exactly when an odd number of
    // the others enclose it.
    Some(
        rings
            .iter()
            .enumerate()
            .map(|(index, vertices)| OutlineRing {
                hole: vertices.first().is_some_and(|&point| {
                    rings
                        .iter()
                        .enumerate()
                        .filter(|&(other, _)| other != index)
                        .filter(|(_, ring)| ring_contains(ring, point))
                        .count()
                        % 2
                        == 1
                }),
                vertices: vertices.clone(),
            })
            .collect(),
    )
}

fn scale_vertices(vertices: &[(f64, f64)], width: f64, height: f64) -> Vec<(f64, f64)> {
    vertices
        .iter()
        .map(|(vx, vy)| (vx * width, vy * height))
        .collect()
}

/// Whether `point` is inside `ring`, by crossing count.
fn ring_contains(ring: &[(f64, f64)], point: (f64, f64)) -> bool {
    let mut inside: bool = false;
    for index in 0..ring.len() {
        let (ax, ay): (f64, f64) = ring[index];
        let (bx, by): (f64, f64) = ring[(index + 1) % ring.len()];
        if (ay > point.1) != (by > point.1) && point.0 < (bx - ax) * (point.1 - ay) / (by - ay) + ax
        {
            inside = !inside;
        }
    }
    inside
}

/// Write a crisp shadow as the silhouette's outline offset by
/// `geometry.expansion` — a Minkowski dilation, not a copy scaled onto an
/// expanded bounding box (issue #1206).
///
/// The coordinates are written relative to the outline's own bounding box so
/// none of them goes negative, and `#place` carries the difference.
fn write_offset_ring_layer(
    out: &mut String,
    rings: &[OutlineRing],
    geometry: RingGeometry,
    shadow: &Shadow,
    alpha: u8,
    origin: (f64, f64),
) {
    // A crisp shadow is the silhouette itself, turned by whichever join the
    // outline declares (#1090).
    let reach: CornerReach = if geometry.silhouette_radius > 0.0 {
        CornerReach::Round
    } else {
        CornerReach::Mitre
    };
    let corners: Vec<Vec<OffsetCorner>> = rings
        .iter()
        .map(|ring| {
            let sign: f64 = if ring.hole { -1.0 } else { 1.0 };
            offset_ring(&ring.vertices, sign * geometry.expansion, reach)
        })
        .filter(|corners| !corners.is_empty())
        .collect();
    if corners.is_empty() {
        return;
    }

    let segments: Vec<Vec<RingSegment>> = corners
        .iter()
        .map(|corners| ring_segments(corners))
        .collect();
    let (left, top): (f64, f64) = segments
        .iter()
        .flatten()
        .flat_map(RingSegment::points)
        .fold((f64::MAX, f64::MAX), |(left, top), (x, y)| {
            (left.min(x), top.min(y))
        });
    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[#curve(fill-rule: \"even-odd\", fill: rgb({}, {}, {}, {})",
        format_f64(origin.0 + left),
        format_f64(origin.1 + top),
        shadow.color.r,
        shadow.color.g,
        shadow.color.b,
        alpha,
    );
    for ring in &segments {
        for segment in ring {
            segment.write(out, (left, top));
        }
        out.push_str(", curve.close()");
    }
    out.push_str(")]\n");
}

/// One drawn step of a ring's outline.
enum RingSegment {
    Move((f64, f64)),
    Line((f64, f64)),
    Cubic((f64, f64), (f64, f64), (f64, f64)),
}

impl RingSegment {
    fn points(&self) -> Vec<(f64, f64)> {
        match self {
            RingSegment::Move(point) | RingSegment::Line(point) => vec![*point],
            // A cubic never leaves its own control hull, so bounding the four
            // points bounds the curve.
            RingSegment::Cubic(first, second, end) => vec![*first, *second, *end],
        }
    }

    fn write(&self, out: &mut String, offset: (f64, f64)) {
        let at = |(x, y): (f64, f64)| -> String {
            format!(
                "({}pt, {}pt)",
                format_f64(x - offset.0),
                format_f64(y - offset.1),
            )
        };
        let _ = match self {
            RingSegment::Move(point) => write!(out, ", curve.move({})", at(*point)),
            RingSegment::Line(point) => write!(out, ", curve.line({})", at(*point)),
            RingSegment::Cubic(first, second, end) => write!(
                out,
                ", curve.cubic({}, {}, {})",
                at(*first),
                at(*second),
                at(*end),
            ),
        };
    }
}

/// The drawn steps of one offset ring: a straight run into every corner, and
/// the arc that turns it.
fn ring_segments(corners: &[OffsetCorner]) -> Vec<RingSegment> {
    let mut segments: Vec<RingSegment> = Vec::with_capacity(2 * corners.len());
    for (index, corner) in corners.iter().enumerate() {
        segments.push(if index == 0 {
            RingSegment::Move(corner.entry)
        } else {
            RingSegment::Line(corner.entry)
        });
        if let Some(arc) = &corner.arc {
            segments.extend(
                arc_beziers(arc)
                    .into_iter()
                    .map(|(first, second, end)| RingSegment::Cubic(first, second, end)),
            );
        }
    }
    segments
}

/// Write fill color, using rgb with 4 args when opacity is set, rgb with 3 args otherwise.
pub(super) fn write_fill_color(out: &mut String, fill: &Color, opacity: Option<f64>) {
    if let Some(op) = opacity {
        let alpha = (op * 255.0).round() as u8;
        let _ = write!(out, ", fill: {}", rgb_with_alpha(fill, alpha));
    } else {
        let _ = write!(out, ", fill: {}", rgb(fill));
    }
}

fn write_shape_params(out: &mut String, shape: &Shape, width: f64, height: f64) {
    let _ = write!(
        out,
        "width: {}pt, height: {}pt",
        format_f64(width),
        format_f64(height),
    );
    if let Some(pattern) = &shape.pattern_fill {
        out.push_str(", fill: ");
        write_pattern_fill(out, pattern);
    } else if let Some(gradient) = &shape.gradient_fill {
        out.push_str(", fill: ");
        write_gradient_fill(out, gradient);
    } else if let Some(fill) = &shape.fill {
        write_fill_color(out, fill, shape.opacity);
    }
    write_shape_stroke(out, &shape.stroke);
}

/// Write stroke parameter for shapes, handling dash patterns.
pub(super) fn write_shape_stroke(out: &mut String, stroke: &Option<BorderSide>) {
    match stroke {
        Some(stroke) if stroke.style == BorderLineStyle::None => {
            out.push_str(", stroke: none");
        }
        Some(stroke) => {
            // DrawingML dashes scale with the line width (issue #678).
            let _ = write!(out, ", stroke: {}", drawingml_stroke_value(stroke));
        }
        // An omitted source line keeps Typst's existing default behavior.
        // Explicit no-line is carried by `BorderLineStyle::None` instead.
        None => {}
    }
}

/// Write a border stroke value for image box wrapping (no leading comma).
///
/// A picture's border comes from its own `a:ln`, whose `a:prstDash` is the
/// same DrawingML preset a shape's is, so it takes the width-proportional
/// dashes too (issue #678).
pub(super) fn write_image_border_stroke(out: &mut String, stroke: &BorderSide) {
    out.push_str(&drawingml_stroke_value(stroke));
}

/// Write polygon vertex coordinates scaled to actual dimensions.
fn write_polygon_vertices(out: &mut String, width: f64, height: f64, vertices: &[(f64, f64)]) {
    for (i, (vx, vy)) in vertices.iter().enumerate() {
        if i > 0 {
            out.push_str(", ");
        }
        let _ = write!(
            out,
            "({}pt, {}pt)",
            format_f64(vx * width),
            format_f64(vy * height),
        );
    }
}

/// Generate a Typst `#curve(...)` for a shape made of several closed
/// subpaths, filled under the even-odd rule.
///
/// A DrawingML `a:custGeom` is one path however many subpaths it holds, so an
/// inner boundary carves a hole. Drawing each subpath as its own filled
/// polygon painted that hole solid (issue #870).
fn write_subpath_curve(
    out: &mut String,
    shape: &Shape,
    width: f64,
    height: f64,
    subpaths: &[Subpath],
) {
    out.push_str("#curve(fill-rule: \"even-odd\"");
    if let Some(pattern) = &shape.pattern_fill {
        out.push_str(", fill: ");
        write_pattern_fill(out, pattern);
    } else if let Some(gradient) = &shape.gradient_fill {
        out.push_str(", fill: ");
        write_gradient_fill(out, gradient);
    } else if let Some(fill) = &shape.fill {
        write_fill_color(out, fill, shape.opacity);
    }
    write_shape_stroke(out, &shape.stroke);
    for subpath in subpaths {
        write_curve_subpath(out, width, height, subpath);
    }
    out.push_str(")\n");
}

/// One subpath as `curve.move` / `curve.line` …, closed with `curve.close`
/// only when the geometry said `a:close`.
///
/// Typst closes an open curve for filling but not for stroking, which is what
/// DrawingML does: an unclosed connector's outline stops at its last point
/// (issue #1205).
fn write_curve_subpath(out: &mut String, width: f64, height: f64, subpath: &Subpath) {
    for (index, (vx, vy)) in subpath.vertices.iter().enumerate() {
        let verb: &str = if index == 0 { "move" } else { "line" };
        let _ = write!(
            out,
            ", curve.{}(({}pt, {}pt))",
            verb,
            format_f64(vx * width),
            format_f64(vy * height),
        );
    }
    if subpath.closed {
        out.push_str(", curve.close()");
    }
}

/// Generate a Typst `#polygon(...)` for an arbitrary polygon shape.
fn write_polygon(
    out: &mut String,
    shape: &Shape,
    width: f64,
    height: f64,
    vertices: &[(f64, f64)],
) {
    out.push_str("#polygon(");
    write_polygon_vertices(out, width, height, vertices);
    if let Some(pattern) = &shape.pattern_fill {
        out.push_str(", fill: ");
        write_pattern_fill(out, pattern);
    } else if let Some(gradient) = &shape.gradient_fill {
        out.push_str(", fill: ");
        write_gradient_fill(out, gradient);
    } else if let Some(fill) = &shape.fill {
        write_fill_color(out, fill, shape.opacity);
    }
    write_shape_stroke(out, &shape.stroke);
    out.push_str(")\n");
}

/// Write a Typst `gradient.linear(...)` expression.
///
/// Stops are sorted by position before rendering because Typst requires
/// gradient stop offsets to be in monotonic (non-decreasing) order.
/// The first stop is clamped to 0% and the last to 100% as Typst requires.
pub(super) fn write_gradient_fill(out: &mut String, gradient: &GradientFill) {
    // Typst requires at least 2 stops for gradient.linear().
    // Fall back to solid fill if fewer than 2 stops.
    if gradient.stops.len() < 2 {
        if let Some(stop) = gradient.stops.first() {
            out.push_str(&rgb(&stop.color));
        }
        return;
    }
    let mut sorted_stops = gradient.stops.clone();
    sorted_stops.sort_by(|a, b| {
        a.position
            .partial_cmp(&b.position)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    // Typst requires first stop at 0% and last stop at 100%.
    if let Some(first) = sorted_stops.first_mut() {
        first.position = 0.0;
    }
    if let Some(last) = sorted_stops.last_mut() {
        last.position = 1.0;
    }
    out.push_str("gradient.linear(");
    for (i, stop) in sorted_stops.iter().enumerate() {
        if i > 0 {
            out.push_str(", ");
        }
        let pos_pct = (stop.position * 100.0).round() as i64;
        let _ = write!(out, "({}, {}%)", rgb(&stop.color), pos_pct);
    }
    if gradient.angle.abs() > 0.001 {
        let _ = write!(out, ", angle: {}deg", format_f64(gradient.angle));
    }
    // typst 0.14 ignored the Oklab default in PDF and blended every gradient in
    // sRGB; typst 0.15 honours it, moving a red-to-green midpoint from
    // (127,128,0) to (208,168,0). Pinning sRGB keeps fills from recolouring.
    out.push_str(", space: rgb");
    out.push(')');
}

/// Write a Typst tiling that approximates a DrawingML preset-pattern fill.
///
/// The line-family presets use vector strokes. The remaining presets use
/// compact vector motifs so they stay resolution-independent in the PDF.
pub(super) fn write_pattern_fill(out: &mut String, pattern: &PatternFill) {
    use PatternPreset::*;

    let tile_size = match pattern.preset {
        LightHorizontal | LightVertical | LightDownwardDiagonal | LightUpwardDiagonal => 2.72,
        NarrowHorizontal | NarrowVertical => 2.0,
        WideDownwardDiagonal | WideUpwardDiagonal | LargeCheck | LargeGrid | LargeConfetti
        | HorizontalBrick | DiagonalBrick | SolidDiamond | OpenDiamond | DottedDiamond | Plaid
        | Sphere | Weave | Divot | Shingle | Wave | Trellis | ZigZag => 8.0,
        _ => 4.0,
    };
    let size = format_f64(tile_size);
    let _ = write!(out, "tiling(size: ({size}pt, {size}pt))[");
    let _ = write!(
        out,
        "#place(rect(width: 100%, height: 100%, fill: {}, stroke: none))",
        rgb(&pattern.background),
    );

    match pattern.preset {
        Percent5 => write_percentage_motif(out, &pattern.foreground, 5),
        Percent10 => write_percentage_motif(out, &pattern.foreground, 10),
        Percent20 => write_percentage_motif(out, &pattern.foreground, 20),
        Percent25 => write_percentage_motif(out, &pattern.foreground, 25),
        Percent30 => write_percentage_motif(out, &pattern.foreground, 30),
        Percent40 => write_percentage_motif(out, &pattern.foreground, 40),
        Percent50 => write_percentage_motif(out, &pattern.foreground, 50),
        Percent60 => write_percentage_motif(out, &pattern.foreground, 60),
        Percent70 => write_percentage_motif(out, &pattern.foreground, 70),
        Percent75 => write_percentage_motif(out, &pattern.foreground, 75),
        Percent80 => write_percentage_motif(out, &pattern.foreground, 80),
        Percent90 => write_percentage_motif(out, &pattern.foreground, 90),
        Horizontal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 50%)",
            "(100%, 50%)",
            0.75,
            false,
        ),
        Vertical => write_pattern_line(
            out,
            &pattern.foreground,
            "(50%, 0%)",
            "(50%, 100%)",
            0.75,
            false,
        ),
        LightHorizontal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 50%)",
            "(100%, 50%)",
            0.24,
            false,
        ),
        LightVertical => write_pattern_line(
            out,
            &pattern.foreground,
            "(50%, 0%)",
            "(50%, 100%)",
            0.24,
            false,
        ),
        DarkHorizontal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 50%)",
            "(100%, 50%)",
            1.5,
            false,
        ),
        DarkVertical => write_pattern_line(
            out,
            &pattern.foreground,
            "(50%, 0%)",
            "(50%, 100%)",
            1.5,
            false,
        ),
        NarrowHorizontal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 50%)",
            "(100%, 50%)",
            0.75,
            false,
        ),
        NarrowVertical => write_pattern_line(
            out,
            &pattern.foreground,
            "(50%, 0%)",
            "(50%, 100%)",
            0.75,
            false,
        ),
        DashedHorizontal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 50%)",
            "(100%, 50%)",
            0.75,
            true,
        ),
        DashedVertical => write_pattern_line(
            out,
            &pattern.foreground,
            "(50%, 0%)",
            "(50%, 100%)",
            0.75,
            true,
        ),
        Cross | SmallGrid | LargeGrid | Plaid => {
            write_pattern_line(
                out,
                &pattern.foreground,
                "(0%, 50%)",
                "(100%, 50%)",
                0.75,
                false,
            );
            write_pattern_line(
                out,
                &pattern.foreground,
                "(50%, 0%)",
                "(50%, 100%)",
                0.75,
                false,
            );
        }
        DownwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 0%)",
            "(100%, 100%)",
            0.75,
            false,
        ),
        UpwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 100%)",
            "(100%, 0%)",
            0.75,
            false,
        ),
        LightDownwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 0%)",
            "(100%, 100%)",
            0.24,
            false,
        ),
        LightUpwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 100%)",
            "(100%, 0%)",
            0.24,
            false,
        ),
        DarkDownwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 0%)",
            "(100%, 100%)",
            1.5,
            false,
        ),
        DarkUpwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 100%)",
            "(100%, 0%)",
            1.5,
            false,
        ),
        WideDownwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 0%)",
            "(100%, 100%)",
            1.0,
            false,
        ),
        WideUpwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 100%)",
            "(100%, 0%)",
            1.0,
            false,
        ),
        DashedDownwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 0%)",
            "(100%, 100%)",
            0.75,
            true,
        ),
        DashedUpwardDiagonal => write_pattern_line(
            out,
            &pattern.foreground,
            "(0%, 100%)",
            "(100%, 0%)",
            0.75,
            true,
        ),
        DiagonalCross | Trellis => {
            write_pattern_line(
                out,
                &pattern.foreground,
                "(0%, 0%)",
                "(100%, 100%)",
                0.75,
                false,
            );
            write_pattern_line(
                out,
                &pattern.foreground,
                "(0%, 100%)",
                "(100%, 0%)",
                0.75,
                false,
            );
        }
        SmallCheck | LargeCheck => write_checker_motif(out, &pattern.foreground, tile_size),
        DotGrid | SmallConfetti | LargeConfetti | Sphere | Divot => {
            write_dot_motif(
                out,
                &pattern.foreground,
                tile_size,
                matches!(pattern.preset, LargeConfetti | Sphere),
            );
        }
        HorizontalBrick => write_brick_motif(out, &pattern.foreground, tile_size),
        DiagonalBrick => {
            write_pattern_line(
                out,
                &pattern.foreground,
                "(0%, 0%)",
                "(100%, 100%)",
                0.75,
                false,
            );
            write_pattern_line(
                out,
                &pattern.foreground,
                "(50%, 0%)",
                "(100%, 50%)",
                0.75,
                false,
            );
        }
        SolidDiamond => write_diamond_motif(out, &pattern.foreground, tile_size, true),
        OpenDiamond | DottedDiamond => {
            write_diamond_motif(out, &pattern.foreground, tile_size, false)
        }
        Weave => {
            write_pattern_line(
                out,
                &pattern.foreground,
                "(0%, 25%)",
                "(100%, 25%)",
                1.0,
                false,
            );
            write_pattern_line(
                out,
                &pattern.foreground,
                "(25%, 0%)",
                "(25%, 100%)",
                1.0,
                false,
            );
        }
        Shingle | Wave | ZigZag => write_zigzag_motif(out, &pattern.foreground, tile_size),
    }

    out.push(']');
}

fn write_pattern_line(
    out: &mut String,
    color: &Color,
    start: &str,
    end: &str,
    thickness: f64,
    dashed: bool,
) {
    if dashed {
        let _ = write!(
            out,
            "#place(line(start: {start}, end: {end}, stroke: (paint: {}, thickness: {}pt, dash: \"dashed\")))",
            rgb(color),
            format_f64(thickness),
        );
    } else {
        let _ = write!(
            out,
            "#place(line(start: {start}, end: {end}, stroke: {}pt + {}))",
            format_f64(thickness),
            rgb(color),
        );
    }
}

fn write_percentage_motif(out: &mut String, color: &Color, percentage: usize) {
    // A 4x4 Bayer order keeps low and high densities evenly distributed.
    const BAYER_ORDER: [(usize, usize); 16] = [
        (0, 0),
        (2, 2),
        (2, 0),
        (0, 2),
        (1, 1),
        (3, 3),
        (3, 1),
        (1, 3),
        (1, 0),
        (3, 2),
        (3, 0),
        (1, 2),
        (0, 1),
        (2, 3),
        (2, 1),
        (0, 3),
    ];
    let count = ((percentage * BAYER_ORDER.len() + 50) / 100).clamp(1, BAYER_ORDER.len());
    for &(x, y) in BAYER_ORDER.iter().take(count) {
        let _ = write!(
            out,
            "#place(dx: {}pt, dy: {}pt, rect(width: 1pt, height: 1pt, fill: {}, stroke: none))",
            x,
            y,
            rgb(color),
        );
    }
}

fn write_checker_motif(out: &mut String, color: &Color, tile_size: f64) {
    let half = tile_size / 2.0;
    let half = format_f64(half);
    for (x, y) in [("0", "0"), (&half, &half)] {
        let _ = write!(
            out,
            "#place(dx: {x}pt, dy: {y}pt, rect(width: {half}pt, height: {half}pt, fill: {}, stroke: none))",
            rgb(color),
        );
    }
}

fn write_dot_motif(out: &mut String, color: &Color, tile_size: f64, large: bool) {
    let radius = if large { 1.0 } else { 0.5 };
    let center = format_f64(tile_size / 2.0 - radius);
    let _ = write!(
        out,
        "#place(dx: {center}pt, dy: {center}pt, circle(radius: {}pt, fill: {}, stroke: none))",
        format_f64(radius),
        rgb(color),
    );
}

fn write_brick_motif(out: &mut String, color: &Color, tile_size: f64) {
    write_pattern_line(out, color, "(0%, 0%)", "(100%, 0%)", 0.75, false);
    write_pattern_line(out, color, "(0%, 50%)", "(100%, 50%)", 0.75, false);
    let center = format_f64(tile_size / 2.0);
    let _ = write!(
        out,
        "#place(line(start: ({center}pt, 0pt), end: ({center}pt, {}pt), stroke: 0.75pt + {}))",
        format_f64(tile_size / 2.0),
        rgb(color),
    );
}

fn write_diamond_motif(out: &mut String, color: &Color, tile_size: f64, solid: bool) {
    let half = format_f64(tile_size / 2.0);
    let size = format_f64(tile_size);
    let fill = if solid {
        rgb(color)
    } else {
        "none".to_string()
    };
    let stroke = if solid {
        "none".to_string()
    } else {
        format!("0.75pt + {}", rgb(color))
    };
    let _ = write!(
        out,
        "#place(polygon(({half}pt, 0pt), ({size}pt, {half}pt), ({half}pt, {size}pt), (0pt, {half}pt), fill: {fill}, stroke: {stroke}))",
    );
}

fn write_zigzag_motif(out: &mut String, color: &Color, tile_size: f64) {
    let half = format_f64(tile_size / 2.0);
    let size = format_f64(tile_size);
    let _ = write!(
        out,
        "#place(curve(stroke: 0.75pt + {}, fill: none, curve.move((0pt, {half}pt)), curve.line(({half}pt, 0pt)), curve.line(({size}pt, {half}pt))))",
        rgb(color),
    );
}

// ── Polyline & arrowhead rendering ──────────────────────────────────

/// Render a multi-segment polyline as consecutive `#line()` calls,
/// each wrapped in `#place(top + left)` so they overlay at the same origin.
fn write_polyline(out: &mut String, stroke: &Option<BorderSide>, points: &[(f64, f64)]) {
    for segment in points.windows(2) {
        let (x1, y1) = segment[0];
        let (x2, y2) = segment[1];
        out.push_str("#place(top + left)[#line(");
        let _ = write!(
            out,
            "start: ({}pt, {}pt), end: ({}pt, {}pt)",
            format_f64(x1),
            format_f64(y1),
            format_f64(x2),
            format_f64(y2),
        );
        write_shape_stroke(out, stroke);
        out.push_str(")]\n");
    }
}

/// Draw a triangle arrowhead at `tip`, pointing in the direction from `from` → `tip`.
pub(super) fn write_arrowhead_at(
    out: &mut String,
    stroke: &Option<BorderSide>,
    from: (f64, f64),
    tip: (f64, f64),
) {
    let Some(stroke) = stroke
        .as_ref()
        .filter(|stroke| stroke.style != BorderLineStyle::None)
    else {
        return;
    };
    let dx: f64 = tip.0 - from.0;
    let dy: f64 = tip.1 - from.1;
    let len: f64 = (dx * dx + dy * dy).sqrt();
    if len < 0.001 {
        return;
    }
    // Arrow size proportional to stroke width, with min/max bounds.
    let arrow_len: f64 = (stroke.width * 4.0).clamp(3.0, 12.0);
    let arrow_half_w: f64 = arrow_len * 0.45;

    // Unit direction vector from `from` toward `tip`.
    let ux: f64 = dx / len;
    let uy: f64 = dy / len;
    // Perpendicular vector.
    let px: f64 = -uy;
    let py: f64 = ux;

    // Three vertices: tip, and two base corners.
    let base_x: f64 = tip.0 - ux * arrow_len;
    let base_y: f64 = tip.1 - uy * arrow_len;
    let v1 = (tip.0, tip.1);
    let v2 = (base_x + px * arrow_half_w, base_y + py * arrow_half_w);
    let v3 = (base_x - px * arrow_half_w, base_y - py * arrow_half_w);

    out.push_str("#place(top + left)[#polygon(");
    let _ = write!(
        out,
        "({}pt, {}pt), ({}pt, {}pt), ({}pt, {}pt), fill: {}",
        format_f64(v1.0),
        format_f64(v1.1),
        format_f64(v2.0),
        format_f64(v2.1),
        format_f64(v3.0),
        format_f64(v3.1),
        rgb(&stroke.color),
    );
    out.push_str(")]\n");
}

/// Render a non-rectangular shape background for a text box.
///
/// Emits a `#place(top + left)` overlay with the shape geometry, offset by
/// negative insets so it covers the full bounding box (the text box block's
/// coordinate origin is inside the inset).
#[allow(clippy::too_many_arguments)]
pub(super) fn write_text_box_shape_background(
    out: &mut String,
    shape_kind: &ShapeKind,
    width: f64,
    height: f64,
    padding: &Insets,
    fill: Option<&Color>,
    opacity: Option<f64>,
    stroke: &Option<BorderSide>,
) {
    // Offset the placed shape to compensate for the block's inset.
    let _ = write!(
        out,
        "  #place(top + left, dx: -{}pt, dy: -{}pt)[",
        format_f64(padding.left),
        format_f64(padding.top),
    );
    match shape_kind {
        ShapeKind::RoundedRectangle { radius_fraction } => {
            let radius: f64 = radius_fraction * width.min(height);
            let _ = write!(
                out,
                "#rect(width: {}pt, height: {}pt, radius: {}pt",
                format_f64(width),
                format_f64(height),
                format_f64(radius),
            );
            if let Some(c) = fill {
                write_fill_color(out, c, opacity);
            }
            write_shape_stroke(out, stroke);
            out.push(')');
        }
        ShapeKind::Polygon { vertices } => {
            out.push_str("#polygon(");
            write_polygon_vertices(out, width, height, vertices);
            if let Some(c) = fill {
                write_fill_color(out, c, opacity);
            }
            write_shape_stroke(out, stroke);
            out.push(')');
        }
        ShapeKind::Ellipse => {
            let _ = write!(
                out,
                "#ellipse(width: {}pt, height: {}pt",
                format_f64(width),
                format_f64(height),
            );
            if let Some(c) = fill {
                write_fill_color(out, c, opacity);
            }
            write_shape_stroke(out, stroke);
            out.push(')');
        }
        // Rectangle or line/polyline — shouldn't reach here, but handle gracefully.
        _ => {
            let _ = write!(
                out,
                "#rect(width: {}pt, height: {}pt",
                format_f64(width),
                format_f64(height),
            );
            if let Some(c) = fill {
                write_fill_color(out, c, opacity);
            }
            write_shape_stroke(out, stroke);
            out.push(')');
        }
    }
    out.push_str("]\n");
}
