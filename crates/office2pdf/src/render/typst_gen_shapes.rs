use std::fmt::Write;

use super::*;

pub(super) fn generate_shape(out: &mut String, shape: &Shape, width: f64, height: f64) {
    // Render shadow as offset duplicate before main shape
    if let Some(shadow) = &shape.shadow {
        write_shadow_shape(out, shape, width, height, shadow);
    }

    let has_rotation = shape.rotation_deg.is_some();
    if let Some(deg) = shape.rotation_deg {
        let _ = write!(out, "#rotate({}deg)[", format_f64(deg));
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
            let has_arrowheads: bool =
                *tail_end != ArrowHead::None || *head_end != ArrowHead::None;
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
                format_f64(*x1),
                format_f64(*y1),
                format_f64(*x2),
                format_f64(*y2),
            );
            write_shape_stroke(out, &shape.stroke);
            out.push_str(")\n");
            if has_arrowheads {
                out.push_str("]\n");
            }
            if *tail_end != ArrowHead::None {
                write_arrowhead_at(out, &shape.stroke, (*x1, *y1), (*x2, *y2));
            }
            if *head_end != ArrowHead::None {
                write_arrowhead_at(out, &shape.stroke, (*x2, *y2), (*x1, *y1));
            }
        }
        ShapeKind::Polyline {
            points,
            head_end,
            tail_end,
        } => {
            write_polyline(out, &shape.stroke, points);
            if points.len() >= 2 {
                if *tail_end != ArrowHead::None {
                    let last = points[points.len() - 1];
                    let second_last = points[points.len() - 2];
                    write_arrowhead_at(out, &shape.stroke, second_last, last);
                }
                if *head_end != ArrowHead::None {
                    let first = points[0];
                    let second = points[1];
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
    }

    if has_rotation {
        out.push_str("]\n");
    }
}

/// Render a shadow approximation as an offset duplicate shape with reduced opacity.
fn write_shadow_shape(out: &mut String, shape: &Shape, width: f64, height: f64, shadow: &Shadow) {
    let dir_rad = shadow.direction.to_radians();
    let dx = shadow.distance * dir_rad.cos();
    let dy = shadow.distance * dir_rad.sin();
    let alpha = (shadow.opacity * 255.0).round() as u8;

    let _ = write!(
        out,
        "#place(top + left, dx: {}pt, dy: {}pt)[",
        format_f64(dx),
        format_f64(dy),
    );

    match &shape.kind {
        ShapeKind::Line { .. } => {
            // Lines don't have meaningful shadows; skip
            out.push_str("]\n");
            return;
        }
        ShapeKind::Polygon { vertices } => {
            // Shadow for polygon: duplicate polygon with shadow color
            out.push_str("#polygon(");
            write_polygon_vertices(out, width, height, vertices);
            let _ = write!(
                out,
                ", fill: rgb({}, {}, {}, {})",
                shadow.color.r, shadow.color.g, shadow.color.b, alpha,
            );
            out.push_str(")]\n");
            return;
        }
        _ => {}
    }
    let shape_cmd = match &shape.kind {
        ShapeKind::Rectangle => "#rect(",
        ShapeKind::Ellipse => "#ellipse(",
        ShapeKind::RoundedRectangle { radius_fraction } => {
            let _ = writeln!(
                out,
                "#rect(width: {}pt, height: {}pt, radius: {}pt, fill: rgb({}, {}, {}, {}))]",
                format_f64(width),
                format_f64(height),
                format_f64(radius_fraction * width.min(height)),
                shadow.color.r,
                shadow.color.g,
                shadow.color.b,
                alpha,
            );
            return;
        }
        // Line and Polygon are handled by early returns above; any future
        // variants gracefully skip the shadow rather than panicking.
        _ => {
            out.push_str("]\n");
            return;
        }
    };
    out.push_str(shape_cmd);
    let _ = write!(
        out,
        "width: {}pt, height: {}pt, fill: rgb({}, {}, {}, {})",
        format_f64(width),
        format_f64(height),
        shadow.color.r,
        shadow.color.g,
        shadow.color.b,
        alpha,
    );
    out.push_str(")]\n");
}

/// Write fill color, using rgb with 4 args when opacity is set, rgb with 3 args otherwise.
pub(super) fn write_fill_color(out: &mut String, fill: &Color, opacity: Option<f64>) {
    if let Some(op) = opacity {
        let alpha = (op * 255.0).round() as u8;
        let _ = write!(
            out,
            ", fill: rgb({}, {}, {}, {})",
            fill.r, fill.g, fill.b, alpha
        );
    } else {
        let _ = write!(out, ", fill: rgb({}, {}, {})", fill.r, fill.g, fill.b);
    }
}

fn write_shape_params(out: &mut String, shape: &Shape, width: f64, height: f64) {
    let _ = write!(
        out,
        "width: {}pt, height: {}pt",
        format_f64(width),
        format_f64(height),
    );
    if let Some(gradient) = &shape.gradient_fill {
        out.push_str(", fill: ");
        write_gradient_fill(out, gradient);
    } else if let Some(fill) = &shape.fill {
        write_fill_color(out, fill, shape.opacity);
    }
    write_shape_stroke(out, &shape.stroke);
}

/// Write stroke parameter for shapes, handling dash patterns.
pub(super) fn write_shape_stroke(out: &mut String, stroke: &Option<BorderSide>) {
    if let Some(stroke) = stroke {
        match stroke.style {
            BorderLineStyle::Solid | BorderLineStyle::None => {
                let _ = write!(
                    out,
                    ", stroke: {}pt + rgb({}, {}, {})",
                    format_f64(stroke.width),
                    stroke.color.r,
                    stroke.color.g,
                    stroke.color.b,
                );
            }
            _ => {
                let _ = write!(
                    out,
                    ", stroke: (paint: rgb({}, {}, {}), thickness: {}pt, dash: \"{}\")",
                    stroke.color.r,
                    stroke.color.g,
                    stroke.color.b,
                    format_f64(stroke.width),
                    border_line_style_to_typst(stroke.style),
                );
            }
        }
    }
}

/// Write a border stroke value for image box wrapping (no leading comma).
pub(super) fn write_image_border_stroke(out: &mut String, stroke: &BorderSide) {
    match stroke.style {
        BorderLineStyle::Solid | BorderLineStyle::None => {
            let _ = write!(
                out,
                "{}pt + rgb({}, {}, {})",
                format_f64(stroke.width),
                stroke.color.r,
                stroke.color.g,
                stroke.color.b,
            );
        }
        _ => {
            let _ = write!(
                out,
                "(paint: rgb({}, {}, {}), thickness: {}pt, dash: \"{}\")",
                stroke.color.r,
                stroke.color.g,
                stroke.color.b,
                format_f64(stroke.width),
                border_line_style_to_typst(stroke.style),
            );
        }
    }
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
    if let Some(gradient) = &shape.gradient_fill {
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
            let _ = write!(
                out,
                "rgb({}, {}, {})",
                stop.color.r, stop.color.g, stop.color.b,
            );
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
        let _ = write!(
            out,
            "(rgb({}, {}, {}), {}%)",
            stop.color.r, stop.color.g, stop.color.b, pos_pct,
        );
    }
    if gradient.angle.abs() > 0.001 {
        let _ = write!(out, ", angle: {}deg", format_f64(gradient.angle));
    }
    out.push(')');
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
fn write_arrowhead_at(
    out: &mut String,
    stroke: &Option<BorderSide>,
    from: (f64, f64),
    tip: (f64, f64),
) {
    let Some(stroke) = stroke else { return };
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
        "({}pt, {}pt), ({}pt, {}pt), ({}pt, {}pt), fill: rgb({}, {}, {})",
        format_f64(v1.0),
        format_f64(v1.1),
        format_f64(v2.0),
        format_f64(v2.1),
        format_f64(v3.0),
        format_f64(v3.1),
        stroke.color.r,
        stroke.color.g,
        stroke.color.b,
    );
    out.push_str(")]\n");
}
