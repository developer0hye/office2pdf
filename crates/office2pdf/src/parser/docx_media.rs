use super::contexts::DocxConversionContext;
use super::{
    Block, DrawingTextBoxInfo, FloatingImage, FloatingTextBox, HyperlinkMap, ImageData,
    ImageFormat, ImageMap, StyleMap, VmlTextBoxInfo, WrapContext, convert_paragraph_blocks,
    convert_table,
};

fn emu_to_pt(emu: u32) -> f64 {
    emu as f64 / 12700.0
}

fn emu_to_pt_signed(emu: i32) -> f64 {
    emu as f64 / 12700.0
}

pub(super) fn extract_drawing_image(
    drawing: &docx_rs::Drawing,
    images: &ImageMap,
    wraps: &WrapContext,
) -> Option<Block> {
    let pic = match &drawing.data {
        Some(docx_rs::DrawingData::Pic(pic)) => pic,
        _ => return None,
    };

    let data = images.get(&pic.id)?;
    let (w_emu, h_emu) = pic.size;
    let width = if w_emu > 0 {
        Some(emu_to_pt(w_emu))
    } else {
        None
    };
    let height = if h_emu > 0 {
        Some(emu_to_pt(h_emu))
    } else {
        None
    };

    let image_data = ImageData {
        data: data.clone(),
        format: ImageFormat::Png,
        width,
        height,
        crop: None,
    };

    if pic.position_type == docx_rs::DrawingPositionType::Anchor {
        let wrap_mode = wraps.consume_next();
        let offset_x = match pic.position_h {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };
        let offset_y = match pic.position_v {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };

        Some(Block::FloatingImage(FloatingImage {
            image: image_data,
            wrap_mode,
            offset_x,
            offset_y,
        }))
    } else {
        Some(Block::Image(image_data))
    }
}

pub(super) fn extract_shape_image(shape: &docx_rs::Shape, images: &ImageMap) -> Option<Block> {
    let image_id = shape.image_data.as_ref()?.id.as_str();
    let data = images.get(image_id)?;

    let width = extract_vml_style_dimension(shape.style.as_deref(), "width");
    let height = extract_vml_style_dimension(shape.style.as_deref(), "height");

    Some(Block::Image(ImageData {
        data: data.clone(),
        format: ImageFormat::Png,
        width,
        height,
        crop: None,
    }))
}

pub(super) fn extract_vml_shape_text_box(
    shape: &docx_rs::Shape,
    text_box: &VmlTextBoxInfo,
) -> Option<FloatingTextBox> {
    if text_box.paragraphs.is_empty() {
        return None;
    }

    let style = shape.style.as_deref()?;
    if !is_positioned_vml_text_box(style) {
        return None;
    }

    let width = extract_vml_style_length(Some(style), "width")?;
    let height = extract_vml_style_length(Some(style), "height")?;
    let offset_x = extract_vml_style_length(Some(style), "margin-left")
        .or_else(|| extract_vml_style_length(Some(style), "left"))
        .unwrap_or(0.0);
    let offset_y = extract_vml_style_length(Some(style), "margin-top")
        .or_else(|| extract_vml_style_length(Some(style), "top"))
        .unwrap_or(0.0);
    let wrap_mode = text_box
        .wrap_mode
        .or_else(|| extract_vml_style_wrap_mode(Some(style)))
        .unwrap_or(crate::ir::WrapMode::Square);

    Some(FloatingTextBox {
        content: text_box.clone().into_blocks(),
        wrap_mode,
        width,
        height,
        offset_x,
        offset_y,
    })
}

fn is_positioned_vml_text_box(style: &str) -> bool {
    has_vml_style_value(style, "position", "absolute")
        || extract_vml_style_length(Some(style), "margin-left").is_some()
        || extract_vml_style_length(Some(style), "margin-top").is_some()
}

fn has_vml_style_value(style: &str, key: &str, expected: &str) -> bool {
    extract_vml_style_value(style, key)
        .map(|value| value.eq_ignore_ascii_case(expected))
        .unwrap_or(false)
}

fn extract_vml_style_wrap_mode(style: Option<&str>) -> Option<crate::ir::WrapMode> {
    let value = extract_vml_style_value(style?, "mso-wrap-style")?;
    match value.to_ascii_lowercase().as_str() {
        "square" => Some(crate::ir::WrapMode::Square),
        "none" => Some(crate::ir::WrapMode::None),
        "tight" | "through" => Some(crate::ir::WrapMode::Tight),
        "top-and-bottom" | "topandbottom" => Some(crate::ir::WrapMode::TopAndBottom),
        _ => None,
    }
}

fn extract_vml_style_value(style: &str, key: &str) -> Option<String> {
    for part in style.split(';') {
        let Some((name, value)) = part.split_once(':') else {
            continue;
        };
        if name.trim() == key {
            return Some(value.trim().to_string());
        }
    }

    None
}

fn extract_vml_style_length(style: Option<&str>, key: &str) -> Option<f64> {
    let value = extract_vml_style_value(style?, key)?;
    let value = value.trim();
    if let Some(raw) = value.strip_suffix("pt") {
        return raw.trim().parse::<f64>().ok();
    }
    if let Some(raw) = value.strip_suffix("px") {
        return raw
            .trim()
            .parse::<f64>()
            .ok()
            .map(|px| px * crate::defaults::POINTS_PER_INCH / crate::defaults::PIXELS_PER_INCH);
    }

    None
}

fn extract_vml_style_dimension(style: Option<&str>, key: &str) -> Option<f64> {
    let style = style?;
    for part in style.split(';') {
        let Some((name, value)) = part.split_once(':') else {
            continue;
        };
        if name.trim() != key {
            continue;
        }

        let value = value.trim();
        if let Some(raw) = value.strip_suffix("pt") {
            return raw.trim().parse::<f64>().ok();
        }
        if let Some(raw) = value.strip_suffix("px") {
            return raw.trim().parse::<f64>().ok().map(|px| {
                px * crate::defaults::POINTS_PER_INCH / crate::defaults::PIXELS_PER_INCH
            });
        }
        if let Ok(points) = value.parse::<f64>() {
            return Some(points);
        }
    }

    None
}

pub(super) fn extract_drawing_text_box_blocks(
    drawing: &docx_rs::Drawing,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    ctx: &DocxConversionContext,
) -> Vec<Block> {
    let Some(docx_rs::DrawingData::TextBox(text_box)) = &drawing.data else {
        return Vec::new();
    };

    let layout: DrawingTextBoxInfo = ctx.drawing_text_boxes.consume_next();
    let mut blocks: Vec<Block> = Vec::new();
    for child in &text_box.children {
        match child {
            docx_rs::TextBoxContentChild::Paragraph(para) => {
                convert_paragraph_blocks(para, &mut blocks, images, hyperlinks, style_map, ctx);
            }
            docx_rs::TextBoxContentChild::Table(table) => {
                blocks.push(Block::Table(convert_table(
                    table, images, hyperlinks, style_map, ctx, 0,
                )));
            }
        }
    }

    if text_box.position_type == docx_rs::DrawingPositionType::Anchor {
        let wrap_mode = ctx.wraps.consume_next();
        let offset_x = match text_box.position_h {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };
        let offset_y = match text_box.position_v {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };
        let (width, height) = resolve_drawing_text_box_size(text_box, layout);

        vec![Block::FloatingTextBox(FloatingTextBox {
            content: blocks,
            wrap_mode,
            width,
            height,
            offset_x,
            offset_y,
        })]
    } else {
        blocks
    }
}

fn resolve_drawing_text_box_size(
    text_box: &docx_rs::TextBox,
    layout: DrawingTextBoxInfo,
) -> (f64, f64) {
    let width = layout.width_pt.unwrap_or_else(|| {
        if text_box.size.0 > 0 {
            emu_to_pt(text_box.size.0)
        } else {
            0.0
        }
    });
    let height = layout.height_pt.unwrap_or_else(|| {
        if text_box.size.1 > 0 {
            emu_to_pt(text_box.size.1)
        } else {
            0.0
        }
    });

    (width, height)
}
