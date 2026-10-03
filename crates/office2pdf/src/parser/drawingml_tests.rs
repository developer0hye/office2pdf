use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::Event;

use super::*;
use crate::ir::{Color, ImageFormat};

fn scheme_with(colors: &[(&str, Color)], aliases: &[(&str, &str)]) -> (ColorsMap, AliasMap) {
    let colors: ColorsMap = colors
        .iter()
        .map(|(name, color)| ((*name).to_string(), *color))
        .collect();
    let aliases: AliasMap = aliases
        .iter()
        .map(|(from, to)| ((*from).to_string(), (*to).to_string()))
        .collect();
    (colors, aliases)
}

type ColorsMap = HashMap<String, Color>;
type AliasMap = HashMap<String, String>;

#[test]
fn supported_raster_formats_survive_sniffing_and_full_decode() {
    let image = image::DynamicImage::new_rgb8(2, 2);
    let formats = [
        (image::ImageFormat::Png, ImageFormat::Png),
        (image::ImageFormat::Jpeg, ImageFormat::Jpeg),
        (image::ImageFormat::Gif, ImageFormat::Gif),
        (image::ImageFormat::Bmp, ImageFormat::Bmp),
        (image::ImageFormat::Tiff, ImageFormat::Tiff),
    ];

    for (encoded_format, expected) in formats {
        let mut encoded = std::io::Cursor::new(Vec::new());
        image
            .write_to(&mut encoded, encoded_format)
            .expect("test raster encodes");
        assert_eq!(
            validated_raster_format(&encoded.into_inner()),
            Some(expected),
            "{encoded_format:?} must remain supported"
        );
    }
}

#[test]
fn malformed_raster_is_rejected_even_when_its_signature_is_present() {
    assert_eq!(validated_raster_format(b"not an image"), None);
    assert_eq!(validated_raster_format(b"\x89PNG\r\n\x1a\ntruncated"), None);
}

#[test]
fn resolve_scheme_color_direct_lookup() {
    let (colors, aliases) = scheme_with(&[("accent1", Color::new(68, 114, 196))], &[]);
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    assert_eq!(
        resolve_scheme_color(&scheme, "accent1"),
        Some(Color::new(68, 114, 196))
    );
    assert_eq!(resolve_scheme_color(&scheme, "accent2"), None);
}

#[test]
fn resolve_scheme_color_follows_alias() {
    // pptx clrMap maps bg1 → lt1 for typical slides.
    let (colors, aliases) = scheme_with(&[("lt1", Color::new(255, 255, 255))], &[("bg1", "lt1")]);
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    assert_eq!(
        resolve_scheme_color(&scheme, "bg1"),
        Some(Color::new(255, 255, 255))
    );
}

#[test]
fn resolve_scheme_color_falls_back_to_unaliased_name() {
    // Alias points at a missing entry; the raw name still resolves.
    let (colors, aliases) = scheme_with(&[("bg1", Color::new(1, 2, 3))], &[("bg1", "lt9")]);
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    assert_eq!(
        resolve_scheme_color(&scheme, "bg1"),
        Some(Color::new(1, 2, 3))
    );
}

/// `<a:clrScheme>` names its slots `dk1`/`lt1`/`dk2`/`lt2`, but the parts that
/// *use* a colour name them `tx1`/`bg1`/`tx2`/`bg2`. PowerPoint bridges the two
/// with the slide master's `<p:clrMap>`; Word and Excel carry no such part, so
/// nothing maps them and the pair has to be understood implicitly (#1145).
#[test]
fn text_and_background_names_resolve_onto_the_theme_light_dark_slots() {
    let (colors, aliases) = scheme_with(
        &[
            ("dk1", Color::new(0, 0, 0)),
            ("lt1", Color::new(255, 255, 255)),
            ("dk2", Color::new(0x44, 0x54, 0x6A)),
            ("lt2", Color::new(0xE7, 0xE6, 0xE6)),
        ],
        &[],
    );
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };

    assert_eq!(
        resolve_scheme_color(&scheme, "tx1"),
        Some(Color::new(0, 0, 0))
    );
    assert_eq!(
        resolve_scheme_color(&scheme, "bg1"),
        Some(Color::new(255, 255, 255))
    );
    assert_eq!(
        resolve_scheme_color(&scheme, "tx2"),
        Some(Color::new(0x44, 0x54, 0x6A))
    );
    assert_eq!(
        resolve_scheme_color(&scheme, "bg2"),
        Some(Color::new(0xE7, 0xE6, 0xE6))
    );
    // The implicit pairing is exactly those four names; nothing else acquires
    // a slot it was never given.
    assert_eq!(resolve_scheme_color(&scheme, "tx3"), None);
    assert_eq!(resolve_scheme_color(&scheme, "accent1"), None);
}

/// A deck whose `<p:clrMap>` swaps the pair — a light-on-dark master — must
/// keep its own mapping, not the implicit one.
#[test]
fn a_declared_alias_outranks_the_implicit_light_dark_pairing() {
    let (colors, aliases) = scheme_with(
        &[
            ("dk1", Color::new(0, 0, 0)),
            ("lt1", Color::new(255, 255, 255)),
        ],
        &[("tx1", "lt1"), ("bg1", "dk1")],
    );
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };

    assert_eq!(
        resolve_scheme_color(&scheme, "tx1"),
        Some(Color::new(255, 255, 255))
    );
    assert_eq!(
        resolve_scheme_color(&scheme, "bg1"),
        Some(Color::new(0, 0, 0))
    );
}

/// PowerPoint's fallback colour map, used when no master declares one, spells
/// every entry as itself. The identity has to fall through to the light/dark
/// slot the same way an absent map does.
#[test]
fn an_identity_alias_still_reaches_the_light_dark_slot() {
    let (colors, aliases) = scheme_with(&[("dk1", Color::new(17, 17, 17))], &[("tx1", "tx1")]);
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };

    assert_eq!(
        resolve_scheme_color(&scheme, "tx1"),
        Some(Color::new(17, 17, 17))
    );
}

/// A theme that names the slot under the using spelling is still read that
/// way: the implicit pairing is a fallback, not a redirect.
#[test]
fn a_theme_keyed_by_the_using_name_keeps_winning() {
    let (colors, aliases) = scheme_with(
        &[("tx1", Color::new(1, 2, 3)), ("dk1", Color::new(4, 5, 6))],
        &[],
    );
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };

    assert_eq!(
        resolve_scheme_color(&scheme, "tx1"),
        Some(Color::new(1, 2, 3))
    );
}

#[test]
fn tint_blends_toward_white() {
    // OOXML tint 0.4: channel = 255 - (255 - c) * 0.4
    let out = apply_color_transforms(Color::new(0, 100, 255), &[ColorTransform::Tint(0.4)]);
    assert_eq!(out, Color::new(153, 193, 255));
}

#[test]
fn shade_scales_in_linear_light_not_on_srgb_bytes() {
    // accent1 #4472C4 shaded 50% is #2F528F in PowerPoint (issue #667). Scaling
    // the sRGB bytes instead gives #223962 — visibly darker on every channel.
    // The difference is the sRGB transfer curve: the scale belongs in linear
    // light, so the bytes are decoded, halved, and re-encoded.
    let out = apply_color_transforms(Color::new(0x44, 0x72, 0xC4), &[ColorTransform::Shade(0.5)]);
    assert_eq!(out, Color::new(0x2F, 0x52, 0x8F));
}

#[test]
fn shade_of_one_is_identity() {
    // Triangulation: a full-strength shade must round-trip the colour through
    // the transfer curve unchanged, which a wrong or lossy curve would not.
    let color = Color::new(200, 100, 50);
    assert_eq!(
        apply_color_transforms(color, &[ColorTransform::Shade(1.0)]),
        color
    );
}

#[test]
fn shade_of_zero_is_black() {
    // Triangulation at the other end: no light left is black regardless of curve.
    assert_eq!(
        apply_color_transforms(Color::new(200, 100, 50), &[ColorTransform::Shade(0.0)]),
        Color::new(0, 0, 0)
    );
}

#[test]
fn shade_keeps_black_and_white_fixed() {
    // Triangulation: the curve's endpoints must be exact, or every shaded
    // colour picks up a bias.
    assert_eq!(
        apply_color_transforms(Color::new(0, 0, 0), &[ColorTransform::Shade(0.5)]),
        Color::new(0, 0, 0)
    );
    assert_eq!(
        apply_color_transforms(Color::new(255, 255, 255), &[ColorTransform::Shade(1.0)]),
        Color::new(255, 255, 255)
    );
}

#[test]
fn lum_mod_and_off_adjust_lightness_in_hsl() {
    // lumMod 0.5 halves lightness; pure red keeps its hue.
    let out = apply_color_transforms(Color::new(255, 0, 0), &[ColorTransform::LumMod(0.5)]);
    assert_eq!(out, Color::new(128, 0, 0));

    // lumOff +0.2 raises lightness toward a lighter red.
    let out = apply_color_transforms(Color::new(255, 0, 0), &[ColorTransform::LumOff(0.2)]);
    assert_eq!(out, Color::new(255, 102, 102));
}

#[test]
fn saturation_modifier_multiplies_hsl_saturation() {
    let (colors, aliases) = scheme_with(&[], &[]);
    let parsed = parse_first_color(
        r#"<a:srgbClr val="6496C8"><a:satMod val="150000"/></a:srgbClr>"#,
        &colors,
        &aliases,
    );

    // #6496C8 is HSL(210, 47.6%, 58.8%). Multiplying its saturation by
    // 150% produces HSL(210, 71.4%, 58.8%), or #4B96E1. Treating 150000 as
    // an offset or ignoring satMod cannot produce this color.
    assert_eq!(parsed.color, Some(Color::new(75, 150, 225)));
}

#[test]
fn hue_and_saturation_offsets_use_drawingml_units() {
    let (colors, aliases) = scheme_with(&[], &[]);
    let parsed = parse_first_color(
        r#"<a:srgbClr val="00FFFF"><a:hueOff val="-3600000"/><a:satOff val="-50000"/></a:srgbClr>"#,
        &colors,
        &aliases,
    );

    // -3,600,000 is -60 degrees (1/60,000 degree units), rotating cyan to
    // green; removing half its HSL saturation produces #40BF40.
    assert_eq!(parsed.color, Some(Color::new(64, 191, 64)));
}

#[test]
fn no_transforms_is_identity() {
    let color = Color::new(12, 34, 56);
    assert_eq!(apply_color_transforms(color, &[]), color);
}

fn parse_first_color(xml: &str, colors: &ColorsMap, aliases: &AliasMap) -> ParsedColor {
    let mut reader = Reader::from_str(xml);
    let scheme = SchemeColors { colors, aliases };
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                return parse_color_from_start(&mut reader, e, &scheme);
            }
            Ok(Event::Empty(ref e)) => {
                return parse_color_from_empty(e, &scheme);
            }
            Ok(Event::Eof) => panic!("no color element in fixture"),
            _ => {}
        }
    }
}

#[test]
fn parse_srgb_color_with_lum_transforms() {
    // Realistic DrawingML: accent fill darkened via lumMod as PowerPoint emits.
    let (colors, aliases) = scheme_with(&[], &[]);
    let parsed = parse_first_color(
        r#"<a:srgbClr val="FF0000"><a:lumMod val="50000"/></a:srgbClr>"#,
        &colors,
        &aliases,
    );
    assert_eq!(parsed.color, Some(Color::new(128, 0, 0)));
    assert_eq!(parsed.alpha, None);
}

#[test]
fn parse_scheme_color_with_alpha() {
    let (colors, aliases) = scheme_with(&[("accent1", Color::new(68, 114, 196))], &[]);
    let parsed = parse_first_color(
        r#"<a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr>"#,
        &colors,
        &aliases,
    );
    assert_eq!(parsed.color, Some(Color::new(68, 114, 196)));
    assert_eq!(parsed.alpha, Some(0.5));
}

#[test]
fn parse_sys_color_uses_last_clr() {
    let (colors, aliases) = scheme_with(&[], &[]);
    let parsed = parse_first_color(
        r#"<a:sysClr val="windowText" lastClr="000000"/>"#,
        &colors,
        &aliases,
    );
    assert_eq!(parsed.color, Some(Color::new(0, 0, 0)));
}

#[test]
fn theme_color_scheme_parses_non_empty_color_elements() {
    let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"></a:sysClr></a:dk1>
      <a:accent1><a:srgbClr val="4F81BD"></a:srgbClr></a:accent1>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;
    let colors = parse_theme_color_scheme(theme_xml);
    assert_eq!(colors.get("dk1"), Some(&Color::new(0, 0, 0)));
    assert_eq!(colors.get("accent1"), Some(&Color::new(0x4f, 0x81, 0xbd)));
}

#[test]
fn theme_accent_palette_returns_all_six_in_order() {
    // The Office 2007 palette both audited fixtures declare (issue #670).
    let colors: ColorsMap = [
        ("accent1", Color::new(0x4F, 0x81, 0xBD)),
        ("accent2", Color::new(0xC0, 0x50, 0x4D)),
        ("accent3", Color::new(0x9B, 0xBB, 0x59)),
        ("accent4", Color::new(0x80, 0x64, 0xA2)),
        ("accent5", Color::new(0x4B, 0xAC, 0xC6)),
        ("accent6", Color::new(0xF7, 0x96, 0x46)),
        ("dk1", Color::new(0, 0, 0)),
    ]
    .into_iter()
    .map(|(name, color)| (name.to_string(), color))
    .collect();

    let palette = theme_accent_palette(&colors);

    assert_eq!(palette.len(), 6);
    assert_eq!(palette[0], Color::new(0x4F, 0x81, 0xBD), "accent1 leads");
    assert_eq!(palette[5], Color::new(0xF7, 0x96, 0x46), "accent6 trails");
}

#[test]
fn theme_accent_palette_is_empty_when_an_accent_is_missing() {
    // Triangulation: a short list would make the renderer repeat colours the
    // file never named, so it keeps its own palette instead.
    let colors: ColorsMap = (1..=5)
        .map(|index| (format!("accent{index}"), Color::new(1, 2, 3)))
        .collect();

    assert!(theme_accent_palette(&colors).is_empty());
}

// ----- A chart's text face against the package theme (issue #668) -----

fn office_font_scheme() -> ThemeFontScheme {
    ThemeFontScheme {
        major_latin: Some("Calibri Light".to_string()),
        minor_latin: Some("Calibri".to_string()),
    }
}

#[test]
fn a_chart_naming_no_face_takes_the_theme_body_font() {
    // `bar-chart.pptx` is this shape: `c:txPr` carries only `sz`, so the face
    // has to resolve through `+mn-lt` to the theme's minor font.
    assert_eq!(
        office_font_scheme().resolve_chart_text_typeface(None),
        Some("Calibri".to_string())
    );
}

#[test]
fn a_chart_naming_a_theme_token_resolves_it() {
    let scheme = office_font_scheme();
    assert_eq!(
        scheme.resolve_chart_text_typeface(Some("+mn-lt")),
        Some("Calibri".to_string())
    );
    assert_eq!(
        scheme.resolve_chart_text_typeface(Some("+mj-lt")),
        Some("Calibri Light".to_string())
    );
}

#[test]
fn a_chart_naming_a_real_face_keeps_it() {
    // `office2pdf_introduction_ko.pptx`'s chart1.xml names Calibri outright.
    assert_eq!(
        office_font_scheme().resolve_chart_text_typeface(Some("Impact")),
        Some("Impact".to_string())
    );
}

#[test]
fn a_theme_with_no_body_font_leaves_the_face_unset() {
    // Better to fall through to the renderer's default than to name a face
    // spelled `+mn-lt`, which would select nothing.
    assert_eq!(
        ThemeFontScheme::default().resolve_chart_text_typeface(None),
        None
    );
    assert_eq!(
        ThemeFontScheme::default().resolve_chart_text_typeface(Some("+mn-lt")),
        None
    );
}

#[test]
fn a_declared_family_class_is_read_from_pitch_family_and_panose() {
    // `Avenir Next LT Pro` carries no `sans` token in its name, so only the
    // declaration tells us it is one (issue #891).
    assert_eq!(
        declared_font_class(Some("020B0502020202020204"), Some("34")),
        Some(DeclaredFontClass::SansSerif)
    );
    // pitchFamily wins where both are present: its high nibble is the family.
    assert_eq!(
        declared_font_class(Some("02020603050405020304"), Some("18")),
        Some(DeclaredFontClass::Serif)
    );
    // PANOSE alone still answers when pitchFamily is absent or unusable.
    assert_eq!(
        declared_font_class(Some("020B0604020202020204"), None),
        Some(DeclaredFontClass::SansSerif)
    );
    assert_eq!(
        declared_font_class(Some("02040503050406030204"), None),
        Some(DeclaredFontClass::Serif)
    );
    // A non-latin PANOSE family (first byte != 2) classifies nothing.
    assert_eq!(
        declared_font_class(Some("05000000000000000000"), None),
        None
    );
    assert_eq!(declared_font_class(None, None), None);
}

#[test]
fn the_font_class_sweep_skips_theme_slots() {
    // `+mj-lt` names a theme slot rather than a face, so recording a class
    // under that spelling would attach it to every theme-following run.
    let xml = r#"<a:t xmlns:a="x">
        <a:latin typeface="+mj-lt" panose="020B0502" pitchFamily="34"/>
        <a:latin typeface="Posterama" panose="020B0502" pitchFamily="34"/>
        <a:latin typeface="Posterama" panose="02020603" pitchFamily="18"/>
    </a:t>"#;
    let mut out: HashMap<String, DeclaredFontClass> = HashMap::new();
    scan_declared_font_classes(xml, &mut out);

    // Keyed by the normalized spelling, which is what the substitution chain
    // looks the family up with.
    assert_eq!(out.get("posterama"), Some(&DeclaredFontClass::SansSerif));
    assert!(!out.contains_key("+mj-lt"));
    // First declaration wins, so a later contradictory one cannot flip it.
    assert_eq!(out.len(), 1);
}
