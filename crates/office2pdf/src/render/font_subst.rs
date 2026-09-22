//! Font substitution and fallback chains.
//!
//! Explicit family mappings provide preferred alternatives, including
//! metric-compatible Office substitutes. A name the table does not list is
//! answered in two further steps: the class the document itself declared for
//! that family, then — failing that — the name itself, read either as a class
//! token or, for a short list of brands that carry no such token, as a
//! known-brand match on its first word. A monospace answer gets a fixed-pitch
//! chain and a sans-serif one a sans chain, so a missing face does not land on
//! the document's default serif. The requested family remains first.
//!
//! Every name in the table is itself a font a host may not have, so each
//! listed family also states its own class, and the list Typst paints from
//! ends on that class's generic faces. Without that a Calibri run on a machine
//! carrying neither Carlito nor Liberation Sans still landed on the default
//! serif. A metrics lookup gets no such tail — see [`ChainPurpose`] (issue
//! #1213).
//!
//! Only PPTX populates the declared-class map today; DOCX `w:family` in
//! `word/fontTable.xml` is not read yet, so a DOCX face still relies on the
//! table and the name (issue #891).

// The substitution index also serves document- and caller-provided in-memory
// fonts. Filesystem-facing fallback paths remain compiled on WASM but unused.
#![cfg_attr(target_arch = "wasm32", allow(dead_code))]

use std::cell::RefCell;
use std::collections::{BTreeSet, HashMap};

#[cfg(target_arch = "wasm32")]
use std::path::PathBuf;
use typst::text::FontWeight;

use crate::ir::{
    Block, Document, FixedElementKind, HFInline, HeaderFooter, Page, Paragraph, Table,
};

use super::font_context::FontSearchContext;
use super::typst_gen::escape_typst_string;
use crate::ir::DeclaredFontClass;

const MONOSPACE_SUBSTITUTES: &[&str] = &[
    "DejaVu Sans Mono",
    "Noto Sans Mono",
    "Liberation Mono",
    "Cousine",
];

/// Where a missing sans-serif family lands. Liberation Sans leads because it
/// is metric-compatible with Arial, which most Office sans faces are sized
/// against (issue #848).
const SANS_SERIF_SUBSTITUTES: &[&str] = &["Liberation Sans", "Arimo", "DejaVu Sans", "Helvetica"];

/// Where a family that is known to be serif lands when it is missing. A
/// family with no class signal at all still falls through to the document's
/// default face, which is itself a serif (issue #891).
const SERIF_SUBSTITUTES: &[&str] = &["Liberation Serif", "Tinos", "DejaVu Serif"];

/// The class a family belongs to, and so the generic chain its own chain ends
/// on.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FamilyClass {
    SansSerif,
    Serif,
    Monospace,
}

impl FamilyClass {
    /// The faces that keep the class when nothing closer to the family is
    /// installed.
    fn substitutes(self) -> &'static [&'static str] {
        match self {
            Self::SansSerif => SANS_SERIF_SUBSTITUTES,
            Self::Serif => SERIF_SUBSTITUTES,
            Self::Monospace => MONOSPACE_SUBSTITUTES,
        }
    }
}

impl From<DeclaredFontClass> for FamilyClass {
    fn from(class: DeclaredFontClass) -> Self {
        match class {
            DeclaredFontClass::SansSerif => Self::SansSerif,
            DeclaredFontClass::Monospace => Self::Monospace,
            DeclaredFontClass::Serif => Self::Serif,
        }
    }
}

thread_local! {
    static ACTIVE_FONT_CONTEXT: RefCell<Option<FontSearchContext>> = const { RefCell::new(None) };
    /// The family classes the document itself declares, keyed by lowercased
    /// family name. A face that states its class outranks any guess drawn
    /// from its name (issue #891).
    static DECLARED_FONT_CLASSES: RefCell<HashMap<String, DeclaredFontClass>> =
        RefCell::new(HashMap::new());
}

/// Install the family classes a document declares, for the duration of one
/// render. Returns the previous map so a nested render can restore it.
pub(crate) fn set_declared_font_classes(
    classes: HashMap<String, DeclaredFontClass>,
) -> HashMap<String, DeclaredFontClass> {
    DECLARED_FONT_CLASSES.with(|cell| cell.replace(classes))
}

/// The class the document itself declared for a family, if it declared one.
fn declared_class(normalized_family: &str) -> Option<FamilyClass> {
    DECLARED_FONT_CLASSES.with(|cell| {
        cell.borrow()
            .get(normalized_family)
            .copied()
            .map(FamilyClass::from)
    })
}

fn normalized_lookup_key(font_family: &str) -> String {
    let trimmed = font_family.trim();
    let lower = trimmed.to_ascii_lowercase();
    if lower.starts_with("pretendard") {
        return "pretendard".to_string();
    }
    // Map Korean font names to their English equivalents for lookup.
    // OOXML files often use localized names (e.g., "맑은 고딕" instead of
    // "Malgun Gothic") which Typst doesn't recognise by default.
    match trimmed {
        "맑은 고딕" => "malgun gothic".to_string(),
        "굴림" => "gulim".to_string(),
        "돋움" => "dotum".to_string(),
        "바탕" => "batang".to_string(),
        "궁서" => "gungsuh".to_string(),
        "나눔고딕" | "나눔 고딕" => "nanum gothic".to_string(),
        "나눔명조" | "나눔 명조" => "nanum myeongjo".to_string(),
        "MS 고딕" => "ms gothic".to_string(),
        "MS 명조" => "ms mincho".to_string(),
        "メイリオ" => "meiryo".to_string(),
        "MS ゴシック" => "ms gothic".to_string(),
        "MS 明朝" => "ms mincho".to_string(),
        "游ゴシック" => "yu gothic".to_string(),
        "微软雅黑" => "microsoft yahei".to_string(),
        "宋体" | "宋體" => "simsun".to_string(),
        _ => lower,
    }
}

fn alias_family(font_family: &str) -> Option<&'static str> {
    let trimmed = font_family.trim();
    let lower = trimmed.to_ascii_lowercase();
    if lower.starts_with("pretendard") && lower != "pretendard" {
        return Some("Pretendard");
    }
    // Map localized CJK font names to their English names so Typst can
    // find them when the system registers fonts under English names.
    match trimmed {
        "맑은 고딕" => Some("Malgun Gothic"),
        "굴림" => Some("Gulim"),
        "돋움" => Some("Dotum"),
        "바탕" => Some("Batang"),
        "궁서" => Some("Gungsuh"),
        "나눔고딕" | "나눔 고딕" => Some("Nanum Gothic"),
        "나눔명조" | "나눔 명조" => Some("Nanum Myeongjo"),
        "MS 고딕" => Some("MS Gothic"),
        "MS 명조" => Some("MS Mincho"),
        "メイリオ" => Some("Meiryo"),
        "MS ゴシック" => Some("MS Gothic"),
        "MS 明朝" => Some("MS Mincho"),
        "游ゴシック" => Some("Yu Gothic"),
        "微软雅黑" => Some("Microsoft YaHei"),
        "宋体" | "宋體" => Some("SimSun"),
        _ => None,
    }
}

/// The family name Typst indexes a face under: the name-table family with its
/// style suffixes trimmed.
///
/// A port of typst-library's private `typographic_family`, kept byte-for-byte
/// in its suffix, modifier and separator lists. Its job is to trim a *request*
/// the way the book trimmed the *face*: `calibril.ttf` declares `Calibri
/// Light` and is registered as `Calibri` at weight 300, so a lookup keyed on
/// `Calibri Light` finds nothing even where Office ships the face. A trailing
/// suffix only counts when a separator precedes it, so `CalibriLight` and
/// `Blackadder ITC` stay whole (issue #1286).
fn typographic_family(family: &str) -> &str {
    const SEPARATORS: [char; 3] = [' ', '-', '_'];
    const MODIFIERS: &[&str] = &[
        "extra", "ext", "ex", "x", "semi", "sem", "sm", "demi", "dem", "ultra",
    ];
    #[rustfmt::skip]
    const SUFFIXES: &[&str] = &[
        "normal", "italic", "oblique", "slanted",
        "thin", "th", "hairline", "light", "lt", "regular", "medium", "med",
        "md", "bold", "bd", "demi", "extb", "black", "blk", "bk", "heavy",
        "narrow", "condensed", "cond", "cn", "cd", "compressed", "expanded", "exp",
        "vf", "var", "variable",
    ];

    let family: &str = family.trim().trim_start_matches('.');
    let lower: String = family.to_ascii_lowercase();
    let mut len: usize = usize::MAX;
    let mut trimmed: &str = lower.as_str();

    while trimmed.len() < len {
        len = trimmed.len();

        let mut tail: &str = trimmed;
        let mut shortened: bool = false;
        while let Some(rest) = SUFFIXES.iter().find_map(|suffix| tail.strip_suffix(suffix)) {
            shortened = true;
            tail = rest;
        }
        if !shortened {
            break;
        }

        if let Some(rest) = tail.strip_suffix(SEPARATORS) {
            trimmed = rest;
            tail = rest;
        }

        if let Some(rest) = MODIFIERS
            .iter()
            .find_map(|modifier| tail.strip_suffix(modifier))
            && let Some(stripped) = rest.strip_suffix(SEPARATORS)
        {
            trimmed = stripped;
        }
    }

    &family[..len]
}

/// The weight a family name states in its style suffix, as the member of the
/// base family it denotes: `Calibri Light` is `Calibri` at 300, `Segoe UI
/// Semibold` is `Segoe UI` at 600, `Arial Black` is `Arial` at 900.
///
/// Only the suffix counts, and only a weight suffix: a stretch such as `Arial
/// Narrow` states no weight, and a weight word inside the name — `Blackadder
/// ITC`, `Lightning Sans` — names a family, not a member of one.
///
/// TODO(no Typst name for 350): `Semilight` and `Demilight` are trimmed by the
/// book but have no named weight the emitter can state, so they answer `None`
/// and keep the pre-#1286 behaviour.
pub(crate) fn weight_stated_by_family_name(font_family: &str) -> Option<FontWeight> {
    let requested: &str = font_family.trim();
    let base_family: &str = typographic_family(requested);
    if base_family.len() >= requested.len() {
        return None;
    }
    let suffix: String = requested[base_family.len()..]
        .chars()
        .filter(|character| !matches!(character, ' ' | '-' | '_'))
        .map(|character| character.to_ascii_lowercase())
        .collect();
    Some(match suffix.as_str() {
        "thin" | "hairline" => FontWeight::THIN,
        "extralight" | "ultralight" => FontWeight::EXTRALIGHT,
        "light" => FontWeight::LIGHT,
        "medium" => FontWeight::MEDIUM,
        "semibold" | "demibold" | "demi" => FontWeight::SEMIBOLD,
        "bold" => FontWeight::BOLD,
        "extrabold" | "ultrabold" | "extb" => FontWeight::EXTRABOLD,
        "black" | "heavy" => FontWeight::BLACK,
        _ => return None,
    })
}

/// The base family a weight-suffixed request has to reach for its member:
/// `Some("Calibri")` for `Calibri Light`, `None` for a name stating no
/// weight.
/// The family a variable-font request is filed under: typst 0.15 trims
/// `Variable`, `Var` and `VF` from a face's family, so `Pretendard Variable` is
/// indexed as `Pretendard`. Unlike a weight suffix, the suffix names the face
/// the request means, not another member of its family.
fn variable_font_base_family(font_family: &str) -> Option<&str> {
    let requested: &str = font_family.trim();
    let base_family: &str = typographic_family(requested);
    if base_family.len() >= requested.len() {
        return None;
    }
    let suffix: String = requested[base_family.len()..]
        .chars()
        .filter(|character| !matches!(character, ' ' | '-' | '_'))
        .map(|character| character.to_ascii_lowercase())
        .collect();
    matches!(suffix.as_str(), "variable" | "var" | "vf").then_some(base_family)
}

fn weight_member_base_family(font_family: &str) -> Option<&str> {
    weight_stated_by_family_name(font_family)?;
    Some(typographic_family(font_family.trim()))
}

/// What a candidate list is being built for.
///
/// The two answers differ only in the generic class tail. Painting wants it:
/// Typst walks the list per glyph, so a face of the family's own class beats
/// falling through to the engine's default serif. Reading a family's *metrics*
/// does not — [`family_candidates`] takes the numbers of the first candidate
/// that resolves, whole, and a generic class face's numbers are not the
/// family's. A fit-to-page sheet whose Normal font is `Trebuchet MS` re-scaled
/// by 13% on a host without Ubuntu once the tail handed its column unit
/// Liberation Sans's digit advance, and a Korean footer on a host without a
/// Korean font reported a Latin line box (issue #1213).
///
/// A weight-suffixed request's base family is *not* one of those differences:
/// both chains carry it, because the resolver walks them at the weight the
/// name states and so lands on the same member (issue #1643).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ChainPurpose {
    /// Choosing the faces Typst paints with.
    Paint,
    /// Resolving one face to read its metrics from.
    Metrics,
}

fn fallback_candidates(
    font_family: &str,
    context: Option<&FontSearchContext>,
    purpose: ChainPurpose,
) -> Vec<String> {
    let mut candidates: Vec<String> = Vec::new();
    let requested = font_family.trim();

    // Typst never finds a family called `Calibri Light`: the book holds that
    // face as `Calibri` at 300. Both chains follow the base family so the
    // weight the run states lands on the member it names — the resolver walks
    // this list at that weight, so a metrics lookup reads the same face the
    // paint chain shapes with rather than the family's regular member (issues
    // #1286 and #1643). A variable-font suffix, which typst 0.15 trims too,
    // names the very face requested, so both follow it as well.
    let base_family: Option<&str> =
        variable_font_base_family(requested).or_else(|| weight_member_base_family(requested));
    if let Some(base_family) = base_family {
        candidates.push(base_family.to_string());
    }

    if let Some(alias) = alias_family(requested)
        && !alias.eq_ignore_ascii_case(requested)
        && !candidates
            .iter()
            .any(|candidate| candidate.eq_ignore_ascii_case(alias))
    {
        candidates.push(alias.to_string());
    }

    // Rank the family's own substitutes by source, then rank the generic class
    // tail separately. The tail is a last resort and must never jump ahead of
    // a metric-compatible stand-in merely because its file came from a
    // higher-priority search path (issue #1213).
    let normalized_family: String = normalized_lookup_key(requested);
    let mut append_ranked_group = |group: &'static [&'static str]| {
        let mut ordered: Vec<&'static str> = Vec::new();
        for sub in group.iter().copied() {
            let already_named: bool = sub.eq_ignore_ascii_case(requested)
                || candidates
                    .iter()
                    .any(|candidate| candidate.eq_ignore_ascii_case(sub))
                || ordered.iter().any(|kept| kept.eq_ignore_ascii_case(sub));
            if !already_named {
                ordered.push(sub);
            }
        }

        let mut ranked: Vec<(u8, usize, &'static str)> = ordered
            .iter()
            .enumerate()
            .map(|(index, sub)| {
                let rank = context.map(|ctx| ctx.family_source_rank(sub)).unwrap_or(2);
                (rank, index, *sub)
            })
            .collect();
        ranked.sort_by_key(|(rank, index, _)| (*rank, *index));
        candidates.extend(ranked.into_iter().map(|(_, _, sub)| sub.to_string()));
    };

    append_ranked_group(substitutes(requested).unwrap_or(&[]));
    if purpose == ChainPurpose::Paint {
        append_ranked_group(class_tail(&normalized_family));
    }

    candidates
}

/// The substitution table's entry for a family it lists: the class the family
/// belongs to, and the faces that stand in for it, in preference order.
///
/// The class travels with the faces because a chain of real fonts can run out.
/// Every name here is a font a host may simply not have, and once the last one
/// is missing the family has nothing left — see [`class_tail`] (issue #1213).
fn table_entry(normalized_family: &str) -> Option<(FamilyClass, &'static [&'static str])> {
    use FamilyClass::{Monospace, SansSerif, Serif};
    Some(match normalized_family {
        "calibri" => (SansSerif, &["Carlito", "Liberation Sans"]),
        // LibreOffice 25.2 resolves the unembedded Aptos footer in the Office
        // poster fixture to Noto Sans 2.015. The paint-chain builder also
        // suppresses an incidental host Aptos while preserving a caller- or
        // package-provided face (issue #1463).
        "aptos" => (SansSerif, &["Noto Sans"]),
        // The Gift Budget workbook from issue #982 declares Segoe UI but does
        // not embed it. A host without that proprietary face used Typst's
        // Libertinus Serif default, changing the native instruction panel's
        // wrap boundaries. Reuse the bundled, reproducible sans face instead
        // of making this result depend on the host font book (issue #1472).
        "segoe ui" => (SansSerif, &["Selawik"]),
        // `Calibri Light` is the `majorHAnsi` face of every Office theme since
        // 2013, so it is the face every built-in `Heading N` resolves to. It
        // needs its own entry even where Office ships it: Typst keys its font
        // book on the family name with style suffixes trimmed, so `calibril.ttf`
        // is indexed under `Calibri` at weight 300 and the stated name matches
        // nothing. Without an entry the run got no chain and, through
        // `best_face`, no metrics either — so `word_line_height_settings` left
        // a theme-headed paragraph on Typst's glyph-tight line box and the gap
        // below the heading came out a descender short (issue #1197).
        //
        // `Calibri` leads because the light member declares Calibri's own hhea
        // line — both are 1950/-550/0 on 2048 upem — so it reproduces the line
        // box either way. Since #1643 the chain builder adds `Calibri` to both
        // chains for every weight-suffixed name, and the resolver picks the 300
        // member off it; this entry stays because it also orders the substitutes
        // a host without Calibri falls back on.
        "calibri light" => (SansSerif, &["Calibri", "Carlito", "Liberation Sans"]),
        "carlito" => (SansSerif, &["Calibri", "Liberation Sans", "Arimo", "Arial"]),
        "cambria" => (Serif, &["Caladea", "Liberation Serif"]),
        "arial" => (SansSerif, &["Liberation Sans", "Arimo"]),
        "times new roman" => (Serif, &["Liberation Serif", "Tinos"]),
        "courier new" => (Monospace, &["Liberation Mono", "Cousine"]),
        "comic sans ms" => (SansSerif, &["Comic Neue"]),
        "verdana" => (SansSerif, &["DejaVu Sans"]),
        "georgia" => (Serif, &["DejaVu Serif"]),
        "consolas" => (Monospace, &["Inconsolata"]),
        "trebuchet ms" => (SansSerif, &["Ubuntu"]),
        "impact" => (SansSerif, &["Oswald"]),
        // LibreOffice 25.2 resolves the unembedded display faces in the
        // Office poster template from issue #1458 to Noto Serif 2.015. Keep
        // those declarations on the same reproducible OFL face instead of
        // letting Typst fall through to its unrelated default serif.
        "avenir next lt pro" | "avenir next w1g medium" | "the hand black" => {
            (Serif, &["Noto Serif"])
        }
        "raleway" => (
            SansSerif,
            &[
                "Helvetica",
                "Arial",
                "Arial Unicode MS",
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Malgun Gothic",
                "Liberation Sans",
            ],
        ),
        "lato" => (
            SansSerif,
            &[
                "Helvetica",
                "Arial",
                "Arial Unicode MS",
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Malgun Gothic",
                "Liberation Sans",
            ],
        ),
        "pretendard" => (
            SansSerif,
            &[
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Malgun Gothic",
                "Arial Unicode MS",
                "Helvetica",
                "Arial",
                "Liberation Sans",
            ],
        ),
        // Korean font names → English equivalents + fallbacks
        "malgun gothic" => (
            SansSerif,
            &[
                "Malgun Gothic",
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Arial Unicode MS",
            ],
        ),
        "gulim" => (
            SansSerif,
            &[
                "Gulim",
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Malgun Gothic",
                "Arial Unicode MS",
            ],
        ),
        "dotum" => (
            SansSerif,
            &[
                "Dotum",
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Malgun Gothic",
                "Arial Unicode MS",
            ],
        ),
        "batang" => (
            Serif,
            &[
                "Batang",
                "Noto Serif CJK KR",
                "Apple Myungjo",
                "Arial Unicode MS",
            ],
        ),
        "gungsuh" => (
            Serif,
            &[
                "Gungsuh",
                "Noto Serif CJK KR",
                "Apple Myungjo",
                "Arial Unicode MS",
            ],
        ),
        "nanum gothic" => (
            SansSerif,
            &[
                "Nanum Gothic",
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Malgun Gothic",
                "Arial Unicode MS",
            ],
        ),
        "nanum myeongjo" => (
            Serif,
            &[
                "Nanum Myeongjo",
                "Noto Serif CJK KR",
                "Apple Myungjo",
                "Batang",
                "Arial Unicode MS",
            ],
        ),
        // Japanese font names → English equivalents + fallbacks
        "ms gothic" => (
            SansSerif,
            &["MS Gothic", "Noto Sans CJK JP", "Hiragino Sans"],
        ),
        "ms mincho" => (
            Serif,
            &["MS Mincho", "Noto Serif CJK JP", "Hiragino Mincho ProN"],
        ),
        "meiryo" => (SansSerif, &["Meiryo", "Noto Sans CJK JP", "Hiragino Sans"]),
        "yu gothic" => (
            SansSerif,
            &["Yu Gothic", "Noto Sans CJK JP", "Hiragino Sans"],
        ),
        // Chinese font names → English equivalents + fallbacks
        "microsoft yahei" => (
            SansSerif,
            &[
                "Microsoft YaHei",
                "Noto Sans CJK SC",
                "PingFang SC",
                "Arial Unicode MS",
            ],
        ),
        "simsun" => (
            Serif,
            &["SimSun", "Noto Serif CJK SC", "STSong", "Arial Unicode MS"],
        ),
        // Noto CJK families are common in documents authored on Linux or with
        // Google Fonts, but are rarely installed on macOS/Windows. Without a
        // chain the renderer emits a one-element font stack and Typst's own
        // fallback picks a regular-only face, silently dropping bold/italic.
        // Short names ("Noto Sans KR") are the Google Fonts per-language
        // builds of the same designs.
        "noto sans cjk kr" | "noto sans kr" => (
            SansSerif,
            &[
                "Noto Sans CJK KR",
                "Noto Sans KR",
                "Apple SD Gothic Neo",
                "Malgun Gothic",
                "Arial Unicode MS",
            ],
        ),
        "noto sans cjk sc" | "noto sans sc" => (
            SansSerif,
            &[
                "Noto Sans CJK SC",
                "Noto Sans SC",
                "PingFang SC",
                "Microsoft YaHei",
                "Apple SD Gothic Neo",
                "Arial Unicode MS",
            ],
        ),
        "noto sans cjk tc" | "noto sans tc" => (
            SansSerif,
            &[
                "Noto Sans CJK TC",
                "Noto Sans TC",
                "PingFang TC",
                "Microsoft JhengHei",
                "Arial Unicode MS",
            ],
        ),
        "noto sans cjk jp" | "noto sans jp" => (
            SansSerif,
            &[
                "Noto Sans CJK JP",
                "Noto Sans JP",
                "Hiragino Sans",
                "Yu Gothic",
                "Meiryo",
                "Arial Unicode MS",
            ],
        ),
        "noto serif cjk kr" | "noto serif kr" => (
            Serif,
            &[
                "Noto Serif CJK KR",
                "Noto Serif KR",
                "Apple Myungjo",
                "Batang",
                "Arial Unicode MS",
            ],
        ),
        "noto serif cjk sc" | "noto serif sc" => (
            Serif,
            &[
                "Noto Serif CJK SC",
                "Noto Serif SC",
                "STSong",
                "SimSun",
                "Arial Unicode MS",
            ],
        ),
        "noto serif cjk tc" | "noto serif tc" => (
            Serif,
            &["Noto Serif CJK TC", "Noto Serif TC", "Arial Unicode MS"],
        ),
        "noto serif cjk jp" | "noto serif jp" => (
            Serif,
            &[
                "Noto Serif CJK JP",
                "Noto Serif JP",
                "Hiragino Mincho ProN",
                "Yu Mincho",
                "Arial Unicode MS",
            ],
        ),
        "corbel" | "candara" => (SansSerif, SANS_SERIF_SUBSTITUTES),
        _ => return None,
    })
}

/// The class a family the table does not list states for itself: the class the
/// document declared for it, then a class token in its own name, then a
/// known-brand match on its first word.
///
/// The document's own declaration outranks anything read off the name (issue
/// #891).
fn inferred_class(normalized_family: &str) -> Option<FamilyClass> {
    declared_class(normalized_family)
        // Monospace first: a name carrying both tokens, as "… Sans Mono"
        // does, is fixed-pitch.
        .or_else(|| {
            family_name_declares_monospace(normalized_family).then_some(FamilyClass::Monospace)
        })
        .or_else(|| {
            (family_name_declares_sans_serif(normalized_family)
                || family_name_is_known_sans_serif_brand(normalized_family))
            .then_some(FamilyClass::SansSerif)
        })
}

/// The generic chain a family's own substitutes end on when Typst paints with
/// them, so a family that exhausts them lands on its own class instead of on
/// the engine's default face — a serif, whatever the family was (issue #1213).
///
/// Only a family the table lists has a tail to add. One the table does not
/// list is already answered with its class chain and nothing else, so there is
/// nothing left to append. See [`ChainPurpose`] for why a metrics lookup gets
/// no tail at all.
fn class_tail(normalized_family: &str) -> &'static [&'static str] {
    match table_entry(normalized_family) {
        Some((class, _)) => class.substitutes(),
        None => &[],
    }
}

/// Return metric- or family-class-compatible substitutes for a font family.
///
/// Returns `None` if no substitution is defined and the name provides no
/// reliable family-class signal.
///
/// The returned slice is ordered by preference. Explicit mappings preserve the
/// known family's intent; class-derived mappings preserve the class the
/// document declared for the family, or failing that the one its own name
/// declares — fixed pitch, or sans-serif. The generic tail a listed family
/// falls back on once these are exhausted is [`class_tail`], which the chain
/// builders append; it stands in for the class rather than for the family, so
/// it is not reported here.
pub fn substitutes(font_family: &str) -> Option<&'static [&'static str]> {
    let normalized_family = normalized_lookup_key(font_family);
    if let Some((_, families)) = table_entry(&normalized_family) {
        return Some(families);
    }
    inferred_class(&normalized_family).map(FamilyClass::substitutes)
}

/// OOXML may provide no usable family class beyond the requested name. A
/// standalone class token still lets a missing fixed-pitch face avoid Typst's
/// proportional default without mistaking brand names such as Monotype.
fn family_name_declares_monospace(normalized_family: &str) -> bool {
    normalized_family
        .split(|character: char| !character.is_ascii_alphanumeric())
        .any(|token| matches!(token, "mono" | "monospace" | "typewriter"))
}

/// A standalone class token marking the requested family as sans-serif.
///
/// Without this an unlisted sans family fell through to the document's default
/// face, which is a serif — a family-class error rather than the metric
/// difference a substitution normally costs (issue #848). Matching whole
/// tokens keeps brand names out of it, the same discipline
/// [`family_name_declares_monospace`] uses: `Sansation` and `Gothamist` are
/// not classified, `Microsoft Sans Serif` and `Franklin Gothic Demi` are.
///
/// `sans` is tested for its own sake, so `Microsoft Sans Serif` resolves as
/// sans despite also carrying `serif`.
fn family_name_declares_sans_serif(normalized_family: &str) -> bool {
    normalized_family
        .split(|character: char| !character.is_ascii_alphanumeric())
        .any(|token| matches!(token, "sans" | "gothic" | "grotesk" | "grotesque"))
}

/// A sans-serif family that says so nowhere a heuristic can read it: no class
/// token in its name, and no declared class either where the source names a
/// face as a bare string — an XLSX header/footer's `&"Font"` code carries
/// nothing else at all.
///
/// Matched on the family's *first* token, so every weight and width of the
/// family lands with it without enumerating them. The monospace check runs
/// first, so a fixed-pitch member such as `Aptos Mono` still resolves as
/// monospace.
///
/// `Aptos` has been Microsoft 365's default face since 2024, which puts it in
/// every document a current Office build creates; without this it fell through
/// to the document's serif default (issue #949).
fn family_name_is_known_sans_serif_brand(normalized_family: &str) -> bool {
    let first_token: Option<&str> = normalized_family
        .split(|character: char| !character.is_ascii_alphanumeric())
        .find(|token| !token.is_empty());
    matches!(first_token, Some("aptos"))
}

/// Whether the declared family itself stays at the front of a conversion's
/// paint and metric chains.
///
/// Aptos is special only when it is unembedded: its incidental presence in a
/// host Office installation made the #1407 DOCX footer narrower on that host
/// than on CI and in LibreOffice. An embedded face, a caller font path, or
/// registered bytes remain explicit conversion inputs and still lead. With no
/// active context, preserve the historical declaration-first behaviour.
fn requested_family_leads(font_family: &str, context: Option<&FontSearchContext>) -> bool {
    normalized_lookup_key(font_family) != "aptos"
        || context.is_none_or(|context| context.is_user_family(font_family))
}

/// Check whether the given font family (or its alias) is available in the
/// current font context. Returns `true` when no context is active to preserve
/// existing behaviour.
///
/// A name stating a weight member — `Calibri Light`, `Segoe UI Semibold` — is
/// available when the base family the book indexes it under ships a face at
/// that weight. The book never lists the suffixed name itself, so without this
/// the check answered `false` for the theme heading face of every Word
/// package on a host where Office ships it, and the inferred weight was
/// dropped for Typst's `#heading` bold (issue #1286).
pub fn is_primary_font_available(font_family: &str) -> bool {
    ACTIVE_FONT_CONTEXT.with(|cell| {
        let guard = cell.borrow();
        let Some(ctx) = guard.as_ref() else {
            return true;
        };
        if !requested_family_leads(font_family, Some(ctx)) {
            return false;
        }
        if ctx.has_family(font_family) {
            return true;
        }
        if let Some(alias) = alias_family(font_family)
            && ctx.has_family(alias)
        {
            return true;
        }
        if let Some(weight) = weight_stated_by_family_name(font_family) {
            let base_family: &str = typographic_family(font_family.trim());
            return ctx.has_face_weight(base_family, weight);
        }
        if let Some(base_family) = variable_font_base_family(font_family) {
            return ctx.has_family(base_family);
        }
        false
    })
}

/// The East Asian script a run's text is written in, as font selection sees it.
///
/// A run states a family; the text states a script, and the two need not agree.
/// A PowerPoint run can declare `<a:ea typeface="Calibri"/>` over Hangul, and a
/// workbook can declare a Simplified Chinese family over Korean text. In both
/// cases the declared family has no glyph for what the run actually holds, and
/// the chain it carries has none either, so the text landed on whatever the
/// font book happened to offer first — Gulim, which has no bold member at all
/// (#543), or a Chinese face 13.5% narrower than the one Excel picks (#537).
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum TextScript {
    /// No East Asian character appears, so no script chain applies.
    Latin,
    Korean,
    Japanese,
    Chinese,
}

impl TextScript {
    /// Faces that cover the script, in the order Word and PowerPoint reach for
    /// them.
    ///
    /// Appending these lets the per-glyph fallback answer the question the text
    /// asks rather than the one its family name asks.
    fn fallbacks(self, serif: bool) -> &'static [&'static str] {
        match (self, serif) {
            (Self::Latin, _) => &[],
            // PowerPoint substitutes a serif CJK face for a serif request, to
            // keep the typographic voice. Reaching the sans chain regardless
            // put 29 slide titles declaring a serif `a:ea` into a geometric
            // sans at a much heavier weight (issue #687).
            (Self::Korean, true) => &[
                "Batang",
                "Noto Serif CJK KR",
                "Apple Myungjo",
                "Gungsuh",
                "Arial Unicode MS",
            ],
            (Self::Korean, false) => &[
                "Malgun Gothic",
                "Apple SD Gothic Neo",
                "Noto Sans CJK KR",
                "Arial Unicode MS",
            ],
            (Self::Japanese, true) => &[
                "MS Mincho",
                "Noto Serif CJK JP",
                "Hiragino Mincho ProN",
                "YuMincho",
            ],
            (Self::Japanese, false) => &[
                "Yu Gothic",
                "Hiragino Sans",
                "Noto Sans CJK JP",
                "MS Gothic",
            ],
            (Self::Chinese, true) => &["SimSun", "Noto Serif CJK SC", "Songti SC", "STSong"],
            (Self::Chinese, false) => &[
                "Microsoft YaHei",
                "PingFang SC",
                "Noto Sans CJK SC",
                "SimSun",
            ],
        }
    }
}

/// Whether `font_family` names an East Asian face.
///
/// Keyed on the family, not on any text: Word's East Asian line metrics follow
/// the face a line is set in, and a CJK family shapes its own Latin glyphs.
/// A run naming `w:eastAsia="Arial"` — which the Latin business fixtures do —
/// is not East Asian and must not be caught here.
pub(crate) fn is_east_asian_family(font_family: &str) -> bool {
    matches!(
        normalized_lookup_key(font_family).as_str(),
        "malgun gothic"
            | "gulim"
            | "dotum"
            | "batang"
            | "gungsuh"
            | "nanum gothic"
            | "nanum myeongjo"
            | "ms gothic"
            | "ms mincho"
            | "meiryo"
            | "yu gothic"
            | "microsoft yahei"
            | "simsun"
    )
}

/// Classify `text` by the first script-specific character it carries.
pub(crate) fn text_script(text: &str) -> TextScript {
    let mut has_han = false;
    for character in text.chars() {
        match character as u32 {
            // Hangul syllables, Jamo, and compatibility Jamo.
            0xAC00..=0xD7AF | 0x1100..=0x11FF | 0x3130..=0x318F => return TextScript::Korean,
            // Hiragana and Katakana.
            0x3040..=0x30FF => return TextScript::Japanese,
            // Han, which Korean and Japanese also use — only decisive when no
            // script-specific character appears anywhere in the run.
            0x4E00..=0x9FFF | 0x3400..=0x4DBF => has_han = true,
            _ => {}
        }
    }
    if has_han {
        TextScript::Chinese
    } else {
        TextScript::Latin
    }
}

/// Whether `font_family` names a serif face.
///
/// Keyed on the family the document *asks for*, because that is what decides
/// the voice the substitute has to keep. Recognises the Office serif families
/// by name and anything that says so in its own name, which covers the
/// metric-compatible substitutes (`Liberation Serif`, `Tinos`, `Caladea`) and
/// the CJK serif families whose names do not contain "serif".
fn family_is_serif(font_family: &str) -> bool {
    let normalized = normalized_lookup_key(font_family);
    if matches!(
        normalized.as_str(),
        "cambria"
            | "times new roman"
            | "georgia"
            | "garamond"
            | "book antiqua"
            | "palatino"
            | "palatino linotype"
            | "constantia"
            | "bookman old style"
            | "century schoolbook"
            | "caladea"
            | "tinos"
    ) {
        return true;
    }
    ["serif", "batang", "gungsuh", "myungjo", "myeongjo", "mincho", "songti", "simsun"]
        .iter()
        .any(|token| normalized.contains(token))
        // "Sans Serif" and "Noto Sans CJK" are not serif despite the substring.
        && !normalized.contains("sans")
}

fn script_fallbacks(font_family: &str, text: &str) -> &'static [&'static str] {
    text_script(text).fallbacks(family_is_serif(font_family))
}

/// Latin faces that carry the symbol blocks list markers are drawn from.
///
/// Word resolves a marker its declared family lacks to Arial, not to that
/// family's own substitute chain. Liberation Sans is metric-compatible and
/// carries U+25E6 at 0.3545em — the advance Word's ArialMT uses.
const SYMBOL_FALLBACKS: &[&str] = &["Arial", "Liberation Sans", "Arimo", "Helvetica"];

/// Whether `text` carries a symbol a CJK family is likely to be missing.
///
/// Restricted to the blocks list bullets actually come from. General
/// punctuation is deliberately excluded: an em dash or a bullet that the
/// declared family *does* carry must keep that family's glyph, and widening
/// this predicate would put a Latin face ahead of the substitute chain for
/// text that never needed one.
fn has_symbol_needing_latin_fallback(text: &str) -> bool {
    text.chars().any(|character| {
        matches!(character as u32,
            0x2190..=0x21FF   // Arrows
            | 0x25A0..=0x25FF // Geometric Shapes — U+25E6 WHITE BULLET
            | 0x2600..=0x26FF // Miscellaneous Symbols
            | 0x2700..=0x27BF // Dingbats
        )
    })
}

/// Latin faces to try for `text`, after the declared family and its script's
/// faces but ahead of the family's own substitute chain.
fn symbol_fallbacks(text: &str) -> &'static [&'static str] {
    if has_symbol_needing_latin_fallback(text) {
        SYMBOL_FALLBACKS
    } else {
        &[]
    }
}

/// The font list for `family` covering the script `text` is written in.
///
/// The declared family leads, so a run naming a face that does cover its own
/// text keeps it. The script's faces come next, ahead of the declared family's
/// substitute chain: those substitutes preserve the requested family's metrics
/// or class, which is the wrong priority for a glyph the family does not have.
/// Placing them first is what sent Korean text through a Chinese face that
/// happens to carry some Hangul (issue #537).
///
/// Symbol-block glyphs — arrows, geometric shapes, dingbats — then get a Latin
/// chain, still ahead of the family's substitutes, because Word resolves a
/// marker its family lacks to Arial rather than through that family's own
/// chain. See [`symbol_fallbacks`] (issue #642).
///
/// This ordering assumes `text` is one run, and so one script. See
/// [`font_for_mixed_script_text`] for the case where one face has to cover
/// several at once.
pub fn font_with_fallbacks_for_text(font_family: &str, text: &str) -> String {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        join_font_list(latin_family_chain(font_family, text, context.as_ref()))
    })
}

/// The ordered candidate list behind [`font_with_fallbacks_for_text`], before
/// it is joined into a Typst value. Shared with
/// [`needs_synthetic_oblique`], which has to reason about the same chain the
/// emitter writes out.
fn latin_family_chain(
    font_family: &str,
    text: &str,
    context: Option<&FontSearchContext>,
) -> Vec<String> {
    let mut families: Vec<String> = Vec::new();
    if requested_family_leads(font_family, context) {
        families.push(font_family.to_string());
    }
    families.extend(
        script_fallbacks(font_family, text)
            .iter()
            .map(|face| (*face).to_string()),
    );
    families.extend(
        symbol_fallbacks(text)
            .iter()
            .map(|face| (*face).to_string()),
    );
    families.extend(fallback_candidates(
        font_family,
        context,
        ChainPurpose::Paint,
    ));
    append_last_resort(&mut families, context);
    families
}

/// The font list for one face that has to cover text in several scripts at
/// once.
///
/// A chart sets a single face for every string it draws, so the sample handed
/// here mixes the title, the categories and the series names — and with them
/// their scripts. [`font_with_fallbacks_for_text`] puts the script's faces
/// ahead of the declared family's substitutes, which is right when the run's
/// text *is* that script: a substitute preserving Latin metrics is the wrong
/// answer for a Hangul glyph. Applied to a mixed sample it is the wrong answer
/// the other way round — the Hangul face covers Latin too, so it takes the
/// Latin glyphs before the chain ever reaches the stand-in for the face the
/// chart asked for, and a Korean chart rendered its `DOCX` label in Malgun
/// Gothic instead of Calibri's substitute.
///
/// So the declared family's own substitutes come first here, and the script
/// faces catch only what neither covers (issue #668).
pub(crate) fn font_for_mixed_script_text(font_family: &str, text: &str) -> String {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        let mut families: Vec<String> = Vec::new();
        if requested_family_leads(font_family, context.as_ref()) {
            families.push(font_family.to_string());
        }
        families.extend(fallback_candidates(
            font_family,
            context.as_ref(),
            ChainPurpose::Paint,
        ));
        families.extend(
            script_fallbacks(font_family, text)
                .iter()
                .map(|face| (*face).to_string()),
        );
        families.extend(
            symbol_fallbacks(text)
                .iter()
                .map(|face| (*face).to_string()),
        );
        append_last_resort(&mut families, context.as_ref());
        join_font_list(families)
    })
}

fn append_last_resort(families: &mut Vec<String>, context: Option<&FontSearchContext>) {
    let Some(family) = context.and_then(FontSearchContext::last_resort_font_family) else {
        return;
    };
    if !families
        .iter()
        .any(|candidate| candidate.eq_ignore_ascii_case(family))
    {
        families.push(family.to_string());
    }
}

/// Whether `character` is written in an East Asian script, so the font list's
/// East Asian entries — not its Latin ones — decide the face it lands on.
///
/// The ranges are [`text_script`]'s, read one character at a time: a run that
/// mixes scripts resolves per glyph in Typst, and the synthetic-oblique
/// decision has to follow it (issue #686).
pub(crate) fn is_east_asian_char(character: char) -> bool {
    matches!(character as u32,
        0xAC00..=0xD7AF | 0x1100..=0x11FF | 0x3130..=0x318F // Hangul
        | 0x3040..=0x30FF                                   // Kana
        | 0x4E00..=0x9FFF | 0x3400..=0x4DBF                 // Han
    )
}

/// Whether a run marked italic has to be slanted by hand for `script`.
///
/// Word and PowerPoint synthesise an oblique when the face they resolve ships
/// no italic member; Typst has no such fallback and renders the text upright,
/// dropping the emphasis silently (issue #686). This answers whether that is
/// about to happen, by naming the family that will actually shape the script
/// and asking the font context what faces it has.
///
/// `false` whenever the answer is not certain — no active font context, no
/// declared family, or a family chain the context has never seen. Guessing the
/// other way would slant text that a real italic face is about to handle.
pub(crate) fn needs_synthetic_oblique(
    latin_family: Option<&str>,
    east_asian_family: Option<&str>,
    text: &str,
    script: TextScript,
) -> bool {
    let Some(latin_family) = latin_family else {
        return false;
    };
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        let Some(context) = context.as_ref() else {
            return false;
        };
        let families: Vec<String> = match east_asian_family {
            Some(east_asian) if !east_asian.eq_ignore_ascii_case(latin_family) => {
                east_asian_family_chain(latin_family, east_asian, text, Some(context))
            }
            _ => latin_family_chain(latin_family, text, Some(context)),
        };
        families
            .iter()
            .find_map(|family| {
                covers_script_with_alias(context, family, script).then(|| {
                    !context.has_italic_face(family)
                        && !alias_family(family).is_some_and(|alias| context.has_italic_face(alias))
                })
            })
            .unwrap_or(false)
    })
}

/// [`FontSearchContext::covers_script`] through the alias table, so a family
/// the context knows only under its substitute's name still answers.
fn covers_script_with_alias(context: &FontSearchContext, family: &str, script: TextScript) -> bool {
    context.covers_script(family, script)
        || alias_family(family).is_some_and(|alias| context.covers_script(alias, script))
}

/// Render a candidate list as a Typst font value, dropping repeats.
///
/// A lone family stays a bare string rather than a one-element list: that is
/// what a document naming a face with nothing to fall back to should emit, and
/// it keeps the generated source readable.
fn join_font_list(families: Vec<String>) -> String {
    let mut kept: Vec<String> = Vec::with_capacity(families.len());
    for family in families {
        if !kept.iter().any(|seen| seen.eq_ignore_ascii_case(&family)) {
            kept.push(family);
        }
    }
    if let [only] = kept.as_slice() {
        return format!("\"{}\"", escape_typst_string(only));
    }
    let mut result = String::with_capacity(64);
    result.push('(');
    for (index, family) in kept.iter().enumerate() {
        if index > 0 {
            result.push_str(", ");
        }
        result.push('"');
        result.push_str(&escape_typst_string(family));
        result.push('"');
    }
    result.push(')');
    result
}

/// The families to try, in order, when resolving a declared family to a real
/// face: the family itself, then its alias and substitutes.
///
/// A metrics lookup needs this as much as rendering does. Word documents name
/// East Asian families in their own script — `맑은 고딕` rather than `Malgun
/// Gothic` — and a font book registers the English name, so selecting on the
/// declared name alone finds nothing and the caller silently loses the font's
/// line metrics (issue #575).
pub(crate) fn family_candidates(font_family: &str) -> Vec<String> {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        let mut candidates: Vec<String> = Vec::new();
        if requested_family_leads(font_family, context.as_ref()) {
            candidates.push(font_family.to_string());
        }
        candidates.extend(fallback_candidates(
            font_family,
            context.as_ref(),
            ChainPurpose::Metrics,
        ));
        append_last_resort(&mut candidates, context.as_ref());
        candidates
    })
}

/// Resolve a document- or caller-provided in-memory face from the active
/// conversion context.
pub(crate) fn active_in_memory_font(
    font_family: &str,
    variant: typst::text::FontVariant,
) -> Option<typst::text::Font> {
    let candidates = family_candidates(font_family);
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        let context = context.as_ref()?;
        candidates
            .iter()
            .find_map(|candidate| context.in_memory_font(candidate, variant))
    })
}

/// The document- or caller-provided face registered under exactly
/// `font_family`, without walking its alias and substitute chain.
///
/// The native metric lookups walk that chain themselves and interleave each
/// candidate's in-memory face with its on-disk resolution, which is the
/// compiler's own order: a family held in memory outranks its copy on disk,
/// but a chain tail held in memory never outranks a candidate that resolves
/// earlier on a search path (issue #1629). [`active_in_memory_font`] answers
/// for the whole chain at once where no on-disk book exists.
pub(crate) fn active_in_memory_font_named(
    font_family: &str,
    variant: typst::text::FontVariant,
) -> Option<typst::text::Font> {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        active_context
            .borrow()
            .as_ref()?
            .in_memory_font(font_family, variant)
    })
}

/// Every document- or caller-provided face in the same priority order the
/// compiler prepends to its fallback font book.
pub(crate) fn active_in_memory_fonts() -> Vec<typst::text::Font> {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        active_context
            .borrow()
            .as_ref()
            .map(|context| context.in_memory_fonts().to_vec())
            .unwrap_or_default()
    })
}

/// The conversion-local filesystem font paths, when code generation is
/// running under a native font context.
///
/// Returning `Some` for an active context with no extra paths is intentional:
/// its fallback chain can still differ through a conversion-local last-resort
/// family, so family-only process caches are not authoritative in that scope.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn active_font_search_paths() -> Option<Vec<std::path::PathBuf>> {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        active_context
            .borrow()
            .as_ref()
            .map(|context| context.search_paths().to_vec())
    })
}

/// The font list for a run that states a Latin family and an East Asian one.
///
/// Word shapes a run's Latin codepoints with `w:ascii` and its East Asian ones
/// with `w:eastAsia`. Typst resolves a font list per glyph, falling through to
/// the next family for a character the current one has no glyph for, so
/// listing the Latin family first and the East Asian family straight after it
/// reproduces that split: a Latin face has no Hangul, so the Hangul lands on
/// the declared East Asian face rather than on whatever the Latin family's own
/// substitutes happen to cover (issue #575).
///
/// The script's own faces come next, then a Latin chain for symbol-block
/// glyphs, then the East Asian family's substitutes and the Latin family's, so
/// a document naming a face the system does not have still degrades the way
/// each family's chain says it should. Both outrank those substitutes for the
/// reason given on [`font_with_fallbacks_for_text`]: a family substitute
/// preserves metrics or class, which is the wrong priority for a glyph the
/// family does not have.
pub fn font_with_east_asian_fallbacks(
    latin_family: &str,
    east_asian_family: &str,
    text: &str,
) -> String {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        join_font_list(east_asian_family_chain(
            latin_family,
            east_asian_family,
            text,
            context.as_ref(),
        ))
    })
}

/// The family from an emitted font list that will paint `text`'s script.
///
/// Typst walks the list per glyph, whereas [`family_candidates`] resolves one
/// declared name through its metric-substitution chain. Those answers differ
/// when the declared face is present but does not cover the text: the Noto face
/// resolved for the Korean quotation fixture is first in the cell's list but
/// has no Hangul, so Malgun Gothic supplies both the glyphs and the line metrics
/// that must seat them (issue #1239).
///
/// Without an active font context, coverage is unknowable. In that case this
/// preserves the prior declared-family answer instead of guessing.
pub(crate) fn painted_family_for_text(
    latin_family: &str,
    east_asian_family: Option<&str>,
    text: &str,
) -> String {
    let script: TextScript = text_script(text);
    let declared_family: &str = if script == TextScript::Latin {
        latin_family
    } else {
        east_asian_family.unwrap_or(latin_family)
    };
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        let Some(context) = context.as_ref() else {
            return declared_family.to_string();
        };
        let families: Vec<String> = family_chain_for_text_with_context(
            latin_family,
            east_asian_family,
            text,
            Some(context),
        );
        families
            .into_iter()
            .find(|family| covers_script_with_alias(context, family, script))
            .unwrap_or_else(|| declared_family.to_string())
    })
}

/// The exact ordered family list the emitter states for one run.
///
/// Metric callers that must reproduce Typst's implicit fallback need the
/// entire list: when every named family is absent, Typst selects a fallback
/// from its font book rather than resolving one declared name in isolation.
pub(crate) fn family_chain_for_text(
    latin_family: &str,
    east_asian_family: Option<&str>,
    text: &str,
) -> Vec<String> {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let context = active_context.borrow();
        family_chain_for_text_with_context(latin_family, east_asian_family, text, context.as_ref())
    })
}

fn family_chain_for_text_with_context(
    latin_family: &str,
    east_asian_family: Option<&str>,
    text: &str,
    context: Option<&FontSearchContext>,
) -> Vec<String> {
    match east_asian_family {
        Some(east_asian) if !east_asian.eq_ignore_ascii_case(latin_family) => {
            east_asian_family_chain(latin_family, east_asian, text, context)
        }
        _ => latin_family_chain(latin_family, text, context),
    }
}

/// The ordered candidate list behind [`font_with_east_asian_fallbacks`], for
/// the same reason [`latin_family_chain`] exists.
fn east_asian_family_chain(
    latin_family: &str,
    east_asian_family: &str,
    text: &str,
    context: Option<&FontSearchContext>,
) -> Vec<String> {
    let mut families: Vec<String> = Vec::new();
    if requested_family_leads(latin_family, context) {
        families.push(latin_family.to_string());
    }
    if requested_family_leads(east_asian_family, context) {
        families.push(east_asian_family.to_string());
    }
    // The East Asian slot names the face whose voice must be kept — the
    // deck in #687 puts a serif Latin family there — and the Latin slot
    // decides when it says nothing.
    let serif_source: &str = if family_is_serif(east_asian_family) {
        east_asian_family
    } else {
        latin_family
    };
    families.extend(
        script_fallbacks(serif_source, text)
            .iter()
            .map(|face| (*face).to_string()),
    );
    families.extend(
        symbol_fallbacks(text)
            .iter()
            .map(|face| (*face).to_string()),
    );
    families.extend(fallback_candidates(
        east_asian_family,
        context,
        ChainPurpose::Paint,
    ));
    families.extend(fallback_candidates(
        latin_family,
        context,
        ChainPurpose::Paint,
    ));
    append_last_resort(&mut families, context);
    families
}

pub(crate) fn with_font_search_context<T>(
    context: Option<&FontSearchContext>,
    operation: impl FnOnce() -> T,
) -> T {
    ACTIVE_FONT_CONTEXT.with(|active_context| {
        let previous = active_context.replace(context.cloned());
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(operation));
        active_context.replace(previous);
        match result {
            Ok(value) => value,
            Err(panic) => std::panic::resume_unwind(panic),
        }
    })
}

/// Walk the IR tree rooted at a `Block`, calling `visitor` with each font
/// family encountered and the run text it was declared over. The visitor
/// returns `true` to continue walking or `false` to short-circuit. Returns
/// `false` when the visitor short-circuited.
///
/// The text travels with the family because the face a run ends up in depends
/// on both: the same family resolves to a different face over Hangul than over
/// Latin (issue #617).
fn visit_block_fonts(block: &Block, visitor: &mut impl FnMut(&str, &str) -> bool) -> bool {
    match block {
        Block::Paragraph(paragraph) => visit_paragraph_fonts(paragraph, visitor),
        // A contents page's entries are built from what it points at, so it
        // names no font of its own.
        Block::TableOfContents(_) => true,
        Block::Caption(caption) => visit_paragraph_fonts(&caption.paragraph, visitor),
        Block::Table(table) => visit_table_fonts(table, visitor),
        Block::FloatingTextBox(text_box) => visit_blocks_fonts(&text_box.content, visitor),
        Block::List(list) => list.items.iter().all(|item| {
            item.content
                .iter()
                .all(|paragraph| visit_paragraph_fonts(paragraph, visitor))
        }),
        Block::Chart(chart) => visit_chart_fonts(chart, visitor),
        Block::Image(_)
        | Block::InlineImages(_)
        | Block::FloatingImage(_)
        | Block::FloatingShape(_)
        | Block::MathEquation(_)
        | Block::PageBreak
        | Block::ColumnBreak => true,
    }
}

/// Offer the chart default and axis faces to the visitor with their text sample.
///
/// Chart strings decide each face's fallback chain as a run's text does.
/// A document whose *only* font request
/// comes from a chart still needs the font search context, or the directories
/// holding the requested face are never scanned and the chart falls back to the
/// engine's default — which is the very thing resolving the theme font was
/// meant to stop (issues #668, #461).
pub(super) fn visit_chart_fonts(
    chart: &crate::ir::Chart,
    visitor: &mut impl FnMut(&str, &str) -> bool,
) -> bool {
    let sample: String = chart.text_sample();
    [
        chart.text_font_family.as_deref(),
        chart.category_axis_text_font_family.as_deref(),
        chart.value_axis_text_font_family.as_deref(),
        // The secondary axis names its face independently too (issue #1374).
        chart
            .secondary_value_axis
            .as_ref()
            .and_then(|axis| axis.text_font_family.as_deref()),
    ]
    .into_iter()
    .flatten()
    .all(|family| visitor(family, &sample))
}

/// Walk a slice of blocks, calling `visitor` for each font family found.
fn visit_blocks_fonts(blocks: &[Block], visitor: &mut impl FnMut(&str, &str) -> bool) -> bool {
    blocks.iter().all(|block| visit_block_fonts(block, visitor))
}

/// Walk a `Paragraph`'s runs, calling `visitor` for each font family.
fn visit_paragraph_fonts(
    paragraph: &Paragraph,
    visitor: &mut impl FnMut(&str, &str) -> bool,
) -> bool {
    paragraph.runs.iter().all(|run| {
        declared_family(run.style.font_family.as_deref())
            .is_none_or(|family| visitor(family, &run.text))
    })
}

/// Walk a `Table`'s cells, calling `visitor` for each font family found.
fn visit_table_fonts(table: &Table, visitor: &mut impl FnMut(&str, &str) -> bool) -> bool {
    table.rows.iter().all(|row| {
        row.cells
            .iter()
            .all(|cell| visit_blocks_fonts(&cell.content, visitor))
    })
}

/// Walk a `HeaderFooter`'s paragraphs, calling `visitor` for each font family.
fn visit_header_footer_fonts(
    header_footer: &HeaderFooter,
    visitor: &mut impl FnMut(&str, &str) -> bool,
) -> bool {
    header_footer.paragraphs.iter().all(|paragraph| {
        paragraph.elements.iter().all(|inline| match inline {
            HFInline::Run(run) => declared_family(run.style.font_family.as_deref())
                .is_none_or(|family| visitor(family, &run.text)),
            HFInline::Image(_)
            | HFInline::PageNumber(_)
            | HFInline::TotalPages(_)
            | HFInline::PositionedTab(_) => true,
        })
    })
}

/// The family a run declares, or `None` when it leaves the choice to defaults.
fn declared_family(font_family: Option<&str>) -> Option<&str> {
    font_family
        .map(str::trim)
        .filter(|family| !family.is_empty())
}

fn block_requests_font_family(block: &Block) -> bool {
    !visit_block_fonts(block, &mut font_family_uses_context_free_fallbacks)
}

fn table_requests_font_family(table: &Table) -> bool {
    !visit_table_fonts(table, &mut font_family_uses_context_free_fallbacks)
}

fn header_footer_requests_font_family(header_footer: &HeaderFooter) -> bool {
    !visit_header_footer_fonts(header_footer, &mut font_family_uses_context_free_fallbacks)
}

fn font_family_uses_context_free_fallbacks(font_family: &str, _text: &str) -> bool {
    // These two families' static substitution chains are sufficient for Typst
    // to select the installed face, so the separate availability scan is
    // skipped for them. Both are ubiquitous, and Times New Roman is also the
    // face a DOCX naming no `w:rFonts` resolves (issue #1196) — the role Arial
    // held before, and the reason the skip has to move with it. Without it,
    // every package that states no font at all reports a substitution on any
    // host that does not install the face: the scan is what emits
    // `ConvertWarning::FallbackUsed`, and a synthetic default is not a request
    // the document made.
    CONTEXT_FREE_FALLBACK_FAMILIES
        .iter()
        .any(|family| font_family.eq_ignore_ascii_case(family))
}

/// Families the availability scan skips; see
/// [`font_family_uses_context_free_fallbacks`].
const CONTEXT_FREE_FALLBACK_FAMILIES: [&str; 2] = ["Arial", "Times New Roman"];

/// A family as declared, paired with the script of the text it was declared
/// over. Both halves decide which face the run ends up in.
type FontRequest = (String, TextScript);

fn collect_block_fonts(block: &Block, fonts: &mut BTreeSet<FontRequest>) {
    visit_block_fonts(block, &mut |font, text| {
        fonts.insert((font.to_string(), text_script(text)));
        true
    });
}

fn collect_table_fonts(table: &Table, fonts: &mut BTreeSet<FontRequest>) {
    visit_table_fonts(table, &mut |font, text| {
        fonts.insert((font.to_string(), text_script(text)));
        true
    });
}

fn collect_header_footer_fonts(header_footer: &HeaderFooter, fonts: &mut BTreeSet<FontRequest>) {
    visit_header_footer_fonts(header_footer, &mut |font, text| {
        fonts.insert((font.to_string(), text_script(text)));
        true
    });
}

fn collect_document_font_requests(doc: &Document) -> BTreeSet<FontRequest> {
    let mut fonts = BTreeSet::new();

    for page in &doc.pages {
        match page {
            Page::Flow(page) => {
                if let Some(header) = &page.header {
                    collect_header_footer_fonts(header, &mut fonts);
                }
                if let Some(footer) = &page.footer {
                    collect_header_footer_fonts(footer, &mut fonts);
                }
                for block in &page.content {
                    collect_block_fonts(block, &mut fonts);
                }
            }
            Page::Fixed(page) => {
                for element in &page.elements {
                    match &element.kind {
                        FixedElementKind::TextBox(text_box) => {
                            for block in &text_box.content {
                                collect_block_fonts(block, &mut fonts);
                            }
                        }
                        FixedElementKind::Table(table) => collect_table_fonts(table, &mut fonts),
                        FixedElementKind::Chart(chart) => {
                            visit_chart_fonts(chart, &mut |font, text| {
                                fonts.insert((font.to_string(), text_script(text)));
                                true
                            });
                        }
                        FixedElementKind::Image(_)
                        | FixedElementKind::Shape(_)
                        | FixedElementKind::SmartArt(_) => {}
                    }
                }
            }
            Page::Sheet(page) => {
                if let Some(header) = &page.header {
                    collect_header_footer_fonts(header, &mut fonts);
                }
                if let Some(footer) = &page.footer {
                    collect_header_footer_fonts(footer, &mut fonts);
                }
                collect_table_fonts(&page.table, &mut fonts);
                for text_box in &page.text_boxes {
                    for paragraph in &text_box.paragraphs {
                        collect_block_fonts(&Block::Paragraph(paragraph.clone()), &mut fonts);
                    }
                }
            }
        }
    }

    fonts
}

/// Whether this document names one of the unembedded poster faces whose
/// reproducible fallback is the bundled Noto Serif 2.015 (issue #1458).
///
/// Loading the two faces only for a document that can select them keeps the
/// generic Typst fallback order unchanged for every unrelated conversion.
pub(crate) fn document_requests_bundled_noto_serif(doc: &Document) -> bool {
    collect_document_font_requests(doc)
        .into_iter()
        .any(|(family, _)| {
            matches!(
                normalized_lookup_key(&family).as_str(),
                "avenir next lt pro" | "avenir next w1g medium" | "the hand black"
            )
        })
}

/// Whether this document names Aptos, whose reproducible fallback bundle is
/// Noto Sans 2.015. Explicitly supplied Aptos still leads the paint chain
/// (issue #1463).
pub(crate) fn document_requests_bundled_noto_sans(doc: &Document) -> bool {
    collect_document_font_requests(doc)
        .into_iter()
        .any(|(family, _)| normalized_lookup_key(&family) == "aptos")
}

/// Whether this document names Segoe UI, whose bundled OFL fallback is
/// Selawik 1.01 (issue #1472). Regular Basic Latin advances match, but its
/// line metrics differ from the source face (#1603).
pub(crate) fn document_requests_bundled_selawik(doc: &Document) -> bool {
    collect_document_font_requests(doc)
        .into_iter()
        .any(|(family, _)| normalized_lookup_key(&family) == "segoe ui")
}

pub(crate) fn document_requests_font_families(doc: &Document) -> bool {
    doc.pages.iter().any(|page| match page {
        Page::Flow(page) => {
            page.header
                .as_ref()
                .is_some_and(header_footer_requests_font_family)
                || page
                    .footer
                    .as_ref()
                    .is_some_and(header_footer_requests_font_family)
                || page.content.iter().any(block_requests_font_family)
        }
        Page::Fixed(page) => page.elements.iter().any(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => {
                text_box.content.iter().any(block_requests_font_family)
            }
            FixedElementKind::Table(table) => table_requests_font_family(table),
            // A slide whose only font request is a chart's face still needs the
            // context resolved (#668).
            FixedElementKind::Chart(chart) => {
                !visit_chart_fonts(chart, &mut font_family_uses_context_free_fallbacks)
            }
            FixedElementKind::Image(_)
            | FixedElementKind::Shape(_)
            | FixedElementKind::SmartArt(_) => false,
        }),
        Page::Sheet(page) => {
            page.header
                .as_ref()
                .is_some_and(header_footer_requests_font_family)
                || page
                    .footer
                    .as_ref()
                    .is_some_and(header_footer_requests_font_family)
                || table_requests_font_family(&page.table)
                // Worksheet drawings carry their own runs; a workbook whose
                // only font request comes from a shape label still needs the
                // font search context, or the compiler never sees the
                // directories that hold the requested face (issue #461).
                || page.text_boxes.iter().any(sheet_text_box_requests_font_family)
        }
    })
}

fn sheet_text_box_requests_font_family(text_box: &crate::ir::SheetTextBox) -> bool {
    text_box
        .paragraphs
        .iter()
        .any(|paragraph| block_requests_font_family(&Block::Paragraph(paragraph.clone())))
}

/// The face a run declaring `font_family` over `script` actually renders in, or
/// `None` when the declared family is installed and nothing is substituted.
///
/// The candidate order mirrors [`font_with_fallbacks_for_text`] exactly — the
/// script's own faces ahead of the family's substitute chain. Consulting only
/// the substitutes named a Chinese face for Korean text that renders in Malgun
/// Gothic, sending anyone reading the log after a substitution that never
/// happened (issue #617).
fn resolve_available_fallback(
    font_family: &str,
    script: TextScript,
    context: &FontSearchContext,
) -> Option<String> {
    if requested_family_leads(font_family, Some(context))
        && family_covers_or_is_unindexed(context, font_family, script)
    {
        return None;
    }

    script
        .fallbacks(family_is_serif(font_family))
        .iter()
        .map(|face| (*face).to_string())
        .chain(fallback_candidates(
            font_family,
            Some(context),
            ChainPurpose::Paint,
        ))
        .chain(context.last_resort_font_family().map(str::to_string))
        .find(|candidate| family_covers_or_is_unindexed(context, candidate, script))
        .or_else(|| (script != TextScript::Latin).then(|| ".notdef".to_string()))
}

fn family_covers_or_is_unindexed(
    context: &FontSearchContext,
    family: &str,
    script: TextScript,
) -> bool {
    context.has_family(family)
        && (!context.knows_script_coverage(family) || context.covers_script(family, script))
}

pub(crate) fn detect_missing_font_fallbacks_with_context(
    doc: &Document,
    context: &FontSearchContext,
) -> Vec<(String, String)> {
    let requested_fonts = collect_document_font_requests(doc);
    if requested_fonts.is_empty() {
        return Vec::new();
    }

    // A family resolving differently per script yields one warning per
    // resolution; a family used over several scripts that all land on the same
    // face yields one, hence the set rather than a plain map.
    requested_fonts
        .into_iter()
        .filter_map(|(font, script)| {
            resolve_available_fallback(&font, script, context).map(|to| (font, to))
        })
        .collect::<BTreeSet<(String, String)>>()
        .into_iter()
        .collect()
}

#[cfg(target_arch = "wasm32")]
pub fn detect_missing_font_fallbacks(
    _doc: &Document,
    _font_paths: &[PathBuf],
) -> Vec<(String, String)> {
    Vec::new()
}

#[cfg(test)]
#[path = "font_subst_tests.rs"]
mod tests;
