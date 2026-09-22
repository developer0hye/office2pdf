// Filesystem and system-font discovery are native-only. The shared family
// index also backs document- and caller-provided in-memory fonts; path-facing
// members remain compiled on WASM but are intentionally unused there.
#![cfg_attr(target_arch = "wasm32", allow(dead_code))]

use std::collections::{HashMap, HashSet};
use std::path::PathBuf;

use super::font_subst::TextScript;

#[cfg(not(target_arch = "wasm32"))]
use tracing::debug;

#[derive(Debug, Clone, Default)]
pub(crate) struct FontSearchContext {
    search_paths: Vec<PathBuf>,
    available_families: HashSet<String>,
    office_families: HashSet<String>,
    user_families: HashSet<String>,
    /// Families that ship at least one italic or oblique face. A family that
    /// is available but absent here renders `style: "italic"` upright, which
    /// is what the synthetic oblique exists to cover (issue #686).
    italic_families: HashSet<String>,
    /// Which scripts each available family has glyphs for, as
    /// [`script_bit`] flags. Deciding *which* family a run's text lands on
    /// needs this: the declared family leads the font list even for a script
    /// it cannot write, and Typst falls through to the next entry per glyph.
    family_scripts: HashMap<String, u8>,
    /// The OS/2 weight classes each available family ships a face at. A
    /// request naming a weight member — `Calibri Light`, `Segoe UI Semibold`
    /// — is only that member where the base family actually holds a face at
    /// the weight the name states (issue #1286).
    family_weights: HashMap<String, HashSet<u16>>,
    /// Filesystem-free faces available to WASM metric lookups while codegen is
    /// running under this context.
    in_memory_book: typst::text::FontBook,
    in_memory_fonts: Vec<typst::text::Font>,
    /// Caller-selected final family for every emitted Typst font chain.
    last_resort_font_family: Option<String>,
}

impl FontSearchContext {
    pub(crate) fn search_paths(&self) -> &[PathBuf] {
        &self.search_paths
    }

    pub(crate) fn has_family(&self, family: &str) -> bool {
        self.available_families
            .contains(&normalize_family_name(family))
    }

    /// Whether `family` ships a real italic or oblique face.
    pub(crate) fn has_italic_face(&self, family: &str) -> bool {
        self.italic_families
            .contains(&normalize_family_name(family))
    }

    /// Whether `family` has glyphs for `script`.
    ///
    /// Unknown families answer `false`: a face the context never saw cannot be
    /// shown to cover anything, and every caller treats "unknown" as "keep
    /// looking".
    pub(crate) fn covers_script(&self, family: &str, script: TextScript) -> bool {
        self.family_scripts
            .get(&normalize_family_name(family))
            .is_some_and(|scripts| scripts & script_bit(script) != 0)
    }

    /// Whether `family` ships a face at exactly `weight`.
    ///
    /// Exact because the font book indexes a weight member under its base
    /// family: `calibril.ttf` is `Calibri` at 300, and asking whether that
    /// family holds a 300 face is the only way to ask whether the light member
    /// exists at all (issue #1286).
    pub(crate) fn has_face_weight(&self, family: &str, weight: typst::text::FontWeight) -> bool {
        self.family_weights
            .get(&normalize_family_name(family))
            .is_some_and(|weights| weights.contains(&weight.to_number()))
    }

    pub(crate) fn knows_script_coverage(&self, family: &str) -> bool {
        self.family_scripts
            .contains_key(&normalize_family_name(family))
    }

    pub(crate) fn family_source_rank(&self, family: &str) -> u8 {
        let normalized = normalize_family_name(family);
        if self.office_families.contains(&normalized) {
            0
        } else if self.user_families.contains(&normalized) {
            1
        } else if self.available_families.contains(&normalized) {
            2
        } else {
            3
        }
    }

    /// Whether a face came from an explicit conversion input: a caller font
    /// path, a package-embedded font directory, or registered font bytes.
    /// System and auto-discovered Office fonts are deliberately excluded.
    pub(crate) fn is_user_family(&self, family: &str) -> bool {
        self.user_families.contains(&normalize_family_name(family))
    }

    pub(crate) fn last_resort_font_family(&self) -> Option<&str> {
        self.last_resort_font_family.as_deref()
    }

    /// Attach a final fallback family, ignoring an empty or whitespace-only
    /// name so native callers cannot emit an invalid empty Typst family.
    pub(crate) fn with_last_resort_family(mut self, family: Option<&str>) -> Self {
        self.last_resort_font_family = family
            .map(str::trim)
            .filter(|family| !family.is_empty())
            .map(str::to_string);
        self
    }

    /// Merge caller-provided faces into this context without involving the
    /// filesystem. The same face list backs metric lookup and compilation.
    pub(crate) fn with_in_memory_fonts(mut self, fonts: &[typst::text::Font]) -> Self {
        if fonts.is_empty() {
            return self;
        }

        self.in_memory_fonts.extend_from_slice(fonts);
        self.in_memory_book = typst::text::FontBook::from_fonts(&self.in_memory_fonts);

        let FamilyIndex {
            available_families,
            italic_families,
            family_scripts,
            family_weights,
        } = index_families_from_book(&self.in_memory_book);
        self.user_families
            .extend(available_families.iter().cloned());
        self.available_families.extend(available_families);
        self.italic_families.extend(italic_families);
        for (family, scripts) in family_scripts {
            *self.family_scripts.entry(family).or_default() |= scripts;
        }
        for (family, weights) in family_weights {
            self.family_weights
                .entry(family)
                .or_default()
                .extend(weights);
        }
        self
    }

    pub(crate) fn in_memory_font(
        &self,
        family: &str,
        variant: typst::text::FontVariant,
    ) -> Option<typst::text::Font> {
        self.in_memory_book
            .select(&normalize_family_name(family), variant)
            .and_then(|index| self.in_memory_fonts.get(index))
            .cloned()
    }

    /// Conversion-local faces in the order they precede the compiler's
    /// fallback font slots.
    pub(crate) fn in_memory_fonts(&self) -> &[typst::text::Font] {
        &self.in_memory_fonts
    }

    #[cfg(test)]
    pub(crate) fn for_test(
        search_paths: Vec<PathBuf>,
        available_families: &[&str],
        office_families: &[&str],
        user_families: &[&str],
    ) -> Self {
        Self {
            search_paths,
            available_families: available_families
                .iter()
                .map(|family| normalize_family_name(family))
                .collect(),
            office_families: office_families
                .iter()
                .map(|family| normalize_family_name(family))
                .collect(),
            user_families: user_families
                .iter()
                .map(|family| normalize_family_name(family))
                .collect(),
            italic_families: HashSet::new(),
            family_scripts: HashMap::new(),
            family_weights: HashMap::new(),
            in_memory_book: typst::text::FontBook::new(),
            in_memory_fonts: Vec::new(),
            last_resort_font_family: None,
        }
    }

    /// Declare, for a test, which families ship an italic face and which
    /// scripts each one writes. Kept off [`Self::for_test`] so the existing
    /// call sites that care only about availability stay as they are.
    #[cfg(test)]
    pub(crate) fn with_italic_and_scripts(
        mut self,
        italic_families: &[&str],
        family_scripts: &[(&str, &[TextScript])],
    ) -> Self {
        self.italic_families = italic_families
            .iter()
            .map(|family| normalize_family_name(family))
            .collect();
        self.family_scripts = family_scripts
            .iter()
            .map(|(family, scripts)| {
                let bits: u8 = scripts
                    .iter()
                    .fold(0, |bits, script| bits | script_bit(*script));
                (normalize_family_name(family), bits)
            })
            .collect();
        self
    }
}

fn normalize_family_name(family: &str) -> String {
    family.trim().to_ascii_lowercase()
}

/// The flag [`FontSearchContext::family_scripts`] stores `script` under.
fn script_bit(script: TextScript) -> u8 {
    match script {
        TextScript::Latin => 1,
        TextScript::Korean => 2,
        TextScript::Japanese => 4,
        TextScript::Chinese => 8,
    }
}

/// A character whose presence in a face proves it writes `script`.
///
/// Latin is probed with a plain capital rather than with the whole alphabet:
/// a face carrying `A` and not the rest of ASCII does not exist in practice,
/// and every extra probe costs a coverage lookup per family.
const SCRIPT_PROBES: [(TextScript, char); 5] = [
    (TextScript::Latin, 'A'),
    (TextScript::Korean, '가'),
    (TextScript::Japanese, 'あ'),
    // `漢` is outside GB2312 while `中` is inside it. Keeping both recognizes
    // the existing broad CJK faces and a strict Simplified Chinese subset.
    (TextScript::Chinese, '中'),
    (TextScript::Chinese, '漢'),
];

// Test-only probe counting full context resolutions, so the contract that a
// default-path lookup never re-indexes the host's fonts is assertable: one
// resolution parses every face on the host, and the fallback metric path
// once paid that per measured run. (A regular comment: rustc discards doc
// comments attached to macro invocations.)
#[cfg(test)]
thread_local! {
    pub(super) static FONT_CONTEXT_RESOLUTIONS: std::cell::Cell<usize> =
        const { std::cell::Cell::new(0) };
}

/// The Office application font directories a conversion searches on top of
/// the system fonts when the caller supplied no font path: the bundles under
/// `/Applications` and `~/Applications` on macOS, nothing elsewhere.
///
/// [`resolve_font_search_context`] with no user paths yields these same paths
/// but also indexes every face on the host, which is far too slow to repeat
/// per measured run. This is at most six directory probes, done once per
/// process: an Office installed after the first conversion is seen by the
/// next process.
#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn default_font_search_paths() -> &'static [PathBuf] {
    static PATHS: std::sync::LazyLock<Vec<PathBuf>> = std::sync::LazyLock::new(|| {
        if cfg!(target_os = "macos") {
            discover_default_macos_office_font_paths()
        } else {
            Vec::new()
        }
    });
    &PATHS
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn resolve_font_search_context(user_font_paths: &[PathBuf]) -> FontSearchContext {
    #[cfg(test)]
    FONT_CONTEXT_RESOLUTIONS.with(|count| count.set(count.get() + 1));
    let office_paths: &[PathBuf] = default_font_search_paths();
    let user_paths = canonicalize_existing_dirs(user_font_paths.iter().cloned());
    let search_paths = merge_prioritized_paths(office_paths, &user_paths);
    let office_families = available_families_from_paths(office_paths, false);
    let user_families = available_families_from_paths(&user_paths, false);
    let FamilyIndex {
        available_families,
        italic_families,
        family_scripts,
        family_weights,
    } = index_families_from_paths(&search_paths, true);

    debug!(
        office_path_count = office_paths.len(),
        user_path_count = user_paths.len(),
        search_path_count = search_paths.len(),
        available_family_count = available_families.len(),
        italic_family_count = italic_families.len(),
        "resolved font search context"
    );

    FontSearchContext {
        search_paths,
        available_families,
        office_families,
        user_families,
        italic_families,
        family_scripts,
        family_weights,
        in_memory_book: typst::text::FontBook::new(),
        in_memory_fonts: Vec::new(),
        last_resort_font_family: None,
    }
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn resolve_font_search_context(_user_font_paths: &[PathBuf]) -> FontSearchContext {
    FontSearchContext::default()
}

/// Build the substitution index for document- or caller-provided in-memory
/// faces.
///
/// These faces have the same priority as an explicit native font path: they
/// were supplied explicitly for this conversion and must lead Typst's fallback
/// fonts during family and script selection.
#[cfg_attr(not(target_arch = "wasm32"), allow(dead_code))]
pub(crate) fn resolve_font_search_context_from_fonts(
    fonts: &[typst::text::Font],
) -> FontSearchContext {
    FontSearchContext::default().with_in_memory_fonts(fonts)
}

#[cfg(not(target_arch = "wasm32"))]
fn available_families_from_paths(paths: &[PathBuf], include_system_fonts: bool) -> HashSet<String> {
    index_families_from_paths(paths, include_system_fonts).available_families
}

/// What one font search says about the families it found.
struct FamilyIndex {
    available_families: HashSet<String>,
    italic_families: HashSet<String>,
    family_scripts: HashMap<String, u8>,
    family_weights: HashMap<String, HashSet<u16>>,
}

#[cfg(not(target_arch = "wasm32"))]
fn index_families_from_paths(paths: &[PathBuf], include_system_fonts: bool) -> FamilyIndex {
    let book = super::pdf::discover_font_book(paths, include_system_fonts, include_system_fonts);
    index_families_from_book(&book)
}

fn index_families_from_book(book: &typst::text::FontBook) -> FamilyIndex {
    use typst::text::FontStyle;

    let mut index = FamilyIndex {
        available_families: HashSet::new(),
        italic_families: HashSet::new(),
        family_scripts: HashMap::new(),
        family_weights: HashMap::new(),
    };
    for (family, indices) in book.families() {
        let key: String = normalize_family_name(family);
        let infos: Vec<&typst::text::FontInfo> =
            indices.filter_map(|index| book.info(index)).collect();
        if infos
            .iter()
            .any(|info| info.variant.style != FontStyle::Normal)
        {
            index.italic_families.insert(key.clone());
        }
        // A family's scripts are the union over its faces: a face-by-face
        // answer would report the Hangul-less italic member of a family whose
        // regular member does carry Hangul.
        let scripts: u8 = SCRIPT_PROBES
            .iter()
            .filter(|(_, probe)| {
                infos
                    .iter()
                    .any(|info| info.coverage.contains(*probe as u32))
            })
            .fold(0, |bits, (script, _)| bits | script_bit(*script));
        index.family_scripts.insert(key.clone(), scripts);
        index.family_weights.insert(
            key.clone(),
            infos
                .iter()
                .map(|info| info.variant.weight.to_number())
                .collect(),
        );
        index.available_families.insert(key);
    }
    index
}

#[cfg(not(target_arch = "wasm32"))]
fn discover_default_macos_office_font_paths() -> Vec<PathBuf> {
    let mut app_roots = vec![PathBuf::from("/Applications")];
    if let Some(home_dir) = std::env::var_os("HOME").map(PathBuf::from) {
        app_roots.push(home_dir.join("Applications"));
    }
    discover_macos_office_font_paths_from(&app_roots)
}

#[cfg(not(target_arch = "wasm32"))]
fn discover_macos_office_font_paths_from(app_roots: &[PathBuf]) -> Vec<PathBuf> {
    // Installed Office application resources are versioned with the app. The
    // per-user CloudFonts and PreviewFont caches are mutable implementation
    // details: their contents depend on which documents the user opened, so
    // auto-discovering them makes an unembedded family render differently on
    // otherwise identical hosts. Callers that intentionally want one of those
    // directories can still pass it as an explicit font path (issue #1409).
    canonicalize_existing_dirs(office_app_font_dir_candidates(app_roots))
}

#[cfg(not(target_arch = "wasm32"))]
fn office_app_font_dir_candidates(app_roots: &[PathBuf]) -> Vec<PathBuf> {
    let mut candidates = Vec::new();
    for app_root in app_roots {
        for app_name in [
            "Microsoft PowerPoint.app",
            "Microsoft Word.app",
            "Microsoft Excel.app",
        ] {
            candidates.push(app_root.join(app_name).join("Contents/Resources/DFonts"));
        }
    }
    candidates
}

#[cfg(not(target_arch = "wasm32"))]
fn merge_prioritized_paths(primary: &[PathBuf], secondary: &[PathBuf]) -> Vec<PathBuf> {
    let mut merged = Vec::with_capacity(primary.len() + secondary.len());
    let mut seen = HashSet::new();
    for path in primary.iter().chain(secondary) {
        if seen.insert(path.clone()) {
            merged.push(path.clone());
        }
    }
    merged
}

#[cfg(not(target_arch = "wasm32"))]
fn canonicalize_existing_dirs<I>(paths: I) -> Vec<PathBuf>
where
    I: IntoIterator<Item = PathBuf>,
{
    let mut canonicalized = Vec::new();
    let mut seen = HashSet::new();
    for path in paths {
        let Ok(canonical) = std::fs::canonicalize(&path) else {
            debug!(path = ?path, "skipping missing font directory");
            continue;
        };
        if !canonical.is_dir() {
            debug!(path = ?canonical, "skipping non-directory font path");
            continue;
        }
        if seen.insert(canonical.clone()) {
            canonicalized.push(canonical);
        }
    }
    canonicalized
}

/// Faces built for tests from the tracked Noto Serif, so a test can index a
/// real font book without depending on what the host installs.
#[cfg(test)]
pub(crate) mod test_faces {
    use typst::foundations::Bytes;
    use typst::text::Font;

    /// The tracked Noto Serif face with its OS/2 `usWeightClass` rewritten.
    ///
    /// Typst indexes a face under its trimmed family name at the weight its
    /// OS/2 table states, so the rewritten face lands as `Noto Serif` at
    /// `weight_class` — the shape of `calibril.ttf`, which the book registers
    /// as `Calibri` at 300 and never as `Calibri Light` (issue #1286).
    pub(crate) fn noto_serif_at_weight(weight_class: u16) -> Font {
        rewritten_noto_serif(weight_class, None)
    }

    /// [`noto_serif_at_weight`] with its declared ascender rewritten too, so
    /// two members of one family disagree on the line box they declare.
    ///
    /// The tracked Noto Serif ships one line box for every weight, which
    /// cannot tell "the member the name denotes" apart from "the family's
    /// regular member". Real families do: `Arial Black` declares an `hhea`
    /// ascender of 2254/2048 against Arial's 1854/2048 (issue #1643).
    ///
    /// Both the `hhea` ascender and OS/2 `sTypoAscender` are rewritten, so the
    /// face declares the taller line whichever table the reader consults:
    /// Noto Serif sets `USE_TYPO_METRICS`, and `ttf_parser`'s `ascender()`
    /// answers from OS/2 rather than `hhea` whenever that bit is set.
    pub(crate) fn noto_serif_at_weight_with_ascender(weight_class: u16, ascender: i16) -> Font {
        rewritten_noto_serif(weight_class, Some(ascender))
    }

    /// Offset of `tag`'s table in a TrueType face's table directory.
    fn table_offset(bytes: &[u8], tag: &[u8; 4]) -> usize {
        let table_count = usize::from(u16::from_be_bytes([bytes[4], bytes[5]]));
        (0..table_count)
            .map(|record| 12 + record * 16)
            .find(|&record| &bytes[record..record + 4] == tag)
            .map(|record| {
                u32::from_be_bytes([
                    bytes[record + 8],
                    bytes[record + 9],
                    bytes[record + 10],
                    bytes[record + 11],
                ]) as usize
            })
            .expect("a TrueType face carries this table")
    }

    /// The tracked Noto Serif with its declared weight, and optionally its
    /// ascender, rewritten in place.
    ///
    /// Neither `ttf-parser` nor Typst verifies a table checksum, so patching
    /// the fixed-layout fields in place needs no rebuild of the directory.
    fn rewritten_noto_serif(weight_class: u16, ascender: Option<i16>) -> Font {
        /// Byte offset of `sTypoAscender` within an OS/2 table of any version.
        const TYPO_ASCENDER_OFFSET: usize = 68;
        let mut bytes: Vec<u8> = include_bytes!("../../fonts/NotoSerif-Regular.ttf").to_vec();
        let os2_offset: usize = table_offset(&bytes, b"OS/2");
        // `OS/2`: version (2 bytes), xAvgCharWidth (2), usWeightClass (2).
        bytes[os2_offset + 4..os2_offset + 6].copy_from_slice(&weight_class.to_be_bytes());
        if let Some(ascender) = ascender {
            let hhea_offset: usize = table_offset(&bytes, b"hhea");
            // `hhea`: version (4 bytes), ascender (2).
            bytes[hhea_offset + 4..hhea_offset + 6].copy_from_slice(&ascender.to_be_bytes());
            let typo_ascender: usize = os2_offset + TYPO_ASCENDER_OFFSET;
            bytes[typo_ascender..typo_ascender + 2].copy_from_slice(&ascender.to_be_bytes());
        }
        Font::new(Bytes::new(bytes), 0).expect("the rewritten face parses")
    }
}

#[cfg(test)]
#[path = "font_context_tests.rs"]
mod tests;
