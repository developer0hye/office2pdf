use std::collections::HashSet;
use std::path::{Path, PathBuf};

#[cfg(not(target_arch = "wasm32"))]
use typst_kit::fonts::FontSearcher;

#[cfg(not(target_arch = "wasm32"))]
use tracing::debug;

#[derive(Debug, Clone, Default)]
pub(crate) struct FontSearchContext {
    search_paths: Vec<PathBuf>,
    available_families: HashSet<String>,
    office_families: HashSet<String>,
    user_families: HashSet<String>,
}

impl FontSearchContext {
    pub(crate) fn search_paths(&self) -> &[PathBuf] {
        &self.search_paths
    }

    pub(crate) fn has_family(&self, family: &str) -> bool {
        self.available_families
            .contains(&normalize_family_name(family))
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
        }
    }
}

fn normalize_family_name(family: &str) -> String {
    family.trim().to_ascii_lowercase()
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn resolve_font_search_context(user_font_paths: &[PathBuf]) -> FontSearchContext {
    let office_paths = if cfg!(target_os = "macos") {
        discover_default_macos_office_font_paths()
    } else {
        Vec::new()
    };
    let user_paths = canonicalize_existing_dirs(user_font_paths.iter().cloned());
    let search_paths = merge_prioritized_paths(&office_paths, &user_paths);
    let office_families = available_families_from_paths(&office_paths, false);
    let user_families = available_families_from_paths(&user_paths, false);
    let available_families = available_families_from_paths(&search_paths, true);

    debug!(
        office_path_count = office_paths.len(),
        user_path_count = user_paths.len(),
        search_path_count = search_paths.len(),
        available_family_count = available_families.len(),
        "resolved font search context"
    );

    FontSearchContext {
        search_paths,
        available_families,
        office_families,
        user_families,
    }
}

#[cfg(target_arch = "wasm32")]
pub(crate) fn resolve_font_search_context(_user_font_paths: &[PathBuf]) -> FontSearchContext {
    FontSearchContext::default()
}

#[cfg(not(target_arch = "wasm32"))]
fn available_families_from_paths(paths: &[PathBuf], include_system_fonts: bool) -> HashSet<String> {
    let mut searcher = FontSearcher::new();
    searcher.include_system_fonts(include_system_fonts);
    searcher.include_embedded_fonts(include_system_fonts);
    let font_data = if paths.is_empty() {
        searcher.search()
    } else {
        searcher.search_with(paths.iter().map(|path| path.as_path()))
    };
    font_data
        .book
        .families()
        .map(|(family, _)| normalize_family_name(family))
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
fn discover_default_macos_office_font_paths() -> Vec<PathBuf> {
    let mut app_roots = vec![PathBuf::from("/Applications")];
    let Some(home_dir) = std::env::var_os("HOME").map(PathBuf::from) else {
        return canonicalize_existing_dirs(
            office_app_font_dir_candidates(&app_roots)
                .into_iter()
                .collect::<Vec<PathBuf>>(),
        );
    };
    app_roots.push(home_dir.join("Applications"));
    discover_macos_office_font_paths_from(&app_roots, &home_dir)
}

#[cfg(not(target_arch = "wasm32"))]
fn discover_macos_office_font_paths_from(app_roots: &[PathBuf], home_dir: &Path) -> Vec<PathBuf> {
    let mut candidates = office_app_font_dir_candidates(app_roots);
    if let Some(font_cache_root) = highest_numeric_child(
        &home_dir.join("Library/Group Containers/UBF8T346G9.Office/FontCache"),
    ) {
        candidates.push(font_cache_root.join("CloudFonts"));
        candidates.push(font_cache_root.join("PreviewFont"));
    }
    candidates.extend(wps_app_font_dir_candidates(app_roots));
    canonicalize_existing_dirs(candidates)
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
fn wps_app_font_dir_candidates(app_roots: &[PathBuf]) -> Vec<PathBuf> {
    app_roots
        .iter()
        .map(|app_root| app_root.join("wpsoffice.app/Contents/Resources/office6/fonts"))
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
fn highest_numeric_child(root: &Path) -> Option<PathBuf> {
    let entries = std::fs::read_dir(root).ok()?;
    entries
        .flatten()
        .filter_map(|entry| {
            let path = entry.path();
            if !path.is_dir() {
                return None;
            }
            let version = path.file_name()?.to_str()?.parse::<u32>().ok()?;
            Some((version, path))
        })
        .max_by_key(|(version, _)| *version)
        .map(|(_, path)| path)
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    struct TempDir {
        path: PathBuf,
    }

    impl TempDir {
        fn new(prefix: &str) -> Self {
            let unique = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("system time should be valid")
                .as_nanos();
            let path = std::env::temp_dir().join(format!("{prefix}-{unique}"));
            fs::create_dir_all(&path).unwrap();
            Self { path }
        }

        fn path(&self) -> &Path {
            &self.path
        }
    }

    impl Drop for TempDir {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.path);
        }
    }

    #[test]
    fn test_discover_macos_office_font_paths_prefers_office_order() {
        let temp = TempDir::new("office-font-discovery-order");
        let apps = temp.path().join("Applications");
        let home = temp.path().join("home");

        fs::create_dir_all(apps.join("Microsoft PowerPoint.app/Contents/Resources/DFonts"))
            .unwrap();
        fs::create_dir_all(apps.join("Microsoft Word.app/Contents/Resources/DFonts")).unwrap();
        fs::create_dir_all(apps.join("wpsoffice.app/Contents/Resources/office6/fonts")).unwrap();
        fs::create_dir_all(
            home.join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts"),
        )
        .unwrap();
        fs::create_dir_all(
            home.join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/PreviewFont"),
        )
        .unwrap();

        let discovered = discover_macos_office_font_paths_from(&[apps], &home);
        let expected =
            vec![
                fs::canonicalize(
                    temp.path()
                        .join("Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"),
                )
                .unwrap(),
                fs::canonicalize(
                    temp.path()
                        .join("Applications/Microsoft Word.app/Contents/Resources/DFonts"),
                )
                .unwrap(),
                fs::canonicalize(temp.path().join(
                    "home/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts",
                ))
                .unwrap(),
                fs::canonicalize(temp.path().join(
                    "home/Library/Group Containers/UBF8T346G9.Office/FontCache/4/PreviewFont",
                ))
                .unwrap(),
                fs::canonicalize(
                    temp.path()
                        .join("Applications/wpsoffice.app/Contents/Resources/office6/fonts"),
                )
                .unwrap(),
            ];

        assert_eq!(discovered, expected);
    }

    #[test]
    fn test_discover_macos_office_font_paths_selects_highest_font_cache_version() {
        let temp = TempDir::new("office-font-discovery-version");
        let apps = temp.path().join("Applications");
        let home = temp.path().join("home");

        fs::create_dir_all(
            home.join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts"),
        )
        .unwrap();
        fs::create_dir_all(
            home.join("Library/Group Containers/UBF8T346G9.Office/FontCache/7/CloudFonts"),
        )
        .unwrap();
        fs::create_dir_all(
            home.join("Library/Group Containers/UBF8T346G9.Office/FontCache/7/PreviewFont"),
        )
        .unwrap();

        let discovered = discover_macos_office_font_paths_from(&[apps], &home);
        assert!(
            discovered
                .iter()
                .any(|path| path.ends_with("FontCache/7/CloudFonts")),
            "highest font cache version should be used"
        );
        assert!(
            discovered
                .iter()
                .all(|path| !path.ends_with("FontCache/4/CloudFonts")),
            "older font cache versions should be ignored"
        );
    }

    #[test]
    fn test_merge_prioritized_paths_keeps_first_occurrence() {
        let temp = TempDir::new("office-font-merge");
        let first = temp.path().join("first");
        let second = temp.path().join("second");
        fs::create_dir_all(&first).unwrap();
        fs::create_dir_all(&second).unwrap();

        let merged = merge_prioritized_paths(
            &canonicalize_existing_dirs(vec![first.clone(), second.clone()]),
            &canonicalize_existing_dirs(vec![second, first]),
        );

        assert_eq!(merged.len(), 2);
        assert!(merged[0].ends_with("first"));
        assert!(merged[1].ends_with("second"));
    }

    #[test]
    fn test_canonicalize_existing_dirs_skips_missing_paths() {
        let temp = TempDir::new("office-font-canonicalize");
        let existing = temp.path().join("existing");
        fs::create_dir_all(&existing).unwrap();
        let missing = temp.path().join("missing");

        let canonicalized =
            canonicalize_existing_dirs(vec![existing.clone(), missing, existing.clone()]);

        assert_eq!(canonicalized.len(), 1);
        assert_eq!(canonicalized[0], fs::canonicalize(existing).unwrap());
    }
}
