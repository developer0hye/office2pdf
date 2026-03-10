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

    fs::create_dir_all(apps.join("Microsoft PowerPoint.app/Contents/Resources/DFonts")).unwrap();
    fs::create_dir_all(apps.join("Microsoft Word.app/Contents/Resources/DFonts")).unwrap();
    fs::create_dir_all(
        home.join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts"),
    )
    .unwrap();
    fs::create_dir_all(
        home.join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/PreviewFont"),
    )
    .unwrap();

    let discovered = discover_macos_office_font_paths_from(&[apps], &home);
    let expected = vec![
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
        fs::canonicalize(
            temp.path()
                .join("home/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts"),
        )
        .unwrap(),
        fs::canonicalize(
            temp.path()
                .join("home/Library/Group Containers/UBF8T346G9.Office/FontCache/4/PreviewFont"),
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
