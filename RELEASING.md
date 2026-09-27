# Release Manual

This is the canonical end-to-end release path. A release is complete only when the version PR is merged, the GitHub Release exists at that merge commit, both crates are visible on crates.io, and all six binary archives are attached.

## Why releases take time

The protected version PR must pass CI before merge. The release workflow then publishes the library and CLI in dependency order while six platform builds run in parallel. A healthy run normally spends most of its 10-15 minutes compiling those binaries; v0.6.3 took about 14 minutes. Repeated local builds, stale authentication, low disk space, or manually creating the tag add avoidable delay.

## 1. Preflight

Run these checks before editing anything:

```bash
git status --short --branch
git fetch origin
git rev-parse HEAD origin/main
git describe --tags --abbrev=0
df -h .
du -sh target 2>/dev/null || true
GH_TOKEN='' gh auth status -h github.com
GH_TOKEN='' gh api user --jq .login
```

Required state:

- The main worktree is clean and `HEAD` equals `origin/main`.
- The GitHub login is exactly `developer0hye`. `GH_TOKEN=''` intentionally ignores any stale token exported by the shell and uses the keyring login.
- A cold workspace test has at least 16 GiB free. A warm shared `target/` may use less, but never start with under 4 GiB free. If a linker reports `No space left on device`, run `cargo clean --target-dir <main-worktree>/target`, confirm the recovered space, and retry once. Reuse the main worktree's `target/` from the release worktree; do not start a second full dependency build.
- The proposed patch version does not already exist in GitHub Releases or crates.io.

For a persistent authentication repair across terminals, remove any `export GH_TOKEN=...` entry reported by `rg -l 'GH_TOKEN' ~/.zshenv ~/.zprofile ~/.zshrc ~/.config`, then run `unset GH_TOKEN`, `gh auth switch -h github.com -u developer0hye`, and `gh auth refresh -h github.com -s repo,workflow`. If the variable was injected through macOS launch services, also run `launchctl unsetenv GH_TOKEN`. Keep credentials in the keyring, never in repository or shell configuration files.

Use a release worktree and branch named `chore/publish-<version>`.

## 2. Version PR

Update every occurrence below to the same version:

- `crates/office2pdf/Cargo.toml`: package version
- `crates/office2pdf-cli/Cargo.toml`: package version
- `crates/office2pdf-cli/Cargo.toml`: `office2pdf` dependency requirement

Add release-process changes to this same PR only when they are required for the release. Verify alignment:

```bash
cargo metadata --offline --locked --no-deps --format-version 1 \
  | jq -r '.packages[] | select(.name == "office2pdf" or .name == "office2pdf-cli") | [.name, .version, ([.dependencies[] | select(.name == "office2pdf") | .req] | first // "-")] | @tsv'
cargo test --offline --locked --workspace
git diff --check
```

Commit with sign-off, push, and merge through the standard PR procedure. Do not create a release from the branch. After merge, synchronize the main worktree and record its full 40-character SHA.

## 3. One-dispatch release

From the synchronized `main`, run exactly one release dispatch:

```bash
GH_TOKEN='' gh workflow run release.yml \
  --repo developer0hye/office2pdf \
  --ref main \
  -f tag=v<version>
```

The workflow validates the tag against both Cargo packages and the CLI dependency, creates the tag and GitHub Release at the exact dispatched `main` SHA, generates release notes and contributors, publishes both crates, builds six platform archives, and uploads them. Normal releases must not use a separate manual `gh release create` command.

Find and watch the workflow run created by the dispatch:

```bash
GH_TOKEN='' gh run list --repo developer0hye/office2pdf \
  --workflow release.yml --event workflow_dispatch --limit 1
GH_TOKEN='' gh run watch <run-id> --repo developer0hye/office2pdf --exit-status
```

The preparation step is idempotent. Re-dispatching the same tag safely reuses an existing matching tag/Release, treats already-published crates as success, and overwrites duplicate assets. It must fail before publication when versions or refs do not match.

## 4. Completion checks

Do not report completion until every check below passes:

```bash
GH_TOKEN='' gh release view v<version> --repo developer0hye/office2pdf \
  --json tagName,targetCommitish,publishedAt,assets,url
GH_TOKEN='' gh run view <run-id> --repo developer0hye/office2pdf \
  --json status,conclusion,jobs,url
curl -sS https://crates.io/api/v1/crates/office2pdf/<version> | jq -r '.version.num, .version.yanked'
curl -sS https://crates.io/api/v1/crates/office2pdf-cli/<version> | jq -r '.version.num, .version.yanked'
```

Also verify:

- The release tag resolves to the recorded merge SHA.
- The release body contains the complete change summary and a Contributors section.
- Exactly six archives exist: Linux GNU x86_64, Linux musl x86_64, Linux GNU aarch64, macOS arm64, macOS x86_64, and Windows x86_64.
- Both crates report the requested version with `yanked: false`.

Only then remove the worktree and delete the merged local and remote branch.

## Recovery

- Authentication failure: rerun with `GH_TOKEN=''` and verify `gh api user` returns `developer0hye`.
- Low disk space: inspect `target/` first and reuse or clean only known build outputs; never delete fixtures or user files.
- Failed version validation: fix the version PR. Never move an existing public tag to a different commit.
- Failed publish/build job: inspect that job, fix the cause if needed, then re-dispatch the same tag. The workflow is designed for safe recovery.

