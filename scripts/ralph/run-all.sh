#!/bin/bash
# run-all.sh - Orchestrate Ralph across multiple phases with automatic PR/merge
#
# Usage:
#   ./scripts/ralph/run-all.sh              # Run all phases (2, 3, 4)
#   ./scripts/ralph/run-all.sh phase3       # Start from phase3 (skip completed)
#   ./scripts/ralph/run-all.sh phase2 25    # Run phase2 with 25 max iterations
#
# Flow per phase:
#   worktree setup → Ralph coding loop → push → PR create → CI wait → merge → cleanup
#
# Resilience: failures in one phase are logged and skipped. The script always
# attempts ALL remaining phases instead of stopping at the first error.
#
# Requirements: bash 3.2+, git, gh (authenticated), jq, claude
# Compatible with macOS default bash (3.2) — no bash 4+ features used.

set -eo pipefail

# ── Configuration ───────────────────────────────────────────────────

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
LOG_FILE="$PROJECT_ROOT/ralph.log"

# Phase definitions — parallel indexed arrays (bash 3.2 compatible)
ALL_PHASES=(    phase2                phase3                  phase4               )
ALL_BRANCHES=(  "ralph/phase2-formatting" "ralph/phase3-advanced"  "ralph/phase4-polish" )
ALL_DESCS=(
  "Phase 2: P1 Features - Formatting, Styles, Font Fallback"
  "Phase 3: P2 Features - Hyperlinks, Footnotes, TOC, PDF/A, Batch"
  "Phase 4: P3 Features - Charts, Equations, SmartArt, Polish"
)
ALL_ITERS=(     20                    18                      15                   )

# ── Argument parsing ────────────────────────────────────────────────

START_PHASE="${1:-phase2}"
OVERRIDE_ITERS="${2:-}"

# Validate start phase
start_index=-1
for i in 0 1 2; do
  if [ "${ALL_PHASES[$i]}" = "$START_PHASE" ]; then
    start_index=$i
    break
  fi
done
if [ "$start_index" -eq -1 ]; then
  echo "Error: Unknown phase '$START_PHASE'. Valid: phase2, phase3, phase4"
  exit 1
fi

# ── Helper functions ────────────────────────────────────────────────

log() {
  local ts
  ts="$(date '+%Y-%m-%d %H:%M:%S')"
  echo "[$ts] $1"
  echo "[$ts] $1" >> "$LOG_FILE"
}

die() {
  log "FATAL: $1"
  exit 1
}

# Skip current phase: log reason, cleanup worktree, record failure, continue
# Usage: skip_phase "reason" && continue
skip_phase() {
  log "SKIPPING $PHASE: $1"
  cleanup_worktree "$WORKTREE_DIR" "$BRANCH"
  failed_phases="$failed_phases $PHASE"
}

# ── Preflight checks ───────────────────────────────────────────────

preflight() {
  local missing=""
  command -v git   >/dev/null 2>&1 || missing="$missing git"
  command -v gh    >/dev/null 2>&1 || missing="$missing gh"
  command -v jq    >/dev/null 2>&1 || missing="$missing jq"
  command -v claude >/dev/null 2>&1 || missing="$missing claude"

  if [ -n "$missing" ]; then
    die "Missing required tools:$missing"
  fi

  if ! gh auth status >/dev/null 2>&1; then
    die "GitHub CLI not authenticated. Run: gh auth login"
  fi

  if ! git -C "$PROJECT_ROOT" rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    die "Not a git repository: $PROJECT_ROOT"
  fi
}

# ── Cleanup trap ────────────────────────────────────────────────────

# Track current worktree and background PID so we can clean up on interrupt
CURRENT_WORKTREE=""
CURRENT_BRANCH=""
CI_WATCH_PID=""

on_interrupt() {
  echo ""
  log "Interrupted by user."
  # Kill background CI watch process if running
  if [ -n "$CI_WATCH_PID" ] && kill -0 "$CI_WATCH_PID" 2>/dev/null; then
    kill "$CI_WATCH_PID" 2>/dev/null || true
    wait "$CI_WATCH_PID" 2>/dev/null || true
    log "Killed background CI watch process."
  fi
  if [ -n "$CURRENT_WORKTREE" ]; then
    log "Worktree may be left at: $CURRENT_WORKTREE"
    log "Branch may be left: $CURRENT_BRANCH"
    log "To clean up manually:"
    log "  rm -rf $CURRENT_WORKTREE"
    log "  git worktree prune"
    log "  git branch -D $CURRENT_BRANCH"
  fi
  exit 130
}
trap on_interrupt INT TERM

# ── Remove untracked conflicting files before git pull ──────────────

remove_untracked_conflicts() {
  cd "$PROJECT_ROOT"
  local f
  for f in scripts/ralph/prd.json scripts/ralph/progress.txt; do
    if [ -f "$f" ]; then
      # Check if the file is tracked by git; if not, remove it
      if ! git ls-files --error-unmatch "$f" >/dev/null 2>&1; then
        rm -f "$f"
        log "  Removed untracked conflict: $f"
      fi
    fi
  done
}

# ── Clean up a worktree + branch (best-effort) ─────────────────────

cleanup_worktree() {
  local wt_dir="$1"
  local branch="$2"

  cd "$PROJECT_ROOT"

  # Remove worktree directory
  if [ -d "$wt_dir" ]; then
    rm -rf "$wt_dir"
  fi
  git worktree prune 2>/dev/null || true

  # Delete local branch
  git branch -D "$branch" 2>/dev/null || true

  # Delete remote branch
  if git ls-remote --heads origin "$branch" 2>/dev/null | grep -q "$branch"; then
    git push origin --delete "$branch" 2>/dev/null || true
  fi

  CURRENT_WORKTREE=""
  CURRENT_BRANCH=""
}

# ════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════

preflight
cd "$PROJECT_ROOT"

log "============================================================"
log "  Ralph All-Phases Runner"
log "  Project: $PROJECT_ROOT"
log "  Starting from: $START_PHASE"
log "============================================================"
echo ""

completed_count=0
failed_phases=""

for idx in $(seq "$start_index" 2); do
  PHASE="${ALL_PHASES[$idx]}"
  BRANCH="${ALL_BRANCHES[$idx]}"
  DESC="${ALL_DESCS[$idx]}"
  MAX_ITERS="${OVERRIDE_ITERS:-${ALL_ITERS[$idx]}}"
  WORKTREE_NAME="office2pdf-${BRANCH//\//-}"
  WORKTREE_DIR="$PROJECT_ROOT/../$WORKTREE_NAME"

  CURRENT_WORKTREE="$WORKTREE_DIR"
  CURRENT_BRANCH="$BRANCH"

  log "============================================================"
  log "  $DESC"
  log "  Branch: $BRANCH  |  Max iterations: $MAX_ITERS"
  log "============================================================"

  # ── Step 1: Prepare worktree ────────────────────────────────────

  cd "$PROJECT_ROOT"

  # Clean up any leftovers from a previous interrupted run
  if [ -d "$WORKTREE_DIR" ]; then
    log "Removing leftover worktree: $WORKTREE_DIR"
    rm -rf "$WORKTREE_DIR"
    git worktree prune 2>/dev/null || true
  fi

  if git show-ref --verify --quiet "refs/heads/$BRANCH" 2>/dev/null; then
    log "Removing leftover local branch: $BRANCH"
    git branch -D "$BRANCH" 2>/dev/null || true
  fi

  if git ls-remote --heads origin "$BRANCH" 2>/dev/null | grep -q "$BRANCH"; then
    log "Removing leftover remote branch: $BRANCH"
    git push origin --delete "$BRANCH" 2>/dev/null || true
  fi

  log "Creating worktree at $WORKTREE_DIR ..."
  if ! git worktree add "$WORKTREE_DIR" -b "$BRANCH" 2>&1 | tee -a "$LOG_FILE"; then
    skip_phase "Failed to create worktree" && continue
  fi

  # ── Step 2: Set up phase PRD and progress ───────────────────────

  cd "$WORKTREE_DIR"

  cp "$SCRIPT_DIR/phases/${PHASE}.json" scripts/ralph/prd.json

  # Preserve Codebase Patterns section from previous progress (if any)
  PATTERNS_TMP="$(mktemp)"
  if [ -f scripts/ralph/progress.txt ] && grep -q "^## Codebase Patterns" scripts/ralph/progress.txt; then
    sed -n '/^## Codebase Patterns/,/^---$/p' scripts/ralph/progress.txt > "$PATTERNS_TMP"
  fi

  {
    if [ -s "$PATTERNS_TMP" ]; then
      cat "$PATTERNS_TMP"
      echo ""
    fi
    echo "# Ralph Progress Log - $DESC"
    echo "Started: $(date)"
    echo "---"
  } > scripts/ralph/progress.txt
  rm -f "$PATTERNS_TMP"

  git add scripts/ralph/prd.json scripts/ralph/progress.txt
  if ! git commit -s -m "docs: set up ${PHASE} user stories for Ralph" 2>&1 | tee -a "$LOG_FILE"; then
    skip_phase "Failed to commit phase setup" && continue
  fi

  # ── Step 3: Run Ralph ───────────────────────────────────────────

  log "Starting Ralph for $PHASE ($MAX_ITERS iterations max) ..."

  set +e  # Don't exit on Ralph failure
  "$WORKTREE_DIR/scripts/ralph/ralph.sh" "$MAX_ITERS" 2>&1 | tee -a "$LOG_FILE"
  set -e

  # Ensure we're back in the worktree
  cd "$WORKTREE_DIR"

  # Report story progress (|| echo 0: protect against corrupted prd.json)
  total=$(jq '[.userStories[]] | length' scripts/ralph/prd.json 2>/dev/null || echo 0)
  passed=$(jq '[.userStories[] | select(.passes == true)] | length' scripts/ralph/prd.json 2>/dev/null || echo 0)
  incomplete=$((total - passed))

  if [ "$incomplete" -gt 0 ]; then
    log "WARNING: $passed/$total stories complete ($incomplete remaining in $PHASE)"
    log "Proceeding with partial progress ..."
  else
    log "All $total stories complete for $PHASE!"
  fi

  # Skip if Ralph made zero implementation commits
  commit_count=$(git rev-list --count "main..HEAD" 2>/dev/null || echo 0)
  if [ "$commit_count" -le 1 ]; then
    skip_phase "Ralph produced 0 implementation commits" && continue
  fi

  # ── Step 4: Push to remote ──────────────────────────────────────

  log "Pushing $BRANCH to origin ..."
  if ! git push -u origin "$BRANCH" 2>&1 | tee -a "$LOG_FILE"; then
    log "Push failed. Retrying in 10s ..."
    sleep 10
    if ! git push -u origin "$BRANCH" 2>&1 | tee -a "$LOG_FILE"; then
      skip_phase "Failed to push branch after retry" && continue
    fi
  fi

  # ── Step 5: Create PR ───────────────────────────────────────────

  commit_log=$(git log --oneline main..HEAD 2>/dev/null | head -30 || echo "(unable to read)")
  story_list=$(jq -r '.userStories[] | "- [" + (if .passes then "x" else " " end) + "] **" + .id + "**: " + .title' scripts/ralph/prd.json 2>/dev/null || echo "- (unable to read)")

  pr_title="feat: ${DESC}"
  # Truncate to 70 chars
  if [ ${#pr_title} -gt 70 ]; then
    pr_title="${pr_title:0:67}..."
  fi

  # Write PR body to a temp file (avoids heredoc-in-subshell escaping issues)
  pr_body_file="$(mktemp)"
  cat > "$pr_body_file" <<ENDOFBODY
## Summary

${DESC} — automated implementation by Ralph agent.

**Progress: ${passed} / ${total} user stories completed.**

### Commits

\`\`\`
${commit_log}
\`\`\`

### User stories

${story_list}

## Test plan

- [ ] \`cargo test --workspace\` passes
- [ ] \`cargo clippy --workspace -- -D warnings\` passes
- [ ] \`cargo fmt --all -- --check\` passes
- [ ] CI checks pass on all platforms

Generated with [Claude Code](https://claude.com/claude-code) via Ralph
ENDOFBODY

  log "Creating PR ..."
  pr_url=""
  if ! pr_url=$(gh pr create --title "$pr_title" --body-file "$pr_body_file" --base main 2>&1); then
    log "ERROR: gh pr create failed: $pr_url"
    rm -f "$pr_body_file"
    skip_phase "Failed to create PR" && continue
  fi
  rm -f "$pr_body_file"

  pr_num=$(echo "$pr_url" | grep -oE '[0-9]+$' || echo "")
  if [ -z "$pr_num" ]; then
    log "ERROR: Could not extract PR number from: $pr_url"
    skip_phase "Failed to extract PR number" && continue
  fi
  log "Created PR #$pr_num: $pr_url"

  # ── Step 6: Wait for CI (with auto-fix retry) ───────────────────

  ci_passed=false

  for ci_attempt in 1 2 3; do
    # Give GitHub Actions time to register the workflow run
    log "Waiting 30s for CI to register (attempt $ci_attempt/3) ..."
    sleep 30

    # Ensure CI checks actually exist before watching
    # (gh pr checks --watch can return immediately if no checks are registered yet)
    ci_check_count=0
    for wait_round in 1 2 3 4 5 6; do
      ci_check_count=$(gh pr checks "$pr_num" 2>&1 | grep -cE "(pass|fail|pending)" || true)
      if [ "$ci_check_count" -ge 3 ]; then
        break
      fi
      log "  Only $ci_check_count checks registered, waiting 10s ..."
      sleep 10
    done

    log "Watching CI checks on PR #$pr_num ($ci_check_count checks registered) ..."

    # Watch with a 30-minute timeout (background process + timer)
    CI_WATCH_PID=""
    ci_timed_out=false
    ci_watch_exit=0

    gh pr checks "$pr_num" --watch > /tmp/ralph-ci-watch-$$.log 2>&1 &
    CI_WATCH_PID=$!

    elapsed=0
    while kill -0 "$CI_WATCH_PID" 2>/dev/null; do
      sleep 10
      elapsed=$((elapsed + 10))
      if [ "$elapsed" -ge 1800 ]; then
        log "CI watch timed out after 30 minutes."
        kill "$CI_WATCH_PID" 2>/dev/null || true
        wait "$CI_WATCH_PID" 2>/dev/null || true
        ci_timed_out=true
        break
      fi
    done

    if [ "$ci_timed_out" = false ]; then
      wait "$CI_WATCH_PID" 2>/dev/null
      ci_watch_exit=$?
    fi
    CI_WATCH_PID=""

    cat /tmp/ralph-ci-watch-$$.log >> "$LOG_FILE" 2>/dev/null
    cat /tmp/ralph-ci-watch-$$.log 2>/dev/null
    rm -f /tmp/ralph-ci-watch-$$.log

    if [ "$ci_timed_out" = true ]; then
      log "ERROR: CI timed out for PR #$pr_num."
      break  # exit CI loop — will try merge anyway below
    fi

    if [ "$ci_watch_exit" -eq 0 ]; then
      ci_passed=true
      break
    fi

    # CI failed — identify what failed
    log "CI checks failed (attempt $ci_attempt). Checking what went wrong ..."
    ci_output=$(gh pr checks "$pr_num" 2>&1 || true)
    echo "$ci_output" >> "$LOG_FILE"

    fmt_failed=false
    clippy_failed=false
    if echo "$ci_output" | grep -i "format" | grep -q "fail"; then
      fmt_failed=true
    fi
    if echo "$ci_output" | grep -i "clippy" | grep -q "fail"; then
      clippy_failed=true
    fi

    # Attempt auto-fix only on attempts 1 and 2
    if [ "$ci_attempt" -lt 3 ]; then
      cd "$WORKTREE_DIR"
      fixed_something=false

      if [ "$fmt_failed" = true ]; then
        log "  Auto-fixing: running cargo fmt ..."
        if cargo fmt --all 2>/dev/null; then
          if ! git diff --quiet 2>/dev/null; then
            git add -u
            git commit -s -m "style: auto-fix formatting for CI" 2>/dev/null && fixed_something=true
          fi
        fi
      fi

      if [ "$clippy_failed" = true ]; then
        log "  Auto-fixing: running claude to fix clippy warnings ..."
        fix_prompt="Run 'cargo clippy --workspace --all-targets -- -D warnings 2>&1' and fix ALL clippy errors/warnings. Then run 'cargo fmt --all'. Then 'cargo test --workspace'. Commit all fixes with: git commit -s -m 'fix: resolve clippy warnings for CI'"
        if claude --dangerously-skip-permissions --print -p "$fix_prompt" > /tmp/ralph-clippy-fix-$$.log 2>&1; then
          fixed_something=true
        fi
        cat /tmp/ralph-clippy-fix-$$.log >> "$LOG_FILE" 2>/dev/null
        rm -f /tmp/ralph-clippy-fix-$$.log
      fi

      if [ "$fixed_something" = true ]; then
        log "  Pushing fixes ..."
        if git push 2>&1 | tee -a "$LOG_FILE"; then
          # Loop back to wait for new CI run
          continue
        else
          log "  Push failed after auto-fix. Retrying push ..."
          sleep 5
          if git push 2>&1 | tee -a "$LOG_FILE"; then
            continue
          else
            log "  Push still failing. Moving on."
          fi
        fi
      else
        log "  No auto-fix available for this failure."
      fi
    fi

    # This attempt failed
    log "CI attempt $ci_attempt failed for PR #$pr_num."
    # Don't break 2 — let the loop continue to try more attempts,
    # or fall through to merge-anyway below
  done

  # ── Step 7: Final diff review (logged) ──────────────────────────

  if [ "$ci_passed" = true ]; then
    log "CI checks passed for PR #$pr_num!"
  else
    log "WARNING: CI did not pass for $PHASE. Will attempt merge anyway ..."
  fi

  log "Changed files in PR #$pr_num:"
  gh pr diff "$pr_num" --name-only 2>&1 | tee -a "$LOG_FILE" || true

  # ── Step 8: Merge PR ────────────────────────────────────────────

  log "Merging PR #$pr_num ..."
  if ! gh pr merge "$pr_num" --merge 2>&1 | tee -a "$LOG_FILE"; then
    log "Merge with --merge failed. Trying --squash ..."
    if ! gh pr merge "$pr_num" --squash 2>&1 | tee -a "$LOG_FILE"; then
      log "ERROR: All merge strategies failed for PR #$pr_num."
      log "  PR left open: $pr_url"
      skip_phase "Merge failed" && continue
    fi
  fi
  log "PR #$pr_num merged!"

  # ── Step 9: Sync main ───────────────────────────────────────────

  cd "$PROJECT_ROOT"
  remove_untracked_conflicts

  if ! git pull 2>&1 | tee -a "$LOG_FILE"; then
    log "WARNING: git pull failed. Trying fetch + reset ..."
    git fetch origin main 2>/dev/null || true
    git reset --hard origin/main 2>/dev/null || true
  fi
  log "Main branch synced."

  # ── Step 10: Cleanup worktree and branch ────────────────────────

  log "Cleaning up worktree and branch ..."
  cleanup_worktree "$WORKTREE_DIR" "$BRANCH"

  completed_count=$((completed_count + 1))
  log ">>> $PHASE DONE ($passed/$total stories) <<<"
  echo ""
done

# ── Summary ─────────────────────────────────────────────────────────

echo ""
log "============================================================"
log "  Completed phases: $completed_count / ${#ALL_PHASES[@]}"

if [ -n "$failed_phases" ]; then
  log "  Failed phases:$failed_phases"
  log "  To resume a failed phase: ./scripts/ralph/run-all.sh <phase>"
  log "============================================================"
  # Exit 0 if at least 1 phase succeeded, 1 if ALL failed
  if [ "$completed_count" -gt 0 ]; then
    exit 0
  else
    exit 1
  fi
else
  log "  All phases complete!"
  log "============================================================"
  exit 0
fi
