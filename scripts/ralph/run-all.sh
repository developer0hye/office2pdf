#!/bin/bash
# run-all.sh - Orchestrate Ralph across multiple phases with automatic PR/merge
#
# Usage:
#   ./scripts/ralph/run-all.sh              # Run all phases (2, 3, 4)
#   ./scripts/ralph/run-all.sh phase3       # Start from phase3 (skip completed)
#   ./scripts/ralph/run-all.sh phase2 25    # Run phase2 with 25 max iterations
#
# Flow per phase:
#   worktree setup → Ralph coding loop (per-story push/PR/CI/merge by Claude) → sync main → cleanup
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

# Track current worktree so we can clean up on interrupt
CURRENT_WORKTREE=""
CURRENT_BRANCH=""

on_interrupt() {
  echo ""
  log "Interrupted by user."
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

  # ── Step 4: Sync main ─────────────────────────────────────────
  # Ralph handles per-story push/PR/CI/merge internally.
  # We just need to sync main with whatever Ralph merged.

  cd "$PROJECT_ROOT"
  remove_untracked_conflicts

  if ! git pull 2>&1 | tee -a "$LOG_FILE"; then
    log "WARNING: git pull failed. Trying fetch + reset ..."
    git fetch origin main 2>/dev/null || true
    git reset --hard origin/main 2>/dev/null || true
  fi
  log "Main branch synced."

  # ── Step 5: Cleanup worktree and branch ────────────────────────

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
