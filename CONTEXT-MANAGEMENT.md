# Context Management

Project memory is stored in `.claude/memory/` and **committed to git** so that context is shared across machines and developers.

## File Structure

```
.claude/memory/
├── INDEX.md                                        # Always read first. One-line summary per file.
├── decision--<short-description>.md                # Technical decisions and rationale
├── workaround--<short-description>.md              # Non-obvious workarounds for library/tool limitations
├── bug--<short-description>.md                     # Bugs that took significant effort to diagnose
└── discovery--<short-description>.md               # Surprising findings about specs, libraries, or behavior
```

## File Naming

- Prefix tag: `decision--`, `workaround--`, `bug--`, `discovery--`
- Suffix: lowercase, hyphen-separated, descriptive keywords
- The filename alone should tell you whether to open the file. Be specific.
  - Good: `decision--pptx-parser-zip-quickxml-over-ppt-rs.md`
  - Bad: `decision--pptx-change.md`

## File Template

Every memory file must start with:

```markdown
# <Title>

> TL;DR: <one or two sentences>

- **Type**: decision | workaround | bug | discovery
- **Date**: YYYY-MM
- **Status**: applied | deprecated | open | superseded-by:<filename>
- **Related**: <PRD sections, file paths, crate names>

## Context
<Why did this come up?>

## Outcome
<What was decided/done/found?>

## Impact
<What else is affected? PRD updates needed? Other files to watch?>
```

## INDEX.md

Maintained in `.claude/memory/INDEX.md`. Updated every time a memory file is added, modified, or deprecated.

```markdown
| File | Status | Related | TL;DR |
|------|--------|---------|-------|
| decision--pptx-parser-zip-quickxml-over-ppt-rs.md | applied | parser/pptx.rs, PRD §5.4 | ppt-rs slide master/theme 미지원 → zip+quick-xml 직접 OOXML 파싱 |
```

INDEX.md alone should be sufficient to decide whether to open a file. Write TL;DR with searchable keywords.

Read INDEX.md first, then open only the relevant files.

## When to WRITE

### Must record
- PRD에 명시된 것과 다르게 구현했을 때
- 외부 크레이트를 교체·제거했을 때
- 2시간 이상 디버깅한 문제를 해결했을 때
- "다음에도 이거 또 헤맬 것 같다"고 느낄 때
- 아키텍처나 설계 방향이 바뀌었을 때

### Do NOT record
- 단순 리팩토링 (rename 등)
- 일반적인 Rust/언어 패턴
- 한 번만 쓰고 끝인 임시 해결책
- 커밋 메시지만으로 충분히 설명되는 변경

### Write procedure
1. Create `.claude/memory/<tag>--<short-description>.md` using the template above
2. Update `.claude/memory/INDEX.md` with a new row
3. If a previous entry is superseded, update its Status to `superseded-by:<new-filename>`

## When to READ

### READ procedure
1. If you know what you're looking for (specific module, crate, feature):
   → Grep INDEX.md for keywords first, read matching entries only
2. If starting fresh or context is unclear:
   → Read full INDEX.md, but Open section only is sufficient for most tasks

### Always read INDEX.md
- Before starting any new task
- When resuming work after a context switch

### Search and read specific files when
- Modifying a file/module mentioned in a memory file's `Related` field
- Introducing or evaluating an external crate → check `decision--*` files
- Hitting an unexpected error or behavior → check `bug--*`, `workaround--*` files
- Working on a feature that diverges from PRD → check `decision--*` files

## Size Limits

Memory files are notes, not documentation. Keep them short.

- **TL;DR**: 2 sentences max
- **Context**: 5 sentences max
- **Outcome**: 5 sentences max
- **Impact**: 3 sentences max
- **Total file**: 30 lines max
- **No code blocks.** Describe behavior in prose. Link to source files instead of copying code.

If it needs more space, the scope is too broad — split into multiple files.

## Maintenance

- When INDEX.md exceeds 50 entries: review and archive deprecated entries
- When opening a memory file and its content is outdated: update or mark deprecated immediately
- Archived files go to `.claude/memory/archive/` (still git-tracked, but removed from INDEX.md)

## Worktree & Concurrency

Multiple worktrees (or Claude Code instances) may run in parallel. To avoid merge conflicts:

- **READ from any branch.** Memory files in `.claude/memory/` are always safe to read.
- **WRITE only on `main`.** After a PR is merged and you return to the main repo, write memory files then.
- **In a worktree**, if you encounter something worth recording, note it in the commit message or PR description. Write the memory file after merge, from main.
- **Never modify `.claude/memory/` in a feature branch.** This prevents INDEX.md conflicts across concurrent worktrees.

## Git Policy

- `.claude/memory/` **must be committed to the repository** — do not add it to `.gitignore`
- `.claude/settings.local.json` is machine-specific and **should be gitignored**
