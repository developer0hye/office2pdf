# PPTX Font Resolution and Text Style Inheritance

## Summary

This note documents a PPTX fidelity issue that initially looked like a layout bug
but was actually caused by incomplete font and text-style resolution.

The visible symptom was that text blocks on several slides, especially the
slide 4 `SKILLS` area, did not match PowerPoint:

- narrow labels wrapped differently
- title badges had different sizing and weight
- some text appeared visually shifted even when box coordinates were correct

The root cause was not hardcoded slide geometry. The real issue was that
`office2pdf` did not resolve PPT text styling the way PowerPoint does.

## Problem

PowerPoint text formatting is not determined only by direct run properties.
For a text box, effective styling can come from multiple OOXML layers:

1. direct run properties in `a:rPr`
2. paragraph properties in `a:pPr`
3. text body defaults in `a:lstStyle`
4. default paragraph properties in `a:defPPr`
5. level-specific paragraph properties in `a:lvlNpPr`
6. default run properties in `a:defRPr`
7. theme font references such as `+mj-lt` / `+mn-lt`

Before the fix, the PPTX parser mostly used direct `a:rPr` and direct
`a:latin` values on runs. That was incomplete.

As a result:

- runs with missing direct font settings lost their inherited family, size,
  color, and weight
- text boxes using `lstStyle` defaults rendered with weaker styling than
  PowerPoint
- line wrapping diverged because the effective font metrics were wrong

There was a second issue in fallback selection:

- fallback ordering preferred "font source rank" too aggressively
- this caused Office-managed fonts such as `Malgun Gothic` to outrank the
  metric-preferred substitute list for `Pretendard`
- on this machine, that produced worse wrapping than using
  `Apple SD Gothic Neo`

## Why PowerPoint Looked Correct

The affected deck does not rely only on theme fonts. It contains explicit
typeface information such as `Pretendard`, `Pretendard SemiBold`, and related
variants in slide XML.

PowerPoint applies:

- direct run formatting
- inherited `lstStyle` defaults
- script-aware typeface resolution across `latin`, `ea`, and `cs`
- local fallback behavior when the exact face is unavailable

`office2pdf` was missing part of that chain, so the generated PDF diverged
even though the original PPTX rendered correctly in PowerPoint.

## Fix

### 1. Parse text body style inheritance

Added a text-body default model in the PPTX parser so a text box can inherit
style from `a:lstStyle`.

Implementation:

- introduced `PptxTextBodyStyleDefaults`
- introduced `PptxTextLevelStyle`
- parse `a:lstStyle`, `a:defPPr`, `a:lvlNpPr`, and nested `a:defRPr`
- merge inherited paragraph and run styles before applying direct paragraph/run
  overrides
- apply typeface from `a:latin`, `a:ea`, and `a:cs`, not only direct `a:latin`

This makes text boxes behave much closer to PowerPoint's effective style model.

### 2. Preserve metric fallback order

Fallback selection for known families such as `Pretendard` now preserves the
substitution table's metric-compatible order first and only uses source rank as
a secondary tiebreaker.

That means:

- preferred substitute family order is stable
- Office-managed fonts no longer automatically jump ahead of better metric
  matches just because they were discovered from Office paths
- the real deck now falls back to `Apple SD Gothic Neo` instead of
  `Malgun Gothic`

## Files Involved

- `crates/office2pdf/src/parser/pptx.rs`
- `crates/office2pdf/src/render/font_subst.rs`
- `crates/office2pdf/src/render/typst_gen.rs`

## Verification

The fix was verified in three ways:

1. unit tests for `lstStyle` default run inheritance
2. unit tests for fallback ordering under a mixed Office/system font context
3. reconversion of the real PPTX deck and manual visual inspection

Observed result after the fix:

- slide 4 `SKILLS` labels no longer wrap as before
- badge/title styling is closer to PowerPoint
- regenerated PDF embeds `Apple SD Gothic Neo` variants instead of
  `Malgun Gothic` for `Pretendard` fallback

The remaining known gap is unrelated to text inheritance:

- slide 17 still warns for unsupported `hdphoto1.wdp`

## Key Lesson

When a PPTX rendering mismatch looks like a textbox position problem, do not
assume the coordinates are wrong first. In PowerPoint, incorrect font
resolution often becomes a layout problem:

- different font family
- different weight
- different inherited size
- different line spacing

Those differences change wrapping and visual alignment even if the text box
geometry is correct.
