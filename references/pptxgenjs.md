# PptxGenJS

- **URL:** https://github.com/gitbrent/PptxGenJS
- **Stars:** ~1.8k
- **Language:** TypeScript
- **License:** MIT

## What it does

Generates PPTX files and includes table auto-paging/layout heuristics.

## Why relevant

- `src/gen-tables.ts` computes row growth from parsed text lines, font size, and cell margins.
- It is a useful contrast to object-model libraries: practical layout code derives table height from content rather than trusting a stored row-height number alone.
- Helpful when reproducing PowerPoint-like table fitting in another backend.

## Key files to study

- `src/gen-tables.ts`

## When to consult

- Designing heuristics for PPTX table text fitting
- Checking how margins and line-height affect row growth in a PowerPoint-oriented renderer
