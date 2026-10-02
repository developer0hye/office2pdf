# openpyxl

- **URL:** https://github.com/ericgazoni/openpyxl
- **Language:** Python
- **License:** MIT

## What it does

Python reader and writer for Office Open XML spreadsheets.

## Why relevant

- `openpyxl/styles/colors.py::COLOR_INDEX` records the 56 default RGB values
  for indexed palette slots 8–63, used by legacy Excel color controls.
- The number format `[Color N]` selects that sequence at slot `N + 7`.

## Key file

- `openpyxl/styles/colors.py`

## When to consult

- Seeding or cross-checking default XLSX indexed color values.
- Comparing indexed palette behavior between OOXML implementations.
