# Bug-fix evidence

Visual evidence lives in an issue-numbered directory and is checked by the Visual
PR Contract job (`scripts/check_visual_pr.py`). There are two modes; a pull
request declares which one it uses in `Visual audit > Evidence mode`.

## Evidence mode: `fix`

A pull request that fixes a visual defect keeps all three images, a layout
report, and one strict render-cluster report per compared page:

```text
assets/bugfixes/issue-<number>/gt.jpg
assets/bugfixes/issue-<number>/before.jpg
assets/bugfixes/issue-<number>/after.jpg
assets/bugfixes/issue-<number>/layout-audit.json
assets/bugfixes/issue-<number>/render-clusters-page-<page>.json
# Only when the supplied GT differs from a verified native Office export:
assets/bugfixes/issue-<number>/reference-exporter-differences.json
assets/bugfixes/issue-<number>/native.jpg
```

Touching any image makes the gate require and validate the full trio. Every fix
pull request with a raster change changes `after.jpg` and `layout-audit.json`; `gt.jpg` and
`before.jpg` may remain from the defect evidence when they are still current.
Generate all evidence from the same input document, page, resolution, and
renderer.

For a text-layer-only fix, the honest before and after page renders can be
byte- and pixel-identical. Regenerate and verify the current `after.jpg` without
annotations; it need not appear in the diff when the render is unchanged. Set
`Text-layer-only: Yes` and `Pixel delta: 0` in the pull request's `Visual audit`,
change `layout-audit.json`, and use that report to demonstrate that the missing
or extra searchable text is resolved. The contract requires both `before.jpg`
and `after.jpg` and independently runs ImageMagick's exact decoded-pixel
comparison; the declaration cannot bypass a raster-changing fix.

Generate the machine-readable layout report from those same GT and after PDFs:

```sh
python3 scripts/compare_layout.py --json --audit --fine-shift 0.5 gt.pdf after.pdf \
  > assets/bugfixes/issue-<number>/layout-audit.json
```

The layout audit retains whole-line text matching but compares independent
cell/object anchors within a shared baseline. A separate text paint with at
least one em of empty space identifies a fragment boundary; a boundary on
either side is mapped to the same text offset on the other. This prevents a
fixed first cell from hiding later cells' movement or visibility differences.
Normal word spacing stays together; split/join topology matching separately
requires distant, independently painted objects. Closely packed text without such a boundary still needs the
pixel comparison and manual crop inspection.

Distant text objects in one row may have slightly reversed baseline order.
The split/join matcher accepts their left-to-right order when horizontal bounds
are disjoint and conservative ink bands overlap. Position and visibility checks
still apply to every recovered fragment; separate-row reorderings remain findings.
Repeated labels require equal occurrence counts and exactly one counterpart
with an overlapping ink band. Unequal counts and ambiguous rows remain findings. Partially joined cells are
recovered only when every fragment of each consumed line has a counterpart.


The command exits nonzero when material findings remain but still writes the
report. Change both `after.jpg` and `layout-audit.json` in a raster-changing fix
pull request; use the text-layer-only exception above when the current render is
byte-identical. Record the same `Fine-detail threshold` in the PR's `Visual
audit`, then mark each layout-audit category as `Pass` or list the open issues
that classify it. Those references must also appear in `Remaining:` deviation
rows. The gate rejects a claimed pass when the report contains a page-count
difference, missing/extra/reflowed text, changed wraps, a painted-text visibility
mismatch, a visible-fill occlusion, or a text or rectangle geometry shift above
the configured fine or large threshold.

### Composed fill coverage

Rectangle diagnostics retain raw matched dimensions and source-operation
indices. Every matched fill pair is checked after rectangular clipping and
paint-order composition, including pairs with equal raw bounds. A visible
coverage difference beyond the active geometry tolerance is a finding.
For a raw fill geometry finding, the audit requires a stricter coverage proof.
Only equivalent coverage makes that finding informational; a real difference
or unmodeled coverage keeps the geometry finding active. Per-edge coverage
separates an already-painted edge from an error elsewhere in the same fill.
The fine/coarse gate is unchanged; coincident trace coordinates use the existing
0.01pt epsilon, clamped to the active gate when it is finer. Each coordinate
cluster is bounded to that full span. Pairs without a raw finding use the active
geometry tolerance for coordinate normalization; unmodeled coverage is reported
but does not alone introduce a geometry finding.

Images, shading, nonrectangular paths/clips, translucent paint, soft masks, and
non-normal groups cannot establish flat-fill equivalence. A mask or tile marks
the page's fill coverage unmodeled because its later graphics-state effect is
not reconstructed. Later opaque fills may cover unknown image/path paint. Comparisons exceeding 100,000 partition cells also
remain unmodeled. Text and strokes have their own audits; this fill-layer proof
does not replace pixel inspection. Raw findings and unknown regions remain in
the JSON report even when other edges match.

### Verified reference-exporter differences

Do not change correct Office behavior merely to match a PDF made by another
exporter. When a native Word, PowerPoint, or Excel export proves that the
supplied GT is different, freshly change the issue-local
`reference-exporter-differences.json` with the visual-fix PR. It has this strict
shape:

```json
{
  "schema_version": 1,
  "source": {
    "url": "https://github.com/owner/repository/files/123/source.pptx",
    "sha256": "<64 lowercase hex characters>"
  },
  "reference_export": {
    "application": "LibreOffice Impress",
    "version": "26.2.5.2",
    "platform": "macOS 26.6.2",
    "pdf_sha256": "<64 lowercase hex characters>",
    "evidence_path": "assets/bugfixes/issue-123/gt.jpg",
    "evidence_sha256": "<64 lowercase hex characters>"
  },
  "native_export": {
    "application": "Microsoft PowerPoint",
    "version": "16.112.3",
    "platform": "macOS 26.6.2 build 25G83",
    "pdf_sha256": "<64 lowercase hex characters>",
    "evidence_path": "assets/bugfixes/issue-123/native.jpg",
    "evidence_sha256": "<64 lowercase hex characters>"
  },
  "verification_url": "https://github.com/owner/repository/issues/456#issuecomment-789",
  "differences": [
    {
      "id": "page-9-slide-number-visibility",
      "page": 9,
      "kind": "painted-text-visibility",
      "layout_finding": {
        "label": "9",
        "gt": "hidden",
        "out": "painted",
        "occurrence": 1
      }
    },
    {
      "id": "page-9-title-native-match",
      "page": 9,
      "kind": "render-clusters",
      "render_cluster_ids": ["p9-0123456789ab"]
    }
  ]
}
```

Both `gt.jpg` and `native.jpg` are full-page renders from the recorded PDFs and
follow the normal JPEG rules. The gate verifies the committed JPEG hashes and
requires structured source/reference/native provenance, SHA-256-shaped binary
hashes, and a stable GitHub comment URL. The reviewer still verifies the
uncommitted source/PDF hashes and comment content manually. Free-form notes do
not create a disposition.

Set these PR fields when the report is used:

```text
Reference exporter differences: assets/bugfixes/issue-123/reference-exporter-differences.json
Native: assets/bugfixes/issue-123/native.jpg
Layout audit text flow: ref:page-9-slide-number-visibility
```

Render `![Native](...)` with a stable image URL. In the deviation table, use
`Reference difference: ref:<id>` once for every difference in the report.
Layout reference differences currently cover only one exact occurrence of a
`painted-text-visibility` finding. Missing/extra text, wrap/reflow, geometry,
fill, or shift findings still need an open issue. A category containing both a
verified visibility difference and another finding must also name that issue.

Generate the cluster IDs once without strict mode, inspect every cluster in the
full-resolution diff and matched crops, then create a disposition file whose
groups enumerate those exact IDs:

```json
{
  "schema_version": 1,
  "renderer_observations": [
    {
      "class": "shape-edge-antialiasing",
      "bbox_pt": {"x": 220, "y": 330, "width": 380, "height": 390},
      "note": "Inspected hairline fragments remain below the material-cluster floor."
    }
  ],
  "groups": [
    {
      "kind": "accepted-rendering",
      "class": "glyph-edge-rasterization",
      "cluster_ids": ["p1-0123456789ab"],
      "note": "Glyph outlines only in the inspected crop."
    },
    {
      "kind": "issue",
      "issue": "#123",
      "cluster_ids": ["p1-fedcba987654"]
    },
    {
      "kind": "reference-exporter-difference",
      "difference_id": "page-9-title-native-match",
      "cluster_ids": ["p9-0123456789ab"]
    }
  ]
}
```

Accepted renderer-only classes are `glyph-edge-rasterization`,
`photo-resampling`, `gradient-rasterization`, and
`shape-edge-antialiasing`. Page-, region-, and bbox-wide selectors are rejected:
every material cluster must be named, so a new cluster cannot inherit an old
approval. `renderer_observations` may record a bounded crop whose inspected
renderer-only fragments stay below the material-cluster floor. They never
disposition a cluster, so a new material cluster inside that bbox still fails.
Rerun each page in strict mode and commit its passing report:

```sh
python3 scripts/compare_render.py gt.pdf after.pdf --page <page> --dpi 300 \
  --fine-shift 0.5 --artifacts-dir target/audit/page-<page> \
  --cluster-dispositions target/audit/page-<page>-dispositions.json \
  --cluster-report assets/bugfixes/issue-<number>/render-clusters-page-<page>.json \
  --strict-clusters
```

When a disposition group names a verified reference-exporter difference, add
`--reference-differences assets/bugfixes/issue-<number>/reference-exporter-differences.json`.

The report contains every cluster's stable ID, page-space bbox, area, region,
and disposition, plus validated bounded renderer observations.
`--strict-clusters` exits nonzero for missing, stale, duplicate, or invalid IDs,
and when the ImageMagick census is unavailable. List every compared page's
report in `Visual audit > Render cluster reports`; the PR gate requires all of
them to be freshly changed and passing. A reference-exporter group must equal
its manifest entry's complete cluster-ID set, so a new or unrelated cluster
remains undispositioned.

## Evidence mode: `defect`

A pull request that files a visual defect commits a single side-by-side image —
ground truth on the left, office2pdf output at filing time on the right:

```text
assets/bugfixes/issue-<number>/compare.jpg
```

Both panels come from the same page and resolution, so `compare.jpg` must be an
even number of pixels wide for the halves to split evenly.

## JPEG rules the gate enforces

Every committed image must be:

- **progressive** — baseline JPEG is rejected;
- **at least 150 DPI** in the JFIF density header, and rendered at that DPI
  rather than upscaled afterwards;
- **free of metadata** — APP1 (Exif/XMP), APP2 (ICC), APP13 (Photoshop/IPTC), and
  COM markers must all be absent.

Use quality 86 and preserve the source pixel dimensions. Set the density *after*
`-strip`, because `-strip` also drops the JFIF header that carries it:

```sh
magick page.png -strip -density 150 -units PixelsPerInch \
  -interlace Plane -quality 86 assets/bugfixes/issue-<number>/after.jpg
```

Validate locally before pushing:

```sh
python3 scripts/check_visual_pr.py --event event.json --base main --head HEAD \
  --repository developer0hye/office2pdf --root .
```

## Non-image files

Markdown and text files under `assets/bugfixes/` — this README included — are
bookkeeping, not evidence. The gate skips them, so a pull request that changes
only such a file may check `No rendered PDF change`. A fix-mode
`layout-audit.json`, `render-clusters-page-<page>.json`, and optional
`reference-exporter-differences.json` are validated separately and must be
updated with the visual evidence in their issue directory.

Anything else in the directory is treated as evidence and must match one of the
layouts above. The nested `issue-<number>/audit/<case>/` layout used by the early
tracking issues (#213, #214) predates the gate and is not accepted for new
evidence.
