## File submission policy

Before attaching or committing files, read the [submission policy](https://github.com/developer0hye/office2pdf/blob/main/CONTRIBUTING.md).
Complete a security review and obtain your organization's internal approval for
public sharing where applicable. You are responsible for lawful disclosure;
the project and maintainers disclaim related liability to the extent permitted
by applicable law.

- [ ] Any submitted sample files or attachments satisfy the submission policy, or none are submitted.

## Summary

<!-- What changed and why? -->

## Related issue

<!-- Use "Fixes #N" when this PR fully resolves an issue. -->

## Testing

<!-- List the exact commands and manual checks run. -->

## Visual impact

<!-- Select exactly one. A reason is required when selecting no rendered change. -->

- [ ] No rendered PDF change
- [ ] Rendered PDF change or visual evidence added
- Reason: <!-- Required for "No rendered PDF change" -->

## Visual audit

<!-- Required when rendered output changes or visual evidence assets under assets/bugfixes/ change; Markdown and text bookkeeping are exempt. Delete only when "No rendered PDF change" is selected. -->

- Issue: #<!-- N -->
- Fixture: <!-- repository path -->
- Page(s): <!-- compared page numbers -->
- Renderer and DPI: <!-- e.g. pdftoppm, 150 DPI -->
- Evidence mode: `fix` <!-- `fix` requires gt/before/after; `defect` requires compare -->
- Layout audit report: `assets/bugfixes/issue-<!-- N -->/layout-audit.json` <!-- fix mode only -->
- Render cluster reports: `assets/bugfixes/issue-<!-- N -->/render-clusters-page-<!-- P -->.json` <!-- fix mode: one strict report per compared page -->
- Reference exporter differences: None <!-- or assets/bugfixes/issue-N/reference-exporter-differences.json -->
- Fine-detail threshold: <!-- e.g. 0.5pt; must match compare_layout.py --fine-shift -->
- Layout audit page count: <!-- Pass, or #N references when the report differs -->
- Layout audit text flow: <!-- Pass, #N, or ref:<id> for an exact verified painted-text visibility or rasterized-text difference -->
- Layout audit visible fills: <!-- Pass, or #N references for visible-fill occlusions -->
- Layout audit rectangle geometry: <!-- Pass, or #N references for matched rectangle position/size/edge findings -->
- Layout audit large shifts: <!-- Pass, #N open-issue references, or ref:<id> for an exact text-shift record (page/label/occurrence/dx/dy) above the report threshold -->
- Layout audit fine shifts: <!-- Pass, #N open-issue references, or ref:<id> for an exact text-shift record (page/label/occurrence/dx/dy) above the fine-detail threshold -->
- New follow-up issues found in this audit: <!-- #N, #N or None; create issues before completing the audit -->
- Model vision findings: <!-- Describe what Codex/Claude saw in the full pages, diff, and crops. Numeric output is insufficient. -->
- GT: `assets/bugfixes/issue-<!-- N -->/gt.jpg`
- Before: `assets/bugfixes/issue-<!-- N -->/before.jpg`
- After: `assets/bugfixes/issue-<!-- N -->/after.jpg`
- Native: None <!-- required as assets/bugfixes/issue-N/native.jpg when reference exporter differences are used -->
- Compare: `assets/bugfixes/issue-<!-- N -->/compare.jpg`

### Visual comparison

<!-- Replace each cell comment with rendered Markdown image syntax using a stable commit-pinned raw URL or GitHub attachment URL. -->

| GT | Before | After |
| --- | --- | --- |
| <!-- ![GT](IMAGE_URL) --> | <!-- ![Before](IMAGE_URL) --> | <!-- ![After](IMAGE_URL) --> |

<!-- Reference exporter difference only: render the native Office evidence. -->

<!-- ![Native](IMAGE_URL) -->

<!-- Defect mode only: replace the cell comment with the rendered Compare image and remove the GT/Before/After table. -->

| Compare |
| --- |
| <!-- ![Compare](IMAGE_URL) --> |

### Required inspection

- [ ] Rendered all evidence at 150 DPI or higher
- [ ] Stored progressive JPEG quality 86 assets with metadata stripped
- [ ] Used Codex/Claude vision to inspect the full GT/output pages, diff, and matched crops
- [ ] Inspected matched region crops at full resolution
- [ ] Ran compare_layout.py --audit --fine-shift PT and dispositioned every fine/large text-instance shift, rectangle geometry deviation, painted-text visibility mismatch, and visible-fill occlusion
- [ ] Ran compare_render.py --cluster-report PATH --strict-clusters and dispositioned every material 5% fuzz diff cluster by explicit ID
- [ ] Inventoried hairlines and border dash styles
- [ ] Inventoried font weight, italic, and underline emphasis

### Deviation audit

<!-- Every result must start with: Matches GT, Fixed, No deviation observed, Reference difference: ref:<id> (comma-separated exact IDs are allowed), or Remaining: #N. -->

| Check | Result |
| --- | --- |
| Page count/order | <!-- status --> |
| Element presence | <!-- status --> |
| Position/size | <!-- status --> |
| Rotation/flip | <!-- status --> |
| Fill | <!-- status --> |
| Stroke/border | <!-- status --> |
| Shape outline geometry | <!-- status --> |
| Text content | <!-- status --> |
| Font family/weight/style | <!-- status --> |
| Text color | <!-- status --> |
| Alignment | <!-- status --> |
| Line/paragraph spacing | <!-- status --> |
| Clipping/overflow | <!-- status --> |

## Checklist

- [ ] Commits include a `Signed-off-by` line
- [ ] PR scope contains one root cause
- [ ] Remaining converter or harness deviations each reference an open issue
