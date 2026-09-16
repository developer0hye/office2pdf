#!/usr/bin/env python3
"""Build a category-count variant of the income bar chart for issue #1621.

`scripts/probe_harness.py` patches exactly one OOXML part per spec, but a
category-count change on this fixture is not expressible that way: native
Excel re-derives the income chart's categories from the workbook-scoped
`CategoriesIncome` defined-name array literal (`xl/workbook.xml`) and its
values from the live range named in `c:val`'s `c:f`
(`chart_calcs!$D$19:$D$23`, `xl/charts/chart4.xml`'s income counterpart is
`chart3.xml`), both outside the chart part itself. A patch confined to
`xl/charts/chart3.xml`'s own `c:cat`/`c:val` cache is silently overwritten by
Excel on open — confirmed empirically (see README.md's "What this falsifies
(new)" section) — so a working probe has to patch the workbook-level array
*and* widen the chart's value range together, in the same package.

This appends one or more synthetic "probe extra"/"probe extra N" categories,
each with a zero value (the value magnitude doesn't matter for a
baseline-placement probe), by:

1. Adding one array element per new category to the end of
   `CategoriesIncome`'s array literal in `xl/workbook.xml`.
2. Widening `c:val`'s `c:f` range in `xl/charts/chart3.xml` by one cell per
   new category -- `chart_calcs!$D$19:$D${23+added}`, e.g. `D24` alone for a
   single added category (`N=6`) or through `D27` for four (`N=9`) -- all of
   which are blank in the base fixture and read as 0.
3. Extending `c:cat`/`c:val`'s own cache and `c:ptCount` to match, so
   office2pdf (which reads the cache directly and does not evaluate formulas
   or defined names) renders the same category count as native Excel.

Usage:
    python3 assets/validation/issue-1621/build_category_count_probe.py N OUT.xlsx

`N` is the total category count (6-9; the base fixture always has 5, and
`NEW_CATEGORIES` below has 4 entries). Each call rebuilds from the base
fixture directly -- it does not chain off a previous output -- so `N=9` in
one call appends all 4 synthetic categories at once. Extend
`NEW_CATEGORIES` to probe past 9.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

BASE = (
    Path(__file__).resolve().parent.parent.parent.parent
    / "tests"
    / "fixtures"
    / "xlsx"
    / "issue_1181_fit_to_height.xlsx"
)
NEW_CATEGORIES = ["probe extra", "probe extra 2", "probe extra 3", "probe extra 4"]


def build(target_count: int, out_path: Path) -> None:
    if target_count <= 5:
        raise SystemExit("target category count must be > 5 (the base fixture already has 5)")
    added = target_count - 5
    if added > len(NEW_CATEGORIES):
        raise SystemExit(f"only {len(NEW_CATEGORIES)} synthetic categories defined")
    new_names = NEW_CATEGORIES[:added]

    with zipfile.ZipFile(BASE) as zin:
        wb = zin.read("xl/workbook.xml").decode("utf-8")
        chart3 = zin.read("xl/charts/chart3.xml").decode("utf-8")

        old_name = (
            '<definedName name="CategoriesIncome">{"financial aid";'
            '"wages (after-tax)";"family help";"from savings";"other"}</definedName>'
        )
        quoted_new = ";".join(f'"{name}"' for name in new_names)
        new_name = old_name.replace(
            '"other"}</definedName>', f'"other";{quoted_new}}}</definedName>'
        )
        if old_name not in wb:
            raise SystemExit("CategoriesIncome defined name not found or already patched")
        wb2 = wb.replace(old_name, new_name)

        last_value_cell = 23 + added  # base range is $D$19:$D$23
        old_range = "<c:f>chart_calcs!$D$19:$D$23</c:f>"
        new_range = f"<c:f>chart_calcs!$D$19:$D${last_value_cell}</c:f>"
        if old_range not in chart3:
            raise SystemExit("c:val range not found or already patched")
        chart3 = chart3.replace(old_range, new_range)

        old_cat_ptcount = '<c:strCache><c:ptCount val="5"/>'
        chart3 = chart3.replace(old_cat_ptcount, f'<c:strCache><c:ptCount val="{target_count}"/>')

        cat_pts = "".join(
            f'<c:pt idx="{5 + i}"><c:v>{name}</c:v></c:pt>' for i, name in enumerate(new_names)
        )
        old_cat_last = '<c:pt idx="4"><c:v>other</c:v></c:pt></c:strCache>'
        chart3 = chart3.replace(old_cat_last, f'<c:pt idx="4"><c:v>other</c:v></c:pt>{cat_pts}</c:strCache>')

        old_val_ptcount = '<c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="5"/>'
        chart3 = chart3.replace(
            old_val_ptcount, f'<c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="{target_count}"/>'
        )

        val_pts = "".join(f'<c:pt idx="{5 + i}"><c:v>0</c:v></c:pt>' for i in range(added))
        old_val_last = '<c:pt idx="4"><c:v>6.1224489795918366E-2</c:v></c:pt></c:numCache>'
        chart3 = chart3.replace(
            old_val_last, f'<c:pt idx="4"><c:v>6.1224489795918366E-2</c:v></c:pt>{val_pts}</c:numCache>'
        )

        with zipfile.ZipFile(out_path, "w") as zout:
            for info in zin.infolist():
                member = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                member.compress_type = zipfile.ZIP_DEFLATED
                member.external_attr = info.external_attr
                member.create_system = 3
                if info.filename == "xl/workbook.xml":
                    data = wb2.encode("utf-8")
                elif info.filename == "xl/charts/chart3.xml":
                    data = chart3.encode("utf-8")
                else:
                    data = zin.read(info.filename)
                zout.writestr(member, data)


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    target_count = int(sys.argv[1])
    out_path = Path(sys.argv[2])
    build(target_count, out_path)
    print(f"wrote {out_path} ({target_count} income categories)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
