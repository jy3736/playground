# Python-based attendance solution

This solution is a reusable Python version of the pure Excel attendance calculation.
It is intended for Excel files with a similar section-based structure:

1. Each employee has a section header such as `員工代號:00337  吳宗恩`.
2. Each section contains attendance rows with columns equivalent to `日期`, `星期`, `上班`, and `下班`.
3. The column order can vary if a recognizable header row is present.
4. Dates may be Excel date cells or text like `2026/06/08` or `2026-06-08`.
5. Times may be Excel time cells, plain text like `07:53`, or text ending in a time like `病假12:23`.

## Script

`attendance_python_solution.py`

## Basic usage

```bash
python3 attendance_python_solution.py ../data/2026.06月份.xlsx -o outputs/2026.06月份-python工時計算.xlsx
```

If the attendance sheet cannot be detected automatically, specify it:

```bash
python3 attendance_python_solution.py ../data/2026.06月份.xlsx --sheet 員工出勤明細表 -o outputs/result.xlsx
```

The default checkout cap is `19:30`. To change it:

```bash
python3 attendance_python_solution.py ../data/2026.06月份.xlsx --cap 18:30 -o outputs/result.xlsx
```

## Output workbook

The script creates a new output file by copying the source workbook structure.
It does not modify the source workbook.

The output keeps the original workbook sheets, including `算工作數` and `員工出勤明細表`.
It fills `算工作數` with one summary row per masked employee and month.
It also adds `計算明細` as an audit sheet for row-level review.
It does not add `處理說明` or other non-audit worksheet structures.

## Calculation rules

1. Saturday and Sunday are excluded.
2. Rows with missing start or end time do not count as work time.
3. Checkout later than the cap time is counted only up to the cap time.
4. The same employee on the same date counts as only one workday.
5. Multiple valid attendance segments on the same day are added to total work time.
6. Output employee identifiers in `算工作數` are masked.

## Verified sample

The script was verified against `../data/2026.06月份.xlsx`.

Result:

- Parsed attendance rows: 538
- Employees: 16
- Summary mismatches compared with `results.md`: 0

## Limits

This is intentionally generic for similar workbooks, not every possible attendance export.
It expects employee section headers in the first column and a recognizable date column in each section.
If a workbook uses a completely different layout, add a header mapping or normalize the workbook first.
