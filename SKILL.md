---
name: excel-score-report
description: Generate a paginated student exam score PDF entirely offline from a Chinese item-level score Excel workbook, detecting and standardizing compatible non-standard layouts first. Use for 小题分、逐题得分、成绩条, or one-student-per-page PDF requests based on the bundled template. Do not use for generic Excel-to-PDF conversion or grade analysis.
---

# Excel Score Report

Use the bundled Node.js converter to turn a compatible `.xlsx` or `.xls` workbook into one landscape PDF with one page per student. It first detects the standard layout; for a recognizable non-standard layout it creates a local template-shaped `.standardized.xlsx`, then generates the PDF from that normalized workbook. The converter, Excel template, JavaScript libraries, and Chinese font are all included in this skill.

Run the bundled converter directly. Never use browser automation, fetch a font, or upload the workbook to an online service. Runtime processing must remain local.

## Generate a report

Resolve the directory containing this `SKILL.md`, then run:

```bash
node "<skill-directory>/scripts/generate-report.cjs" "<input.xlsx>" --output "<output.pdf>" --title "<report title>"
```

- Require Node.js 18 or newer. Do not run `npm install`; runtime dependencies are vendored.
- Omit `--title` to use the workbook filename.
- Omit `--output` to save beside the workbook. If that default filename exists, the script chooses a numbered filename instead of overwriting it.
- For a non-standard workbook, the JSON result includes `standardized: true`, a `mapping`, and `standardizedOutput` (default: `<input>.standardized.xlsx`). Use `--normalized-output PATH` to choose that path explicitly.
- Use `--force` only when the user explicitly asked to replace an existing output.
- Preserve the input workbook unchanged.

The command prints a JSON result. Treat the job as complete only when:

- `offline` is `true` and `engine` is `node`.
- `pages` equals `students`.
- When `standardized` is `true`, inspect `mapping` and retain the reported standardized workbook alongside the PDF.
- The reported output exists, is non-empty, and begins with `%PDF-`.

If workbook compatibility is uncertain, read [references/input-format.md](references/input-format.md) before running the converter.

## Provide the template

Copy the bundled template without using the network:

```bash
node "<skill-directory>/scripts/generate-report.cjs" --copy-template "<destination>/模板.xlsx"
```

Do not recreate the template manually. Use `--force` only when the user explicitly permits replacing an existing file.

## Constraints

- Process only the first worksheet, matching the original web implementation.
- Keep the fixed mapping of 10 multiple-choice columns, 6 fill-in columns, and the remaining free-response columns.
- Skip rows with a blank student name; do not invent or remap scores.
- Recognizable alternate headers such as `班别`/`学生` and grouped question labels are mapped locally; ambiguous layouts, missing columns, or more than 10 select / 6 fill questions fail safely with an actionable error instead of guessing.
- On a format error, report the exact issue and offer the bundled template.
- If Node.js is unavailable, report that requirement rather than falling back to an online conversion service.
