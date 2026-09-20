---
name: excel-score-report
description: Generate a paginated student exam score PDF from a Chinese item-level score Excel workbook with the hosted 小题分报告生成器. Use for 小题分、逐题得分、成绩条, or one-student-per-page PDF requests based on the supported workbook template. Do not use for generic Excel-to-PDF conversion or grade analysis.
---

# Excel Score Report

Turn a compatible exam score workbook into one landscape PDF with one page per student by using <https://excel-score-pdf.vercel.app>.

## Inputs

Resolve these from the request and workspace before asking questions:

- Source `.xlsx` or `.xls` workbook.
- Optional report title. The site defaults to the workbook filename.
- Optional output path. Otherwise, save beside the source workbook with the chosen title.

If workbook compatibility is uncertain, read [references/input-format.md](references/input-format.md) before opening the site.

## Workflow

1. Confirm the source file exists and preserve it unchanged.
2. Open the hosted generator in a browser that supports file upload and download.
3. Upload the workbook through the upload area.
4. Set `PDF 标题` only when the user supplied a title or asked for a different one.
5. Select `生成 PDF` and wait for either `下载 PDF` or a visible error.
6. Download the result and move or rename it to the requested output path when needed.
7. Verify that the result exists, is non-empty, begins with the PDF signature, and—when a PDF inspection tool is available—has one page per parsed student.
8. Return the final file path, title, and parsed student count. Mention skipped blank-name rows or workbook-format problems when relevant.

## Operational Constraints

- The current application processes the workbook in the browser; do not upload it to a different conversion service.
- Use only the first worksheet. Do not merge sheets or silently reinterpret columns.
- Do not change question counts, scores, student names, or class names to make an incompatible workbook pass.
- If the browser cannot upload or download files, give the user the hosted URL and the exact compatibility issue instead of substituting an unrelated converter.
- The generator downloads a Chinese font at runtime. Retry once after a transient font-loading failure, then report the failure clearly.

## Template Requests

When the user needs a blank workbook, use the site's `下载模板` link or <https://excel-score-pdf.vercel.app/%E6%A8%A1%E6%9D%BF.xlsx>. Explain the required rows and columns using [references/input-format.md](references/input-format.md).
