# Supported workbook format

The offline Node.js generator reads only the first worksheet and expects this fixed layout:

| Excel location | Meaning |
| --- | --- |
| Row 1 | Column headers |
| Row 2 | Target/full scores for free-response questions |
| Rows 3 onward | One student per row |
| Column A | `班级` |
| Column B | `姓名` |
| Columns C–L | 10 multiple-choice question scores |
| Columns M–R | 6 fill-in-the-blank question scores |
| Columns S onward | Free-response question scores; the count may vary |

Question labels come from row 1. Free-response target scores come from row 2, columns S onward. Rows with a blank `姓名` are skipped. Both `.xlsx` and `.xls` inputs are supported.

The output is one combined, landscape PDF. Each parsed student occupies one page containing the title, name, class, and separate tables for the three question groups.

If a workbook does not match this layout, ask the user to adapt it or copy `assets/模板.xlsx` from the installed skill. Do not guess a column mapping when doing so could assign scores to the wrong questions or students.
