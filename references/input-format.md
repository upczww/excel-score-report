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

## Recognizable non-standard layouts

The converter can normalize a workbook when the first worksheet contains recognizable aliases for the class and name columns, such as `班别`/`班`/`class` and `学生`/`考生`/`name`. It also recognizes question labels containing `选择题`/`单选`/`多选`, `填空题`, and `解答题`/`计算题`/`证明题`; when labels have no group words, remaining question columns are assigned positionally as 10 select, 6 fill, then free-response columns. A following row containing `满分`/`目标分值` or numeric score-like values is treated as the target row. Summary rows such as `合计` and `平均分` are skipped.

For a normalized input, the command writes a template-shaped workbook (by default `<input>.standardized.xlsx`) and reports a source-to-output `mapping` in JSON before/alongside generating the PDF. Layouts without both class/name columns, without free-response columns, or exceeding the fixed 10 select / 6 fill capacity fail safely. If a workbook is still ambiguous, copy `assets/模板.xlsx` and adapt the data rather than relying on an unsafe guess.
