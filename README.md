# 小题分报告生成器

`excel-score-report` 是一个完全离线的 Codex Skill。它把符合模板的考试逐题得分 Excel 工作簿，生成为一份每名学生一页的 PDF 成绩报告。

转换逻辑直接来自原 Web 应用的 `parseExcel.ts` 和 `generatePdf.ts`，继续使用 `xlsx`、`jsPDF` 与 `jspdf-autotable`。运行时所需的 JavaScript 库、中文字体和 Excel 模板都已经随仓库发布，不依赖在线应用、浏览器、CDN 或 `npm install`。

## 安装

在 Codex 中调用 `$skill-installer`：

```text
$skill-installer 从 https://github.com/upczww/excel-score-report 安装 Skill
```

也可以手动安装：

```bash
mkdir -p ~/.agents/skills
git clone https://github.com/upczww/excel-score-report.git ~/.agents/skills/excel-score-report
```

运行环境只需要 Node.js 18 或更高版本。

## 使用 Skill

```text
$excel-score-report 把 ./八上数学期中小题分.xlsx 生成为每名学生一页的 PDF，标题使用“八上数学期中小题分”。
```

## 直接运行离线转换器

```bash
node scripts/generate-report.cjs ./成绩表.xlsx \
  --output ./成绩报告.pdf \
  --title "期中数学小题分"
```

脚本同时支持 `.xlsx` 和 `.xls`。没有指定输出路径时，PDF 会生成在输入文件旁边；默认不会覆盖已有文件。

## 导出内置模板

```bash
node scripts/generate-report.cjs --copy-template ./模板.xlsx
```

模板文件也可以直接从仓库的 [`assets/模板.xlsx`](assets/%E6%A8%A1%E6%9D%BF.xlsx) 获取。

## Excel 格式

- 第 1 行：表头。
- 第 2 行：解答题分值。
- 第 3 行起：学生数据。
- A、B 列分别为 `班级`、`姓名`。
- C–L 列为 10 道选择题，M–R 列为 6 道填空题，S 列起为数量可变的解答题。
- 只处理第一个工作表；姓名为空的行会被跳过。

完整约束见 [输入格式说明](references/input-format.md)。

## 离线内容

- `scripts/generate-report.cjs`：Node.js 命令行转换器。
- `assets/模板.xlsx`：可直接填写的成绩表模板。
- `assets/NotoSansSC-Bold.ttf`：PDF 中文字体。
- `vendor/`：离线运行所需的 JavaScript 库。

成绩表只在本机内存中处理，不会上传到任何服务器。
