# 小题分报告生成器

`excel-score-report` 是 [小题分报告生成器](https://excel-score-pdf.vercel.app) 的 Codex Skill。它帮助智能体把符合模板的考试逐题得分 Excel 工作簿，生成为一份每名学生一页的 PDF 成绩报告。

## 能力

- 使用 `.xlsx` 或 `.xls` 成绩表生成逐生 PDF 报告。
- 分别呈现选择题、填空题和解答题得分。
- 保留用户指定的报告标题和输出位置。
- 在操作前识别不兼容的工作簿结构，避免错配学生或题目分数。
- 通过浏览器本地处理成绩表，不把文件上传到应用服务器。

## 安装

在 Codex 中调用 `$skill-installer`，并让它从该仓库安装：

```text
$skill-installer 从 https://github.com/upczww/excel-score-report 安装 Skill
```

也可以手动安装到用户级 Skill 目录：

```bash
git clone https://github.com/upczww/excel-score-report.git ~/.agents/skills/excel-score-report
```

安装后，如果 Skill 没有立即出现在列表中，请重启 Codex。

## 使用

```text
$excel-score-report 把 ./八上数学期中小题分.xlsx 生成为每名学生一页的 PDF，标题使用“八上数学期中小题分”。
```

## Excel 格式

- 第 1 行：表头。
- 第 2 行：解答题分值。
- 第 3 行起：学生数据。
- A、B 列分别为 `班级`、`姓名`。
- C–L 列为 10 道选择题，M–R 列为 6 道填空题，S 列起为数量可变的解答题。
- 当前只处理第一个工作表；姓名为空的行会被跳过。

不确定格式时，可下载应用提供的 [`模板.xlsx`](https://excel-score-pdf.vercel.app/%E6%A8%A1%E6%9D%BF.xlsx)。完整约束见 [输入格式说明](references/input-format.md)。

## 在线应用

<https://excel-score-pdf.vercel.app>
