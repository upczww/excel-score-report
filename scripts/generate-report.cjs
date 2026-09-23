#!/usr/bin/env node
'use strict';

// Offline Node.js port of:
//   web/src/lib/parseExcel.ts
//   web/src/lib/generatePdf.ts
// The runtime dependencies and Chinese font are vendored with this skill.

const fs = require('node:fs');
const path = require('node:path');

const XLSX = require('../vendor/xlsx.full.min.js');
const { jsPDF } = require('../vendor/jspdf.umd.min.js');
const autoTable = require('../vendor/jspdf.plugin.autotable.js').default;

const SKILL_DIR = path.resolve(__dirname, '..');
const FONT_PATH = path.join(SKILL_DIR, 'assets', 'NotoSansSC-Bold.ttf');
const TEMPLATE_PATH = path.join(SKILL_DIR, 'assets', '模板.xlsx');

// 页面尺寸 (mm) — 对应原 Python 的 Cm(28) x Cm(20)
const PAGE_WIDTH = 280;
const PAGE_HEIGHT = 200;
const MARGIN = 10;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2;
const FONT_SIZE = 22;

let fontBase64 = null;

function fail(message) {
  const error = new Error(message);
  error.userFacing = true;
  throw error;
}

// Copied from web/src/lib/parseExcel.ts. The only Node-specific change is
// accepting a Buffer/Uint8Array in addition to a browser ArrayBuffer.
function parseExcel(buffer, fileName) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('Excel 文件中没有工作表');
  }
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  if (data.length < 3) {
    throw new Error('Excel 文件至少需要 3 行（表头 + 分值行 + 学生数据）');
  }

  const headers = data[0].map(String);
  const targetRow = data[1];

  // 列结构: 0=班级, 1=姓名, 2-11=选择题, 12-17=填空题, 18+=解答题
  const selectQuestionNames = headers.slice(2, 12);
  const fillQuestionNames = headers.slice(12, 18);
  const solveQuestionNames = headers.slice(18);
  const solveTargets = targetRow.slice(18).map((value) => value ?? '');

  const students = [];

  for (let index = 2; index < data.length; index += 1) {
    const row = data[index];
    if (!row || row.length < 2) continue;

    const className = String(row[0] ?? '');
    const name = String(row[1] ?? '');
    if (!name) continue;

    students.push({
      name,
      className,
      selectQuestionNames,
      selectScores: row.slice(2, 12).map((value) => value ?? ''),
      fillQuestionNames,
      fillScores: row.slice(12, 18).map((value) => value ?? ''),
      solveQuestionNames,
      solveTargets,
      solveScores: row.slice(18).map((value) => value ?? ''),
    });
  }

  return {
    students,
    fileName: fileName.replace(/\.[^.]+$/, ''),
    sheetName,
  };
}

const CLASS_HEADER_ALIASES = new Set([
  '班级', '班别', '班', 'class', 'classname', 'class_name', 'classno', 'classid',
]);
const NAME_HEADER_ALIASES = new Set([
  '姓名', '学生姓名', '学生', '考生', 'name', 'student', 'studentname', 'student_name',
]);
const NON_QUESTION_HEADER = /序号|学号|编号|总分|总成绩|总计|合计|排名|名次|备注|评语|total|rank|remark|comment|^id$/i;
const SUMMARY_NAME = /^(满分|目标分值|分值|总分|平均分|班级平均|合计|最高分|最低分|排名|平均)$/;

function textValue(value) {
  return value == null ? '' : String(value).trim();
}

function normalizedHeader(value) {
  return textValue(value)
    .toLowerCase()
    .replace(/[\s_\-–—]/g, '')
    .replace(/[（(].*?[）)]/g, '');
}

function findHeader(headers, aliases) {
  return headers.findIndex((header) => aliases.has(normalizedHeader(header)));
}

function questionGroup(header) {
  const value = normalizedHeader(header);
  if (/选择题|单选|多选|choice|mcq/.test(value)) return 'select';
  if (/填空题|填充|fill|blank/.test(value)) return 'fill';
  if (/解答题|计算题|证明题|应用题|主观题|solve|response|subjective/.test(value)) return 'solve';
  return null;
}

function isScoreLike(value) {
  const text = textValue(value);
  return text === '' || /^(?:[-+]?\d+(?:\.\d+)?|[-+]?\d+(?:\.\d+)?\s*分|满分|目标分值|分值)$/i.test(text);
}

function looksLikeTargetRow(row, questionColumns, classIndex, nameIndex) {
  const classValue = textValue(row[classIndex]);
  const nameValue = textValue(row[nameIndex]);
  const labelRow = /^(满分|目标分值|分值)$/i.test(classValue) || /^(满分|目标分值|分值)$/i.test(nameValue);
  if (nameValue && !labelRow) return false;
  const values = questionColumns.map((column) => row[column]).filter((value) => textValue(value) !== '');
  if (values.length < 2) return false;
  return values.filter(isScoreLike).length >= Math.max(2, Math.ceil(values.length * 0.7));
}

function extractHeaderTarget(header) {
  const text = textValue(header);
  if (!/分|满分|score/i.test(text)) return '';
  const match = text.match(/([-+]?\d+(?:\.\d+)?)\s*分?/i);
  return match ? `${match[1]}分` : '';
}

function isStandardWorkbook(data) {
  if (data.length < 3) return false;
  const headers = data[0].map(textValue);
  const targetRow = data[1] || [];
  return headers[0] === '班级'
    && headers[1] === '姓名'
    && headers.length >= 19
    && !textValue(targetRow[1]);
}

function uniqueLabels(columns, headers, prefix, count) {
  const labels = columns.map((column) => textValue(headers[column]) || `${prefix}${column + 1}`);
  while (labels.length < count) labels.push(`${prefix}${labels.length + 1}`);
  return labels.slice(0, count);
}

function standardizeWorkbook(buffer, fileName) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('Excel 文件中没有工作表');
  const sourceSheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, defval: '' });

  if (isStandardWorkbook(data)) {
    return { buffer, changed: false, mapping: null, sourceSheet: sheetName };
  }

  const headerRowIndex = data.slice(0, Math.min(data.length, 12)).findIndex((row) => {
    const headers = row.map(textValue);
    return findHeader(headers, CLASS_HEADER_ALIASES) >= 0 && findHeader(headers, NAME_HEADER_ALIASES) >= 0;
  });
  if (headerRowIndex < 0) {
    throw new Error('未检测到“班级”和“姓名”列；无法安全转换为模板格式');
  }

  const headers = data[headerRowIndex].map(textValue);
  const classIndex = findHeader(headers, CLASS_HEADER_ALIASES);
  const nameIndex = findHeader(headers, NAME_HEADER_ALIASES);
  const questionColumns = headers
    .map((header, index) => ({ header, index }))
    .filter(({ header, index }) => index !== classIndex && index !== nameIndex && header && !NON_QUESTION_HEADER.test(header))
    .map(({ index }) => index);
  if (questionColumns.length < 1) throw new Error('未检测到题目列；请检查 Excel 表头');

  let dataStart = headerRowIndex + 1;
  let targetRow = null;
  if (data[dataStart] && looksLikeTargetRow(data[dataStart], questionColumns, classIndex, nameIndex)) {
    targetRow = data[dataStart];
    dataStart += 1;
  }

  const groups = { select: [], fill: [], solve: [] };
  const unclassified = [];
  for (const column of questionColumns) {
    const group = questionGroup(headers[column]);
    if (group) groups[group].push(column);
    else unclassified.push(column);
  }

  if (groups.select.length === 0 && groups.fill.length === 0 && groups.solve.length === 0) {
    groups.select = unclassified.slice(0, 10);
    groups.fill = unclassified.slice(10, 16);
    groups.solve = unclassified.slice(16);
  } else {
    for (const column of unclassified) {
      if (groups.select.length < 10) groups.select.push(column);
      else if (groups.fill.length < 6) groups.fill.push(column);
      else groups.solve.push(column);
    }
  }

  if (groups.select.length > 10) throw new Error(`检测到 ${groups.select.length} 个选择题列，超过模板固定的 10 列`);
  if (groups.fill.length > 6) throw new Error(`检测到 ${groups.fill.length} 个填空题列，超过模板固定的 6 列`);
  if (groups.solve.length === 0) throw new Error('未检测到解答题列，无法生成模板格式的报告');

  const selectNames = uniqueLabels(groups.select, headers, '选择题', 10);
  const fillNames = uniqueLabels(groups.fill, headers, '填空题', 6);
  const solveNames = groups.solve.map((column, index) => textValue(headers[column]) || `解答题${index + 1}`);
  const target = (column) => textValue(targetRow && targetRow[column]) || extractHeaderTarget(headers[column]);
  const solveTargets = groups.solve.map(target);
  const rows = [
    ['班级', '姓名', ...selectNames, ...fillNames, ...solveNames],
    ['', '', ...groups.select.map(target), ...groups.fill.map(target), ...solveTargets],
  ];
  const students = [];
  for (let rowIndex = dataStart; rowIndex < data.length; rowIndex += 1) {
    const row = data[rowIndex] || [];
    const name = textValue(row[nameIndex]);
    if (!name || SUMMARY_NAME.test(name)) continue;
    const valueAt = (column) => row[column] ?? '';
    students.push([
      valueAt(classIndex),
      name,
      ...groups.select.map(valueAt),
      ...groups.fill.map(valueAt),
      ...groups.solve.map(valueAt),
    ]);
  }
  if (students.length === 0) throw new Error('转换后没有找到学生数据；请确认姓名列和学生行');
  rows.push(...students);

  const normalizedSheet = XLSX.utils.aoa_to_sheet(rows);
  const normalizedBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(normalizedBook, normalizedSheet, '小题分');
  const normalizedBuffer = XLSX.write(normalizedBook, { bookType: 'xlsx', type: 'buffer' });
  const mapping = {
    sourceSheet: sheetName,
    headerRow: headerRowIndex + 1,
    sourceClassColumn: headers[classIndex],
    sourceNameColumn: headers[nameIndex],
    targetRow: targetRow ? dataStart : null,
    outputColumns: {
      select: groups.select.map((column, index) => ({ source: headers[column], output: selectNames[index] })),
      fill: groups.fill.map((column, index) => ({ source: headers[column], output: fillNames[index] })),
      solve: groups.solve.map((column, index) => ({ source: headers[column], output: solveNames[index] })),
    },
    skippedRows: data.length - dataStart - students.length,
  };
  return { buffer: normalizedBuffer, changed: true, mapping, sourceSheet: sheetName };
}

function hasChinese(text) {
  return /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(text);
}

function loadChineseFont() {
  if (fontBase64) return fontBase64;
  if (!fs.existsSync(FONT_PATH)) {
    throw new Error(`Skill 缺少内置中文字体: ${FONT_PATH}`);
  }
  fontBase64 = fs.readFileSync(FONT_PATH).toString('base64');
  return fontBase64;
}

function registerFont(doc, base64) {
  doc.addFileToVFS('NotoSansSC-Bold.ttf', base64);
  doc.addFont('NotoSansSC-Bold.ttf', 'NotoSansSC', 'bold');
}

// Copied from web/src/lib/generatePdf.ts.
function drawTable(doc, y, body) {
  autoTable(doc, {
    startY: y,
    margin: { left: MARGIN, right: MARGIN },
    theme: 'grid',
    styles: {
      font: 'times',
      fontStyle: 'bold',
      fontSize: FONT_SIZE,
      halign: 'center',
      valign: 'middle',
      textColor: [0, 0, 0],
      lineColor: [0, 0, 0],
      lineWidth: 0.35,
      fillColor: false,
      cellPadding: { top: 1, bottom: 1, left: 1, right: 1 },
    },
    head: [],
    body,
    tableWidth: CONTENT_WIDTH,
    didParseCell: (data) => {
      const text = data.cell.text.join('');
      data.cell.styles.font = hasChinese(text) ? 'NotoSansSC' : 'times';
      data.cell.styles.fontStyle = 'bold';
    },
  });
  return doc.lastAutoTable.finalY;
}

function generatePdf(students, titleName, onProgress) {
  const fontData = loadChineseFont();

  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: [PAGE_WIDTH, PAGE_HEIGHT],
  });

  registerFont(doc, fontData);

  for (let studentIndex = 0; studentIndex < students.length; studentIndex += 1) {
    const student = students[studentIndex];
    if (studentIndex > 0) doc.addPage([PAGE_WIDTH, PAGE_HEIGHT], 'landscape');

    if (onProgress) onProgress(studentIndex + 1, students.length);

    let y = MARGIN;
    const LINE_HEIGHT = 9;

    doc.setFontSize(FONT_SIZE);
    doc.setFont('NotoSansSC', 'bold');
    const titleText = `${titleName}    姓名: ${student.name}    班级: ${student.className}`;
    doc.text(titleText, MARGIN, y + 7);
    y += LINE_HEIGHT;

    doc.setFontSize(FONT_SIZE);
    doc.setFont('NotoSansSC', 'bold');
    doc.text('一、选择题（每题4分，共40分）', MARGIN, y + 7);
    y += LINE_HEIGHT;

    y = drawTable(doc, y, [
      ['题号', ...student.selectQuestionNames],
      ['得分', ...student.selectScores.map(String)],
    ]);

    doc.setFontSize(FONT_SIZE);
    doc.setFont('NotoSansSC', 'bold');
    doc.text('二、填空题（每题4分，共24分）', MARGIN, y + 7);
    y += LINE_HEIGHT;

    y = drawTable(doc, y, [
      ['题号', ...student.fillQuestionNames],
      ['得分', ...student.fillScores.map(String)],
    ]);

    doc.setFontSize(FONT_SIZE);
    doc.setFont('NotoSansSC', 'bold');
    doc.text('三、解答题', MARGIN, y + 7);
    y += LINE_HEIGHT;

    const maxDataCols = 11;
    const totalSolve = student.solveScores.length;
    let tableIndex = 0;

    for (let start = 0; start < totalSolve; start += maxDataCols) {
      const end = Math.min(start + maxDataCols, totalSolve);
      const names = student.solveQuestionNames.slice(start, end);
      const targets = student.solveTargets.slice(start, end).map(String);
      const scores = student.solveScores.slice(start, end).map(String);

      y = drawTable(doc, y, [
        ['题号', ...names],
        ['分值', ...targets],
        ['得分', ...scores],
      ]);

      tableIndex += 1;
      if (tableIndex < 2 && start + maxDataCols < totalSolve) {
        y += LINE_HEIGHT;
      }
    }
  }

  return {
    buffer: Buffer.from(doc.output('arraybuffer')),
    pages: doc.getNumberOfPages(),
  };
}

function help() {
  return `小题分报告生成器（离线 Node.js 版）

用法:
  node scripts/generate-report.cjs <成绩表.xlsx> [-o 输出.pdf] [-t 标题]
  node scripts/generate-report.cjs --copy-template <目标路径>

选项:
  -o, --output PATH       输出 PDF 路径；默认与输入文件同目录
  -t, --title TEXT        PDF 标题；默认使用输入文件名
      --normalized-output PATH
                          非标准工作簿转换后的模板文件路径；默认自动命名
      --copy-template PATH 复制 Skill 内置的 Excel 模板
      --force             允许覆盖已存在的文件
  -h, --help              显示帮助
`;
}

function parseArgs(argv) {
  const options = {
    input: null,
    output: null,
    title: null,
    normalizedOutput: null,
    copyTemplate: null,
    force: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    const nextValue = () => {
      index += 1;
      if (index >= argv.length) fail(`选项 ${argument} 缺少参数`);
      return argv[index];
    };

    if (argument === '-h' || argument === '--help') {
      options.help = true;
    } else if (argument === '--force') {
      options.force = true;
    } else if (argument === '-o' || argument === '--output') {
      options.output = nextValue();
    } else if (argument.startsWith('--output=')) {
      options.output = argument.slice('--output='.length);
    } else if (argument === '-t' || argument === '--title') {
      options.title = nextValue();
    } else if (argument.startsWith('--title=')) {
      options.title = argument.slice('--title='.length);
    } else if (argument === '--normalized-output') {
      options.normalizedOutput = nextValue();
    } else if (argument.startsWith('--normalized-output=')) {
      options.normalizedOutput = argument.slice('--normalized-output='.length);
    } else if (argument === '--copy-template') {
      options.copyTemplate = nextValue();
    } else if (argument.startsWith('--copy-template=')) {
      options.copyTemplate = argument.slice('--copy-template='.length);
    } else if (argument.startsWith('-')) {
      fail(`未知选项: ${argument}`);
    } else if (options.input) {
      fail('一次只能处理一个 Excel 文件');
    } else {
      options.input = argument;
    }
  }

  return options;
}

function ensureParent(filePath) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}

function copyTemplate(destination, force) {
  if (!fs.existsSync(TEMPLATE_PATH)) fail(`Skill 缺少内置模板: ${TEMPLATE_PATH}`);
  let target = path.resolve(destination);
  if (fs.existsSync(target) && fs.statSync(target).isDirectory()) {
    target = path.join(target, path.basename(TEMPLATE_PATH));
  }
  if (fs.existsSync(target) && !force) {
    fail(`目标文件已存在: ${target}；如需覆盖请使用 --force`);
  }
  ensureParent(target);
  fs.copyFileSync(TEMPLATE_PATH, target);
  return target;
}

function availableOutput(requested) {
  if (!fs.existsSync(requested)) return requested;
  const parsed = path.parse(requested);
  for (let counter = 1; ; counter += 1) {
    const candidate = path.join(parsed.dir, `${parsed.name} (${counter})${parsed.ext}`);
    if (!fs.existsSync(candidate)) return candidate;
  }
}

function writeOutput(destination, buffer, force, explicitOutput) {
  const target = (explicitOutput || force) ? destination : availableOutput(destination);
  if (fs.existsSync(target) && !force) {
    fail(`输出文件已存在: ${target}；如需覆盖请使用 --force`);
  }
  ensureParent(target);
  fs.writeFileSync(target, buffer, { flag: force ? 'w' : 'wx' });
  return target;
}

function main() {
  const nodeMajor = Number(process.versions.node.split('.')[0]);
  if (!Number.isFinite(nodeMajor) || nodeMajor < 18) {
    fail(`需要 Node.js 18 或更高版本，当前版本为 ${process.versions.node}`);
  }

  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    process.stdout.write(help());
    return;
  }

  if (options.copyTemplate) {
    if (options.input) fail('--copy-template 不能与输入文件同时使用');
    const target = copyTemplate(options.copyTemplate, options.force);
    process.stdout.write(`${JSON.stringify({ template: target, offline: true })}\n`);
    return;
  }

  if (!options.input) fail('请提供输入 Excel，或使用 --copy-template 导出内置模板');

  const input = path.resolve(options.input);
  if (!fs.existsSync(input) || !fs.statSync(input).isFile()) {
    fail(`输入文件不存在: ${input}`);
  }
  if (!['.xlsx', '.xls'].includes(path.extname(input).toLowerCase())) {
    fail('只支持 .xlsx 和 .xls 文件');
  }

  const rawBuffer = fs.readFileSync(input);
  const prepared = standardizeWorkbook(rawBuffer, path.basename(input));
  const parsed = parseExcel(prepared.buffer, path.basename(input));
  if (parsed.students.length === 0) {
    fail('未找到学生数据，请检查 Excel 格式');
  }

  const title = options.title && options.title.trim() ? options.title.trim() : parsed.fileName;
  const generated = generatePdf(parsed.students, title);
  let standardizedOutput = null;
  if (prepared.changed || options.normalizedOutput) {
    const requestedNormalized = options.normalizedOutput
      ? path.resolve(options.normalizedOutput)
      : path.join(path.dirname(input), `${parsed.fileName}.standardized.xlsx`);
    standardizedOutput = writeOutput(
      requestedNormalized,
      prepared.buffer,
      options.force,
      Boolean(options.normalizedOutput),
    );
  }
  const requestedOutput = options.output
    ? path.resolve(options.output)
    : path.join(path.dirname(input), `${parsed.fileName}.pdf`);
  const output = writeOutput(requestedOutput, generated.buffer, options.force, Boolean(options.output));

  const signature = fs.readFileSync(output).subarray(0, 5).toString('ascii');
  if (signature !== '%PDF-') fail('生成结果不是有效的 PDF 文件');
  if (generated.pages !== parsed.students.length) {
    fail(`PDF 页数 ${generated.pages} 与学生数 ${parsed.students.length} 不一致`);
  }

  process.stdout.write(`${JSON.stringify({
    output,
    students: parsed.students.length,
    pages: generated.pages,
    title,
    sheet: parsed.sheetName,
    standardized: prepared.changed,
    standardizedOutput,
    mapping: prepared.mapping,
    engine: 'node',
    offline: true,
  })}\n`);
}

try {
  main();
} catch (error) {
  const message = error && error.message ? error.message : String(error);
  process.stderr.write(`error: ${message}\n`);
  process.exitCode = 2;
}
