#!/usr/bin/env node
/**
 * 课程报告生成脚本 — 哈尔滨工程大学学报格式
 * 用法: node generate_report.js
 *
 * 依赖: npm install docx
 * 数据: 同级目录下需有 report_content.json
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun,
  Header, Footer, AlignmentType,
  HeadingLevel, PageNumber, LineRuleType, SectionType,
} = require('docx');

const content = JSON.parse(fs.readFileSync(__dirname + '/report_content.json', 'utf8'));

// ====== 格式常量 ======
const MARGINS = { top: 1474, bottom: 907, left: 1134, right: 1134 };
const FONT_SONG  = '宋体';
const FONT_HEI   = '黑体';
const FONT_KAI   = '楷体';
const FONT_FSONG = '仿宋';
const FONT_TNR   = 'Times New Roman';
const SZ44 = 44;  // 二号 22pt
const SZ28 = 28;  // 四号 14pt
const SZ24 = 24;  // 小四 12pt
const SZ21 = 21;  // 五号 10.5pt
const SZ18 = 18;  // 小五 9pt
const SINGLE   = 240;
const EXACT15  = { line: 300, lineRule: LineRuleType.EXACT };
const EXACT12  = { line: 240, lineRule: LineRuleType.EXACT };
const INDENT2  = 420;   // 2字符缩进 (10.5pt × 2 × 20)

function bodyPara(text) {
  return new Paragraph({
    spacing: { after: 0, line: SINGLE },
    indent: { firstLine: INDENT2 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, size: SZ21, font: FONT_SONG })],
  });
}
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 200, after: 200, ...EXACT15 },
    children: [new TextRun({ text, size: SZ28, bold: true, font: FONT_FSONG })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 60, after: 60, ...EXACT15 },
    children: [new TextRun({ text, size: SZ21, bold: true, font: FONT_HEI })],
  });
}

// Section 1: 单栏（题头 → 关键词）
const s1 = [];
s1.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 1200, after: 300, line: SINGLE },
  children: [new TextRun({ text: content.title, size: SZ44, bold: true, font: FONT_HEI })],
}));
s1.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120, ...EXACT15 },
  children: [new TextRun({ text: '作者姓名', size: SZ24, font: FONT_KAI })],
}));
s1.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200, ...EXACT15 },
  indent: { left: INDENT2, right: INDENT2 },
  children: [new TextRun({ text: '(哈尔滨工程大学 信息与通信工程学院，黑龙江 哈尔滨 150001)', size: SZ18, font: FONT_KAI })],
}));
s1.push(new Paragraph({
  spacing: { after: 0, ...EXACT15 },
  alignment: AlignmentType.JUSTIFIED, indent: { left: INDENT2, right: INDENT2 },
  children: [
    new TextRun({ text: '摘  要：', size: SZ18, bold: true, font: FONT_HEI }),
    new TextRun({ text: content.abstract, size: SZ18, font: FONT_KAI }),
  ],
}));
s1.push(new Paragraph({
  spacing: { after: 200, ...EXACT15 },
  alignment: AlignmentType.JUSTIFIED, indent: { left: INDENT2, right: INDENT2 },
  children: [
    new TextRun({ text: '关键词：', size: SZ18, bold: true, font: FONT_HEI }),
    new TextRun({ text: content.keywords, size: SZ18, font: FONT_SONG }),
  ],
}));

// Section 2: 双栏（正文）
const s2 = [];
for (const sec of content.sections) {
  if (sec.refs) {
    s2.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 200, after: 200, line: SINGLE },
      children: [new TextRun({ text: sec.heading, size: SZ28, bold: true, font: FONT_HEI })],
    }));
    for (const ref of sec.refs) {
      if (ref.type === 'en') {
        s2.push(new Paragraph({
          spacing: { after: 0, ...EXACT12 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun({ text: ref.text, size: SZ18, font: FONT_TNR })],
        }));
      } else if (ref.type === 'cn') {
        s2.push(new Paragraph({
          spacing: { after: 0, ...EXACT12 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun({ text: ref.cn, size: SZ18, font: FONT_SONG })],
        }));
        s2.push(new Paragraph({
          spacing: { after: 0, ...EXACT12 },
          alignment: AlignmentType.JUSTIFIED,
          children: [new TextRun({ text: ref.en, size: SZ18, font: FONT_TNR })],
        }));
      }
    }
    continue;
  }
  s2.push(h1(sec.heading));
  if (sec.subsections) {
    for (const sub of sec.subsections) {
      s2.push(h2(sub.subheading));
      for (const p of sub.paragraphs) s2.push(bodyPara(p));
    }
  }
  if (sec.paragraphs) {
    for (const p of sec.paragraphs) s2.push(bodyPara(p));
  }
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT_SONG, size: SZ21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: SZ28, bold: true, font: FONT_FSONG },
        paragraph: { spacing: { before: 200, after: 200 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: SZ21, bold: true, font: FONT_HEI },
        paragraph: { spacing: { before: 60, after: 60 }, outlineLevel: 1 } },
    ],
  },
  sections: [
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: MARGINS } },
      headers: { default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: '哈尔滨工程大学学报', size: SZ18, font: FONT_SONG, color: '808080' })],
      })] }) },
      footers: { default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: '— ', size: SZ18, font: FONT_SONG }),
          new TextRun({ children: [PageNumber.CURRENT], size: SZ18, font: FONT_SONG }),
          new TextRun({ text: ' —', size: SZ18, font: FONT_SONG }),
        ],
      })] }) },
      children: s1,
    },
    { properties: {
        type: SectionType.CONTINUOUS,
        page: { size: { width: 11906, height: 16838 }, margin: MARGINS },
        column: { count: 2, space: 720, equalWidth: true },
      },
      children: s2,
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('RIS_6G_Report.docx', buf);
  console.log('Done!');
});
