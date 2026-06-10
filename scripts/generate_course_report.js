#!/usr/bin/env node
/**
 * 生成文档结构骨架
 * 用法: node generate_structure.js <args.json>
 * args.json: { "title": "...", "abstract": "...", "keywords": "...", "output": "..." }
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun,
  Header, Footer, AlignmentType,
  HeadingLevel, LineRuleType, SectionType, PageNumber,
} = require('docx');

const args = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const { title, abst, keywords, output } = args;

const SONG  = '宋体';
const HEI   = '黑体';
const KAI   = '楷体';
const FSONG = '仿宋';

const EXACT15 = { line: 300, lineRule: LineRuleType.EXACT };
const SINGLE  = 240;
const INDENT2 = 420;

const s1 = [];

s1.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 1200, after: 300, line: SINGLE },
  children: [new TextRun({ text: title, size: 44, bold: true, font: HEI })],
}));
s1.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120, ...EXACT15 },
  children: [new TextRun({ text: '作者姓名', size: 24, font: KAI })],
}));
s1.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200, ...EXACT15 },
  indent: { left: INDENT2, right: INDENT2 },
  children: [new TextRun({
    text: '(哈尔滨工程大学 信息与通信工程学院，黑龙江 哈尔滨 150001)',
    size: 18, font: KAI,
  })],
}));
s1.push(new Paragraph({
  spacing: { after: 0, ...EXACT15 },
  alignment: AlignmentType.JUSTIFIED, indent: { left: INDENT2, right: INDENT2 },
  children: [
    new TextRun({ text: '摘  要：', size: 18, bold: true, font: HEI }),
    new TextRun({ text: abst, size: 18, font: KAI }),
  ],
}));
s1.push(new Paragraph({
  spacing: { after: 200, ...EXACT15 },
  alignment: AlignmentType.JUSTIFIED, indent: { left: INDENT2, right: INDENT2 },
  children: [
    new TextRun({ text: '关键词：', size: 18, bold: true, font: HEI }),
    new TextRun({ text: keywords, size: 18, font: SONG }),
  ],
}));

// Section 2 placeholder
const s2 = [
  new Paragraph({ children: [new TextRun({ text: '', size: 21, font: SONG })] }),
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: SONG, size: 21 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: FSONG },
        paragraph: { spacing: { before: 200, after: 200 }, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 21, bold: true, font: HEI },
        paragraph: { spacing: { before: 60, after: 60 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: { size: { width: 11906, height: 16838 },
                margin: { top: 1474, bottom: 907, left: 1134, right: 1134 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: '哈尔滨工程大学学报', size: 16, font: SONG, color: '808080' })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: '-- ', size: 16, font: SONG }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, font: SONG }),
              new TextRun({ text: ' --', size: 16, font: SONG }),
            ],
          })],
        }),
      },
      children: s1,
    },
    {
      properties: {
        type: SectionType.CONTINUOUS,
        page: { size: { width: 11906, height: 16838 },
                margin: { top: 1474, bottom: 907, left: 1134, right: 1134 } },
        column: { count: 2, space: 720, equalWidth: true },
      },
      children: s2,
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(output, buf);
  console.log('Structure written');
});
