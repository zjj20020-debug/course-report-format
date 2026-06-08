/**
 * 生成 pandoc 的 --reference-docx 模板
 * 包含 HEU 学报的 Normal、Heading 1、Heading 2 等样式定义
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  HeadingLevel, LineRuleType,
} = require('docx');

const SONG  = '宋体';
const HEI   = '黑体';
const FSONG = '仿宋';
const TNR   = 'Times New Roman';

// 固定行距15磅
const EXACT15 = { line: 300, lineRule: LineRuleType.EXACT };

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: SONG, size: 21 } },  // 五号
      paragraph: { spacing: { after: 0, line: 240 }, alignment: AlignmentType.JUSTIFIED },
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'heading 1', basedOn: 'Normal', next: 'Normal',
        quickFormat: true,
        run: { size: 28, bold: true, font: FSONG },  // 四号仿宋
        paragraph: { spacing: { before: 200, after: 200, ...EXACT15 }, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'heading 2', basedOn: 'Normal', next: 'Normal',
        quickFormat: true,
        run: { size: 21, bold: true, font: HEI },     // 五号黑体
        paragraph: { spacing: { before: 60, after: 60, ...EXACT15 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1474, bottom: 907, left: 1134, right: 1134 },
      },
    },
    children: [
      new Paragraph({ children: [new TextRun('Normal text - 正文五号宋体')] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Heading 1 - 四号仿宋')] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('Heading 2 - 五号黑体')] }),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('C:/Users/ZJLJL/.claude/skills/course-report-format/scripts/heu_template.docx', buf);
  console.log('Template created!');
});
