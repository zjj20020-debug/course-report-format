# course-report-format

中文课程报告/学术论文格式（哈尔滨工程大学学报模板）。

## 依赖

```bash
pip install lxml python-docx
winget install pandoc
```

## 用法

```bash
python generate_report.py report_content.json
```

输出 `RIS_6G_Report.docx`。

## 功能

- 中英双语题头（字号/字体/间距自动匹配）
- LaTeX 公式 → Word 原生 OMML（pandoc --mathml）
- 三线表 + 中英双语图注
- 参考文献中英对照
- 题头单栏 → 正文双栏排版
