# course-report-format

中文课程报告/学术论文格式规范（基于《哈尔滨工程大学学报》模板）。

自动生成 `.docx` 格式课程报告，**公式使用 pandoc + MathML 预渲染为 Word 原生 OMML**，与在 Word 中 `Alt+=` 输入 LaTeX 转换效果一致。

## 安装

将 `.skill` 文件拖入 Claude Code 即可安装。

或手动复制本目录到 `~/.claude/skills/`。

## 依赖

- **pandoc** (≥3.x) — LaTeX 公式渲染为 Word OMML
- **Node.js** + **docx** (npm) — 文档结构生成
- **Python 3** + **lxml** + **python-docx** — 后处理

```bash
winget install pandoc
npm install -g docx
pip install lxml python-docx
```

## 使用

1. 创建 `report_content.json`，填入报告内容（LaTeX 公式用 `$...$` / `$$...$$`）
2. 运行 `python generate_report.py`
3. 输出 `RIS_6G_Report.docx`

## 格式规范

| 元素 | 字体 | 字号 |
|------|------|------|
| 论文题目 | 黑体 | 二号 22pt |
| 正文 | 宋体 | 五号 10.5pt |
| 一级标题 | 仿宋 | 四号 14pt |
| 二级标题 | 黑体 | 五号 10.5pt |
| 摘要 | 楷体 | 小五 9pt |

- 题头单栏 → 正文双栏（连续不分页）
- 参考文献：英 TNR / 中宋体，小五 9pt，固定行距 12pt
