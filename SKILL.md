---
name: course-report-format
description: 中文课程报告/学术论文格式规范（基于《哈尔滨工程大学学报》模板）。当用户要求撰写课程报告、结课论文、学术论文、学期大作业报告时调用此技能。包括页面布局、字体字号对照、双栏排版、参考文献格式、公式渲染等完整规范。也适用于"把报告写成论文格式"、"调整成学报格式"、"按模板排版"、"公式要用Word能编辑的格式"等请求。
---

# 中文课程报告格式规范

此技能定义生成 `.docx` 格式课程报告/学术论文的完整格式。**公式使用 pandoc + MathML 预渲染为 Word 原生 OMML 公式**，效果与在 Word 中按 `Alt+=` 输入 LaTeX 转换一致。

## 依赖

- **pandoc** (≥3.x) — 公式渲染。安装: `winget install pandoc`
- **Node.js** + **docx** (npm) — 文档结构生成。安装: `npm install docx`
- **Python 3** + **lxml** + **python-docx** — 文档合并后处理

## 完整工作流

```
report_content.json (含 LaTeX 公式)
       │
       ▼
  Markdown 生成 ─────────────── 将 JSON 内容转成 Markdown
       │                         保留 $$...$$ 公式原样
       ▼
  pandoc --mathml ───────────── 将 LaTeX 公式渲染为 OMML（Word原生公式）
       │
       ▼
  docx-js 结构生成 ──────────── 创建文档骨架（页边距、分栏、字体字号）
       │
       ▼
  XML 合并 ──────────────────── 将 pandoc 正文（含OMML公式）合并入结构文档
       │
       ▼
  RIS_6G_Report.docx ────────── 最终输出
```

运行方式（推荐）：
```bash
cd 项目目录
python generate_report.py
```

## 页面设置
- **纸张**：A4（宽 11906 DXA，高 16838 DXA）
- **页边距**：上 2.6cm（1474 DXA）、下 1.6cm（907 DXA）、左 2.0cm（1134 DXA）、右 2.0cm（1134 DXA）

## 字体字号对照表

| 元素 | 字体 | 字号 | 其他格式 |
|------|------|------|---------|
| 论文题目 | 黑体 | 二号（22pt，44 半分） | 居中，段前60磅段后15磅，单倍行距 |
| 作者 | 楷体 | 小四（12pt，24 半分） | 居中，固定行距15磅，段后6磅 |
| 单位 | 楷体 | 小五（9pt，18 半分） | 居中，固定行距15磅，左右缩进2字符 |
| 摘要标签「摘  要：」 | 黑体加粗 | 小五（9pt） | — |
| 摘要内容 | 楷体 | 小五（9pt） | 固定行距15磅，左右缩进2字符 |
| 关键词标签 | 黑体加粗 | 小五（9pt） | — |
| 关键词内容 | 宋体 | 小五（9pt） | 固定行距15磅 |
| **正文** | **宋体** | **五号（10.5pt，21 半分）** | **首行缩进2字符，单倍行距，两端对齐** |
| **一级标题** | **仿宋（加粗）** | **四号（14pt，28 半分）** | **固定行距15磅，段前段后10磅** |
| **二级标题** | **黑体（加粗）** | **五号（10.5pt，21 半分）** | **固定行距15磅，段前段后3磅** |
| 参考文献标题 | 黑体（加粗） | 四号（14pt） | 单倍行距 |
| 参考文献正文 | 中文宋体 / 英文 Times New Roman | 小五（9pt） | 固定行距12磅 |

## 版式布局（分栏）

将文档分为 **两个 Section**，通过 `SectionType.CONTINUOUS` 接续在同一页：

### Section 1：单栏
依次包含：论文题目 → 作者 → 单位 → 摘要 → 关键词

### Section 2：双栏
- 栏数：2，等宽，栏间距 0.5英寸
- **无**分隔线
- 从单栏后直接接续，**不分页**
- 包含所有正文章节及参考文献

## 公式格式

### 公式渲染方式
公式使用 **pandoc + --mathml** 预渲染为 Word 原生 OMML 格式。这种方式生成的公式：
- 使用 Cambria Math 字体，Word 公式引擎专业排版
- 变量自动斜体、函数名自动正体
- 可在 Word 中双击编辑（与 Alt+= 插入的公式完全一致）

### 在 JSON 中写公式
在 `report_content.json` 中使用 `$...$`（行内）或 `$$...$$`（单独一行）写 LaTeX 公式：

| 写法 | 效果 | 用途 |
|------|------|------|
| `$N$`, `$\\varphi_n$`, `$H_{\\mathrm{RB}}$` | **行内公式**（与正文同段落） | 数学变量、参数、符号 |
| `$$\\Gamma_n = \\beta_n \\cdot \\exp(j\\varphi_n)$$` | **单独一行居中** | 完整公式表达式 |

```json
{"paragraphs": [
    "系统有 $N$ 个单元，第 $n$ 个单元的反射系数为：$$\\Gamma_n = \\beta_n \\cdot \\exp(j\\varphi_n)$$其中 $\\beta_n$ 为幅度响应，$\\varphi_n$ 为可调相位。"
]}
```

**注意**：JSON 中的反斜杠必须双写（`\\Gamma` 而非 `\Gamma`）。

### 常用 LaTeX 命令
| 含义 | LaTeX |
|------|-------|
| Gamma | `\\Gamma` |
| beta | `\\beta` |
| phi | `\\varphi` |
| Phi | `\\Phi` |
| epsilon | `\\varepsilon` |
| mu | `\\mu` |
| 下标 | `x_{n}` |
| 点乘 | `\\cdot` |
| 函数名 | `\\mathrm{diag}`, `\\exp` |
| 省略号 | `\\ldots` |

## 参考文献格式

### 标题
四号黑体（14pt），单倍行距（与其他一级标题不同，不用仿宋）。

### 正文
- 字号：小五（9pt）
- 行距：**固定值12磅**（`line: 240, lineRule: EXACT`）
- 英文引用：Times New Roman，中文引用：宋体

### 作者规则
- 姓氏在前，**全字母大写**；名在后，首字母大写
- 3人以上：列出前3人，中文加"等"，英文加"et al."

### 中英文对照
中文参考文献下方必须附英文翻译，不编号。JSON 结构：
```json
{"type": "cn", "cn": "中文原文...", "en": "English translation..."}
```

## JSON 数据结构

```json
{
  "title": "论文题目",
  "abstract": "摘要内容（约200字）",
  "keywords": "关键词1；关键词2；关键词3",
  "sections": [
    {
      "heading": "1 一级标题",
      "subsections": [
        {
          "subheading": "1.1 二级标题",
          "paragraphs": ["正文段落1（含 $...$ 公式）", "正文段落2"]
        }
      ]
    },
    {
      "heading": "参考文献",
      "refs": [
        {"type": "en", "text": "[1] 英文参考文献..."},
        {"type": "cn", "cn": "[N] 中文参考文献...", "en": "English translation..."}
      ]
    }
  ]
}
```

## 技能脚本说明

脚本目录 `scripts/` 包含以下文件：

| 文件 | 作用 |
|------|------|
| `generate_report.py` | **主入口**：完整流水线（项目目录运行） |
| `generate_structure.js` | docx-js 生成文档骨架（含分栏） |
| `create_reference_template.js` | 创建 pandoc `--reference-doc` 模板 |
| `heu_template.docx` | pandoc 参考模板（定义正文/标题样式） |

### 使用步骤

1. **创建 `report_content.json`**，填入报告内容
2. **复制 `generate_report.py`** 到项目根目录
3. **运行**：`python generate_report.py`
4. **输出**：`RIS_6G_Report.docx`

### 注意事项
- LaTeX 反斜杠在 JSON 中必须双写：`\\Gamma`
- 确保 pandoc 在 PATH 中可用
- 确保 `course-report-format` skill 的 `scripts/` 目录下有 `heu_template.docx` 和 `generate_structure.js`
- `node_modules` 在 skill 脚本目录中可保留（跨项目复用）
