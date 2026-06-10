---
name: course-report-format
description: |
  中文课程报告/学术论文格式规范（哈尔滨工程大学学报模板）。
  触发条件：用户要求写课程报告、结课论文、学期大作业、学术论文；用户要求"按学报格式排版"、"改成论文格式"、"套模板"；用户给了一个现有 docx 要求"改成模板格式"。
  此技能提供的是【格式排版流水线】，不是普通的格式说明。你必须运行脚本来生成报告。
---

# 中文课程报告格式

## ⚠️ 核心规则（必须遵守）

**你不应该手动拼 XML 或手写 docx。你必须使用本技能提供的脚本流水线来生成最终报告。**

## 工作流（按顺序执行）

### Step 1：提取/构建 report_content.json

在用户指定的工作目录下创建一个 `report_content.json`，结构如下：

```json
{
  "title": "论文题目",
  "abstract": "摘要内容（200字左右）",
  "keywords": "关键词1；关键词2；关键词3；...",
  "sections": [
    {
      "heading": "1 引言",
      "subsections": [
        {
          "subheading": "1.1 小节标题",
          "paragraphs": ["正文段落...", "正文段落..."]
        }
      ]
    },
    {
      "heading": "参考文献",
      "refs": [
        {"type": "en", "text": "[1] AUTHOR N. Title[J]. Journal, Year, Vol(Issue): Pages."},
        {"type": "cn", "cn": "[2] 作者. 中文标题[J]. 期刊, 年, 卷(期): 页码.", "en": "AUTHOR N. English title[J]. Journal, Year, Vol(Issue): Pages."}
      ]
    }
  ]
}
```

**写 JSON 的注意事项**：
- LaTeX 反斜杠双写，例如 `$\\Gamma_n$`、`$\\beta_n$`
- 中文引号 `""` 改成 `「」`，避免破坏 JSON 语法
- 正文段落是一个大字符串，多个段落拆成数组不同元素
- 公式用 `$...$`（行内）或 `$$...$$`（独立成行）

### Step 2：安装运行所需的依赖（如缺失）

以下两个路径的依赖是脚本运行的前提，缺失时先安装：

```bash
# 检查并安装 Node 依赖（在 skill 脚本目录）
cd ~/.claude/skills/course-report-format/scripts
npm install docx

# 检查 pandoc
pandoc --version || winget install pandoc

# 检查 lxml
python -c "import lxml" 2>/dev/null || pip install lxml
```

### Step 3：运行生成脚本

**直接从 skill 目录运行脚本**，不要拷贝脚本到工作目录：

```bash
python ~/.claude/skills/course-report-format/scripts/generate_report.py <工作目录>/report_content.json
```

或者先 cd 到工作目录然后运行（脚本会从 cwd 找 `report_content.json`）：

```bash
cd <工作目录> && python ~/.claude/skills/course-report-format/scripts/generate_report.py
```

脚本会自动处理：
- docx-js 生成结构骨架（页边距、分栏、题头）
- pandoc 将 Markdown 正文转为含 OMML 公式的 docx
- XML 合并两个部件
- 参考文献单独按正确格式追加

### Step 4：验证输出

检查 `RIS_6G_Report.docx` 是否生成成功，确认公式数量和章节结构正确。

## 格式规范参考

### 页面设置
- A4，上 2.6cm，下 1.6cm，左 2.0cm，右 2.0cm

### 字体字号

| 元素 | 字体 | 字号 |
|------|------|------|
| 论文题目 | 黑体 | 二号（22pt） |
| 作者 | 楷体 | 小四（12pt） |
| 单位 | 楷体 | 小五（9pt） |
| 摘要标签「摘  要：」 | 黑体加粗 | 小五（9pt） |
| 摘要内容 | 楷体 | 小五（9pt） |
| 关键词标签 | 黑体加粗 | 小五（9pt） |
| 关键词内容 | 宋体 | 小五（9pt） |
| **正文** | **宋体** | **五号（10.5pt）** |
| **一级标题** | **仿宋（加粗）** | **四号（14pt）** |
| **二级标题** | **黑体（加粗）** | **五号（10.5pt）** |
| 参考文献标题 | 黑体（加粗） | 四号（14pt） |
| 参考文献正文 | 中文宋体 / 英文TNR | 小五（9pt） |

- 正文：首行缩进2字符，单倍行距，两端对齐
- 一级标题：固定行距15磅，段前段后10磅
- 二级标题：固定行距15磅，段前段后3磅
- 参考文献：固定行距12磅
- 题头部分：单栏；正文部分：双栏（不分页连续）

### 公式
- `$...$` → 行内公式
- `$$...$$` → 独立成行居中
- 使用 pandoc + `--mathml` 渲染为 Word 原生可编辑公式

### 参考文献规则
- 英文文献：Times New Roman，中文文献：宋体
- 作者姓氏全大写在前，名字首字母大写在后
- 3人以上：中文"等"，英文"et al."
- 中文文献下方附英文翻译（不编号）

## 常见问题

- **JSON 报错**：检查是否有多余的中文引号 `""`、LaTeX 反斜杠是否双写
- **pandoc 报错**：确认 pandoc 版本 ≥ 3.x
- **Node 报错**：确认 skill 的 `scripts/` 目录下已执行 `npm install docx`
- **公式效果不对**：确认 pandoc 命令包含 `--mathml` 参数
