# 投标佐证材料生成器

将产品彩页和检测报告自动整合为投标佐证材料 Word 文档。

## 架构

两阶段流水线，各司其职：

```
分项报价表.xlsx  +  docs/（产品彩页 & 检测报告）
        │
        ▼
  ┌─ 阶段 1: 编排（Python）─────────────┐
  │  orchestrate.py                     │
  │  • 读取 Excel B 列产品名            │
  │  • 批量转换 PDF/DOCX → JPG          │
  │  • 输出 _manifest.json 清单         │
  └──────────────┬──────────────────────┘
                 │
                 ▼
  ┌─ 阶段 2: 文档生成（C# / .NET）──────┐
  │  OpenXML SDK                        │
  │  • 自动编号标题（Heading1 样式）     │
  │  • 图片居中、独占分页               │
  │  • 宋体 4 号黑色标题                │
  │  • 输出 投标佐证材料.docx            │
  └─────────────────────────────────────┘
```

| 阶段 | 语言 | 职责 | 为什么 |
|------|------|------|--------|
| 编排 | Python | 文件遍历、格式转换（PDF→JPG、DOCX→PDF→JPG） | PyMuPDF 渲染 PDF 简单高效；LibreOffice 命令行调用便捷 |
| 文档生成 | C#/.NET | 结构化排版、样式继承、编号绑定 | OpenXML SDK 是微软官方库，处理复杂文档结构远超 Python-docx |

## 子技能依赖

本技能依赖以下子技能，需与工作目录并列安装：

| 子技能 | 路径 | 职责 | 阶段 |
|--------|------|------|------|
| **minimax-docx** | `~/.workbuddy/skills/minimax-docx/` | 创建最终 DOCX 文档（样式、编号、图片排版） | 阶段 2 |
| **minimax-xlsx** | `~/.workbuddy/skills/minimax-xlsx/` | Excel 读写参考（本技能脚本直接调用 pandas/openpyxl 读取分项报价表） | 阶段 1 |
| **pdf** | `~/.workbuddy/skills/pdf/` | PDF 处理参考（本技能脚本直接调用 PyMuPDF） | 阶段 1 |
| **docx** | `~/.workbuddy/skills/docx/` | Word→图片转换参考（soffice → PDF → JPG 流水线） | 阶段 1 |

> **注意**：阶段 1 由 `scripts/orchestrate.py` 独立完成，`minimax-xlsx`、`pdf` 和 `docx` 子技能仅为参考文档，脚本自身包含所有转换逻辑。阶段 2 必须依赖 `minimax-docx` 完成 DOCX 输出。

## 前置条件

### Python 依赖

```bash
pip install pandas openpyxl pymupdf comtypes
```

### 系统要求

| 工具 | 用途 | 安装 |
|------|------|------|
| **LibreOffice** | Word → PDF 转换（首选） | [libreoffice.org](https://www.libreoffice.org/) |
| **pdftoppm**（可选） | PDF → JPG 渲染（优选） | `choco install poppler`（Windows） |
| **Microsoft Word**（可选） | Word → PDF 回退方案 | Office 自备 |

> 以上均为建议项。脚本内置多重回退链：
> Word → PDF：LibreOffice → Word COM
> PDF → JPG：pdftoppm → PyMuPDF

## 工作目录结构

用户须按如下结构准备素材：

```
{工作目录}/
├── 分项报价表.xlsx          # 必须。B 列为"分项名称"
└── docs/
    └── {产品名称}/          # 目录名须与 B 列分项名称完全一致
        ├── 彩页/            # 可选。PDF、Word、图片
        │   ├── brochure.pdf
        │   ├── spec.docx
        │   └── photo.jpg
        └── 检测报告/        # 可选。PDF、Word、图片
            ├── report.pdf
            └── cert.png
```

## 快速开始

```bash
# 1. 准备素材目录（按上述结构）
# 2. 运行编排脚本
python scripts/orchestrate.py /path/to/workdir

# 3. 脚本输出：
#    _images/          — 转换后的图片
#    _manifest.json    — 产品-图片映射清单
#
# 4. 生成 DOCX（由 WorkBuddy / minimax-docx 技能完成）
# 输出：投标佐证材料.docx
```

## 文档规格

最终生成的 `投标佐证材料.docx` 遵循以下规范：

| 属性 | 值 |
|------|-----|
| 纸张 | A4（210mm × 297mm） |
| 页边距 | 上下左右各 1 英寸 |
| 标题字体 | 宋体 4 号（14pt）黑色加粗 |
| 标题编号 | 自动十进制递增（1. 2. 3. …） |
| 正文字体 | 宋体（通过 DocDefaults 继承） |
| 图片布局 | 居中，每张图片独占一页 |

## 容错

| 情况 | 处理 |
|------|------|
| B 列为空 | 脚本终止，提示检查 |
| 产品目录不存在 | 跳过该产品，继续处理其他 |
| `彩页/` 或 `检测报告/` 不存在 | 跳过该类别 |
| PDF/Word 转换失败 | 跳过单个文件，继续处理 |
| 所有产品无有效图片 | 不生成 DOCX |

## 依赖

| 库 | 用途 |
|----|------|
| pandas + openpyxl | 读取分项报价表 |
| PyMuPDF（fitz） | PDF 渲染为图片（回退方案） |
| comtypes | Windows Word COM 接口（回退方案） |
| DocumentFormat.OpenXml 3.5+ | Word 文档生成（阶段 2） |

## 许可

MIT
