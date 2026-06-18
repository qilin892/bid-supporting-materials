---
name: bid-supporting-materials
description: "投标佐证材料生成器。从结构化工作目录中读取分项报价表（Excel B列分项名称），自动将产品彩页和检测报告（PDF/Word/图片）统一转换为图片，并整合生成一份投标佐证材料 Word 文档。触发词：投标佐证、佐证材料、产品彩页、检测报告、生成证明材料、投标文件佐证。"
agent_created: true
---

# 投标佐证材料生成器

将产品彩页和检测报告自动整合为投标佐证材料 Word 文档。

## 工作目录结构要求

用户须指定一个工作目录，结构如下：

```
{工作目录}/
├── 分项报价表.xlsx          # 必须。B列为"分项名称"
└── docs/
    └── {产品名称}/          # 目录名须与B列分项名称完全一致
        ├── 彩页/            # 可选。放 PDF/Word/图片
        │   ├── brochure.pdf
        │   ├── spec.docx
        │   └── photo.jpg
        └── 检测报告/        # 可选。放 PDF/Word/图片
            ├── report.pdf
            └── cert.png
```

## 前置条件

### Python 依赖

```bash
pip install pandas openpyxl pymupdf comtypes
```

- `pandas` + `openpyxl`：读取 Excel
- `pymupdf` (fitz)：PDF 渲染（回退方案，当 pdftoppm 不可用时）
- `comtypes`：Windows 上调用 Word COM 接口（Word→PDF 回退方案）

### 系统要求

- **LibreOffice**：Word 文档转 PDF 的首选方案（`soffice --headless --convert-to pdf`）
- **poppler-utils**（`pdftoppm`）：PDF 转 JPG 的首选方案（docx 技能推荐），安装方式：
  - Linux: `sudo apt-get install poppler-utils`
  - macOS: `brew install poppler`
  - Windows: `choco install poppler` 或下载 [poppler Windows 二进制包](http://blog.alivate.com.au/poppler-windows/)
- **Microsoft Word**（可选回退）：当 LibreOffice 不可用时的替代方案
- **minimax-docx 技能**：用于最终 DOCX 生成（需已安装且环境就绪）

## 工作流程

### 步骤 1：确认工作目录

向用户获取工作目录的绝对路径。验证以下条件：

1. 路径存在且为目录
2. `分项报价表.xlsx` 文件存在
3. `docs/` 目录存在

若任一条件不满足，提示用户检查并终止。

### 步骤 2：运行编排脚本

执行本技能目录下的编排脚本，自动完成文件发现和格式转换：

```bash
python {SKILL_DIR}/scripts/orchestrate.py "<工作目录路径>"
```

该脚本自动完成：

- 读取 `分项报价表.xlsx`，提取 B 列所有"分项名称"值（去空去重，保持顺序）
- 逐个匹配 `docs/{分项名称}/` 目录（匹配不到则跳过）
- 遍历 `彩页/` 和 `检测报告/` 子目录中的文件：
  - **PDF** → PyMuPDF 渲染为 100 DPI JPG（质量60，最小体积）
  - **Word (.docx/.doc)** → LibreOffice headless 导出 PDF → pdftoppm 渲染为 100 DPI JPG（docx 技能推荐方式，回退 PyMuPDF）
  - **图片 (.jpg/.png/.bmp/.gif/.tif)** → 直接复制
- 在 `_images/{产品名}/{类别}/` 下输出转换后的图片
- 生成 `_manifest.json` 清单文件，记录每个产品每个类别的图片路径
- 打印处理摘要（产品数、图片数、跳过/错误信息）

### 步骤 3：检查处理结果

阅读脚本输出的摘要，确认：

- 多少产品成功匹配、多少被跳过
- 是否有转换失败的 `[ERROR]` 信息
- 总图片数是否合理

使用 Read 工具查看 `_manifest.json` 确认清单内容正确。

**终止条件**：如果没有任何产品生成有效图片（`products` 数组为空），向用户报告原因并终止流程，不创建 DOCX。

### 步骤 4：生成 Word 文档

使用 **minimax-docx** 技能的 Pipeline A (CREATE) 创建 `投标佐证材料.docx`。

#### 4.1 加载 minimax-docx

先加载 minimax-docx 技能文档，然后读取以下关键参考文件：

| 参考文件 | 用途 |
|----------|------|
| `Samples/ImageSamples.cs` | 图片插入：`BuildDrawingElement()`、`CalculateImageDimensions()`、`GetImagePartType()` |
| `Samples/DocumentCreationSamples.cs` | 文档创建：`CreateFullDocument()` 模式 |
| `Samples/StyleSystemSamples.cs` | 样式系统：`SetupDocDefaults()`、`CreateBasicStyles()`、`CreateListStyle()`（样式绑定编号） |
| `Samples/ListAndNumberingSamples.cs` | 编号系统：`SetupAbstractNum()`、`SetupNumberingInstance()`、`CreateDecimalLevel()` |

#### 4.2 读取清单

读取 `{工作目录}/_manifest.json`，获取产品列表和图片路径。

#### 4.3 构建 DOCX

编写 C# 程序（参照 `CreateFullDocument` 模式），关键要求：

**文档设置**：
- A4 纸张：`PageSize Width=11906 Height=16838` (DXA)
- 标准边距：上下左右各 1440 DXA（1英寸）
- CJK 默认字体：正文宋体(SimSun)，标题宋体(SimSun) 4号(14pt) 黑色
  - 标题字体设置（在 Heading1 样式的 StyleRunProperties 中）：
    - `RunFonts EastAsia = "SimSun"`  // 宋体
    - `FontSize Val = "28"`           // 4号 = 14pt → 28 half-points
    - `Color Val = "000000"`          // 黑色
    - `Bold`                          // 加粗

**自动编号设置**（关键：标题编号由 Word 自动生成，不手动填写数字）：

1. 创建 `NumberingDefinitionsPart`，定义一个十进制编号格式：
   ```
   // AbstractNum（abstractNumId = 扫描所有已有 abstractNumId 取 max+1，避免 ID 冲突）
   //   Level 0: NumberFormat = Decimal, LevelText = "%1.", StartVal = 1
   //   缩进: indentLeft = 0, hanging = 0 (编号不缩进，与标题文字对齐)
   // NumberingInstance（numId = 扫描所有已有 numId 取 max+1）→ references new abstractNumId
   ```
   参照 `ListAndNumberingSamples.cs` 的 `SetupAbstractNum()` 和 `SetupNumberingInstance()` 方法。
   **重要**：必须动态分配 abstractNumId 和 numId，扫描 numbering.xml 中已有 ID 取最大值+1。
   文档模板中可能预设了编号定义（如列表符号等），固定 ID 会导致冲突。

2. 在 Heading1 样式定义中绑定编号（参照 `StyleSystemSamples.cs` 的 `CreateListStyle()` 方法）：
   ```csharp
   // 在 Heading1 的 StyleParagraphProperties 中添加:
   // numId 使用上一步动态分配的值
   new NumberingProperties(
       new NumberingId { Val = <动态分配的numId> },        // 指向新增的 NumberingInstance
       new NumberingLevelReference { Val = 0 }             // 使用 Level 0
   )
   ```
   这样所有使用 Heading1 样式的段落都会自动获得递增编号（1.、2.、3. ...）。

3. **重要规则**：
   - AbstractNum 元素必须在 NumberingInstance 之前出现在 Numbering 根元素中
   - Heading1 样式中的 `numId` 必须与 NumberingInstance 的 `numId` 一致
   - 标题文字中**不要手动写编号**，只写"{产品名称} {类别}"，编号由 Word 自动生成

**文档结构**（按 `_manifest.json` 中产品顺序）：

```
对每个 product in manifest.products:
    对每个 category in ["彩页", "检测报告"]:  (按此固定顺序)
        if category in product.categories:
            1. 插入标题段落: "{product.name} {category}"
               - 使用 Heading1 样式（已绑定自动编号，Word 自动显示 "1."、"2." 等编号）
               - 不在文字中手动添加编号
               - 设置 PageBreakBefore 属性（除文档第一个标题外）
            2. 对每张图片:
               a. 调用 ImageSamples.CalculateImageDimensions(imagePath, 6.0)
                  获取适配 A4 页面宽度（最大6英寸）的 EMU 尺寸
               b. 添加 ImagePart 到 MainDocumentPart
               c. 获取 relId = mainPart.GetIdOfPart(imagePart)
               d. 调用 ImageSamples.BuildDrawingElement(relId, cx, cy, docPropId, name, desc)
                  - docPropId 从 1 开始递增，确保全文档唯一
               e. 创建 Paragraph，设置居中（`Justification.Val = Center`），包裹 Run(Drawing)，添加到 Body
               f. 在每张图片的段落上设置 PageBreakBefore = true
                  确保每张图片独占一页
```

**最终文档效果**：
```
1. 产品A 彩页          ← 编号自动生成，宋体4号黑色
   [图片居中]          ← 独占一页，居中对齐
   [图片居中]          ← 独占一页，居中对齐
2. 产品A 检测报告       ← 编号自动生成
...
```

**关键注意事项**：
- `SectionProperties` 必须是 Body 的最后一个子元素
- 每个插入操作使用 `body.AppendChild()` 或在 sectPr 之前 `body.InsertBefore()`
- `DocProperties.Id` 在整个文档中必须唯一（用递增计数器）
- `ImagePart` 必须添加到 `MainDocumentPart`（不是 Header/Footer）
- 图片路径使用 `_manifest.json` 中的绝对路径
- **编号定义顺序**：在 `Numbering` 根元素中，所有 `AbstractNum` 必须在 `NumberingInstance` 之前，否则 Word 报告文档损坏
- **编号与样式绑定**：编号通过 Heading1 样式的 `NumberingProperties` 绑定，不要在标题段落上直接添加 `NumberingProperties`，避免与样式中的重复冲突
- **标题文字不含编号**：标题文字只写"{产品名称} {类别}"，编号由 Word 根据 NumberingInstance 自动递增生成

#### 4.4 编译并运行

```bash
# 在 minimax-docx 的 dotnet 项目目录下
dotnet run --project scripts/dotnet/MiniMaxAIDocx.Cli -- create-bid-doc --manifest "<工作目录>/_manifest.json" --output "<工作目录>/投标佐证材料.docx"
```

如果 CLI 没有现成的 `create-bid-doc` 命令，则编写独立的 C# 脚本文件并编译运行，或直接用 `dotnet script` 执行。

### 步骤 5：验证输出

1. 确认 `{工作目录}/投标佐证材料.docx` 文件已生成
2. 使用 minimax-docx 的验证功能检查文档完整性：
   ```bash
   dotnet run --project scripts/dotnet/MiniMaxAIDocx.Cli -- validate --input "<工作目录>/投标佐证材料.docx"
   ```
3. 若验证失败，根据错误信息修复 C# 代码并重新生成

### 步骤 6：报告结果

向用户报告：
- 成功生成的产品数和总图片数
- 跳过的产品列表及原因（从 `_manifest.json` 的 summary 读取）
- 转换失败的文件（从脚本输出日志的 `[ERROR]` 行读取）
- 输出文件路径：`{工作目录}/投标佐证材料.docx`
- 清理建议：`_images/` 和 `_manifest.json` 为中间产物，可按需保留或删除

## 边界情况处理

| 情况 | 处理方式 |
|------|----------|
| `分项报价表.xlsx` 不存在 | 脚本终止，提示用户检查文件 |
| B列为空或无数据 | 脚本终止，提示 B列无数据 |
| `docs/` 目录不存在 | 脚本终止，提示创建目录结构 |
| 产品目录不存在 | 跳过该产品 `[SKIP]`，继续处理其他 |
| `彩页/` 或 `检测报告/` 子目录不存在 | 跳过该类别，不影响其他类别 |
| 子目录为空 | 跳过该类别 |
| PDF 转换失败 | 跳过该文件 `[ERROR]`，继续处理 |
| Word 转换失败（COM + LibreOffice 均失败） | 跳过该文件 `[ERROR]`，继续处理 |
| 产品无任何有效图片 | 不出现在 DOCX 中 |
| 全部产品无有效图片 | 不生成 DOCX，向用户报告 |
| 中文文件名 | 脚本内部使用 UTF-8；Word COM 转换时复制到临时英文路径 |

## 子技能依赖

| 子技能 | 路径 | 用途 |
|--------|------|------|
| minimax-docx | `~/.workbuddy/skills/minimax-docx/` | 创建最终 DOCX 文档 |
| minimax-xlsx | `~/.workbuddy/skills/minimax-xlsx/` | Excel 读写参考（本技能脚本直接使用 pandas/openpyxl 读取分项报价表） |
| pdf | `~/.workbuddy/skills/pdf/` | PDF 处理参考（本技能脚本直接使用 PyMuPDF） |
| docx | `~/.workbuddy/skills/docx/` | Word→图片转换（soffice → PDF → pdftoppm/JPG 方式） |

## 文件说明

| 文件 | 说明 |
|------|------|
| `scripts/orchestrate.py` | 编排脚本：读取Excel、遍历目录、转换文件、输出清单 |
| `_manifest.json`（运行时生成） | 图片清单文件，记录产品→类别→图片路径的映射 |
| `_images/`（运行时生成） | 转换后的图片存放目录 |
| `投标佐证材料.docx`（运行时生成） | 最终输出的投标佐证文档 |
