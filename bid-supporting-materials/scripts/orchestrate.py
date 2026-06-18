"""
orchestrate.py - 投标佐证材料编排脚本

功能：
  1. 读取 分项报价表.xlsx，提取 B 列"分项名称"
  2. 遍历 docs/产品名/彩页 和 docs/产品名/检测报告
  3. 将 PDF、Word 文件转换为 JPG 图片；已有图片直接复制
  4. 输出 _manifest.json 清单文件供后续 DOCX 生成使用

依赖：pip install pandas openpyxl pymupdf comtypes
"""

import os
import sys
import json
import shutil
import platform
import subprocess
import tempfile

# ── 编码 ──
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ── 常量 ──
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}
CATEGORIES = ["彩页", "检测报告"]
PDF_DPI = 100           # PDF 渲染分辨率（最小可接受）
JPG_QUALITY = 60        # JPG 输出质量（1-100，越小体积越小）


def log(level, msg):
    """带级别的日志输出"""
    print(f"[{level}] {msg}")


# ═══════════════════════════════════════════════════════════════════
# 1. Excel 读取
# ═══════════════════════════════════════════════════════════════════

def read_product_names(excel_path):
    """
    读取分项报价表.xlsx，返回 B 列（第2列）的分项名称列表。
    去除空值和重复项，但保持原始顺序。
    """
    try:
        import pandas as pd
    except ImportError:
        log("ERROR", "缺少 pandas，请运行: pip install pandas openpyxl")
        sys.exit(1)

    df = pd.read_excel(excel_path, engine="openpyxl")
    if df.shape[1] < 2:
        log("ERROR", "分项报价表列数不足，需要至少2列（B列为分项名称）")
        sys.exit(1)

    # B列 = 第2列（索引1）
    col_b = df.iloc[:, 1]
    names = []
    seen = set()
    for val in colb_dropna(col_b):
        name = str(val).strip()
        if name and name not in seen:
            names.append(name)
            seen.add(name)
    return names


def colb_dropna(series):
    """去除 NaN/None 值"""
    return series.dropna()


# ═══════════════════════════════════════════════════════════════════
# 2. PDF → 图片
# ═══════════════════════════════════════════════════════════════════

def convert_pdf_to_images(pdf_path, output_dir, name_prefix=None):
    """
    使用 PyMuPDF (fitz) 将 PDF 每页转为 JPG 图片。
    返回生成的图片路径列表。
    """
    import fitz  # PyMuPDF

    if name_prefix is None:
        name_prefix = os.path.splitext(os.path.basename(pdf_path))[0]

    doc = fitz.open(pdf_path)
    images = []
    zoom = PDF_DPI / 72.0  # 72 DPI 是 PDF 默认，缩放到目标 DPI
    mat = fitz.Matrix(zoom, zoom)

    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=mat)
        out_name = f"{name_prefix}_p{page_num + 1:03d}.jpg"
        out_path = os.path.join(output_dir, out_name)
        pix.save(out_path, output="jpeg", jpg_quality=JPG_QUALITY)
        images.append(os.path.abspath(out_path))
        log("OK", f"  PDF 页 {page_num + 1}/{len(doc)} → {out_name}")

    doc.close()
    return images


# ═══════════════════════════════════════════════════════════════════
# 3. Word → PDF → 图片
# ═══════════════════════════════════════════════════════════════════

def word_to_pdf_com(word_path, pdf_path):
    """
    使用 Word COM 接口将 Word 文档导出为 PDF（Windows + Microsoft Word）。
    返回 True/False。
    """
    try:
        import comtypes.client
    except ImportError:
        log("WARN", "  comtypes 未安装，无法使用 Word COM 接口")
        return False

    # 复制到临时英文路径避免中文路径问题
    temp_dir = tempfile.mkdtemp(prefix="word2pdf_")
    temp_word = os.path.join(temp_dir, "source.docx")
    try:
        shutil.copy2(word_path, temp_word)
    except Exception as e:
        log("ERROR", f"  复制文件失败: {e}")
        shutil.rmtree(temp_dir, ignore_errors=True)
        return False

    word = None
    doc = None
    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(temp_word, ReadOnly=True)
        doc.SaveAs2(pdf_path, 17)  # 17 = wdFormatPDF
        return True
    except Exception as e:
        log("ERROR", f"  Word COM 转换失败: {e}")
        return False
    finally:
        try:
            if doc:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass
        shutil.rmtree(temp_dir, ignore_errors=True)


def _find_soffice():
    """查找 soffice 可执行文件路径"""
    if platform.system() == "Windows":
        candidates = [
            os.path.expandvars(r"%PROGRAMFILES%\LibreOffice\program\soffice.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\LibreOffice\program\soffice.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
            "soffice.exe",
        ]
    else:
        candidates = [
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "soffice",
        ]
    for c in candidates:
        if os.path.isfile(c) or (c in ("soffice", "soffice.exe")):
            # 用 shutil.which 检查 PATH 中的命令
            found = shutil.which(c)
            if found:
                return found
            continue
    return None


def word_to_pdf_libreoffice(word_path, pdf_path):
    """
    使用 LibreOffice headless 模式将 Word 转为 PDF。
    这是 docx 技能推荐的 Word 文档转换方式。
    """
    soffice = _find_soffice()
    if soffice is None:
        log("ERROR", "  LibreOffice (soffice) 未找到")
        return False

    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", os.path.dirname(pdf_path), word_path],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0:
            # LibreOffice 输出文件名 = 输入文件名 + .pdf
            expected = os.path.splitext(os.path.basename(word_path))[0] + ".pdf"
            expected_path = os.path.join(os.path.dirname(pdf_path), expected)
            if os.path.exists(expected_path):
                if expected_path != pdf_path:
                    shutil.move(expected_path, pdf_path)
                return True
        log("ERROR", f"  LibreOffice 转换失败: {result.stderr}")
        return False
    except Exception as e:
        log("ERROR", f"  LibreOffice 转换异常: {e}")
        return False


def pdf_to_jpg_pdftoppm(pdf_path, output_dir, name_prefix=None, dpi=100):
    """
    使用 pdftoppm 将 PDF 每页转为 JPG 图片。
    这是 docx 技能推荐的方式（soffice → pdftoppm）。
    返回生成的图片路径列表。
    """
    if name_prefix is None:
        name_prefix = os.path.splitext(os.path.basename(pdf_path))[0]

    prefix_path = os.path.join(output_dir, name_prefix)
    try:
        result = subprocess.run(
            ["pdftoppm", "-jpeg", f"-r", str(dpi), pdf_path, prefix_path],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode != 0:
            log("ERROR", f"  pdftoppm 失败: {result.stderr}")
            return []

        # pdftoppm 输出: prefix-1.jpg, prefix-2.jpg, ...
        images = []
        i = 1
        while True:
            expected = f"{prefix_path}-{i}.jpg"
            if os.path.isfile(expected):
                images.append(os.path.abspath(expected))
                i += 1
            else:
                break

        if not images:
            log("ERROR", "  pdftoppm 未生成图片文件")
            return []

        log("OK", f"  pdftoppm 生成 {len(images)} 张图片 (100 DPI)")
        return images

    except FileNotFoundError:
        log("WARN", "  pdftoppm 未安装")
        return []
    except Exception as e:
        log("ERROR", f"  pdftoppm 异常: {e}")
        return []


def convert_word_to_images(word_path, output_dir, name_prefix=None):
    """
    将 Word 文档转换为 JPG 图片。
    按 docx 技能推荐方式：soffice --headless → PDF → pdftoppm → JPG。
    回退链：pdftoppm → PyMuPDF（若 pdftoppm 不可用）。
    Word→PDF 回退链：LibreOffice → Word COM。
    """
    if name_prefix is None:
        name_prefix = os.path.splitext(os.path.basename(word_path))[0]

    temp_pdf = os.path.join(output_dir, f"_temp_{name_prefix}.pdf")

    # Step 1: Word → PDF（LibreOffice 优先，Word COM 回退）
    log("INFO", "  使用 LibreOffice 转换 Word → PDF...")
    success = word_to_pdf_libreoffice(word_path, temp_pdf)
    if not success:
        log("INFO", "  尝试 Word COM 回退...")
        success = word_to_pdf_com(word_path, temp_pdf)

    if not success:
        log("ERROR", f"  Word 文档转换失败: {word_path}")
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)
        return []

    # Step 2: PDF → JPG（pdftoppm 优先，PyMuPDF 回退）
    images = pdf_to_jpg_pdftoppm(temp_pdf, output_dir, name_prefix=name_prefix)
    if not images:
        log("INFO", "  回退到 PyMuPDF 渲染...")
        images = convert_pdf_to_images(temp_pdf, output_dir, name_prefix=name_prefix)

    # 清理临时 PDF
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)

    return images


# ═══════════════════════════════════════════════════════════════════
# 4. 图片文件复制
# ═══════════════════════════════════════════════════════════════════

def copy_image_file(src_path, output_dir):
    """直接复制图片文件到输出目录，返回目标路径。"""
    filename = os.path.basename(src_path)
    dest = os.path.join(output_dir, filename)
    # 处理同名冲突
    if os.path.exists(dest) and os.path.abspath(src_path) != os.path.abspath(dest):
        base, ext = os.path.splitext(filename)
        i = 1
        while os.path.exists(dest):
            dest = os.path.join(output_dir, f"{base}_{i}{ext}")
            i += 1
    shutil.copy2(src_path, dest)
    log("OK", f"  图片复制 → {os.path.basename(dest)}")
    return os.path.abspath(dest)


# ═══════════════════════════════════════════════════════════════════
# 5. 单文件处理
# ═══════════════════════════════════════════════════════════════════

def process_file(filepath, output_dir):
    """
    根据文件扩展名处理单个文件，返回图片路径列表。
    """
    ext = os.path.splitext(filepath)[1].lower()
    name_prefix = os.path.splitext(os.path.basename(filepath))[0]

    try:
        if ext == ".pdf":
            log("INFO", f"  转换 PDF: {os.path.basename(filepath)}")
            return convert_pdf_to_images(filepath, output_dir, name_prefix)

        elif ext in (".docx", ".doc"):
            log("INFO", f"  转换 Word: {os.path.basename(filepath)}")
            return convert_word_to_images(filepath, output_dir, name_prefix)

        elif ext in IMAGE_EXTS:
            return [copy_image_file(filepath, output_dir)]

        else:
            log("WARN", f"  不支持的格式，跳过: {os.path.basename(filepath)}")
            return []

    except Exception as e:
        log("ERROR", f"  处理文件失败 {os.path.basename(filepath)}: {e}")
        return []


# ═══════════════════════════════════════════════════════════════════
# 6. 主编排逻辑
# ═══════════════════════════════════════════════════════════════════

def orchestrate(working_dir):
    """
    主编排函数：
      1. 验证工作目录
      2. 读取分项报价表
      3. 遍历产品目录，转换文件
      4. 输出 _manifest.json
    """
    working_dir = os.path.abspath(working_dir)

    # ── 验证工作目录 ──
    if not os.path.isdir(working_dir):
        log("ERROR", f"工作目录不存在: {working_dir}")
        sys.exit(1)

    excel_path = os.path.join(working_dir, "分项报价表.xlsx")
    if not os.path.isfile(excel_path):
        log("ERROR", f"分项报价表.xlsx 不存在: {excel_path}")
        sys.exit(1)

    docs_dir = os.path.join(working_dir, "docs")
    if not os.path.isdir(docs_dir):
        log("ERROR", f"docs/ 目录不存在: {docs_dir}")
        sys.exit(1)

    log("INFO", f"工作目录: {working_dir}")
    log("INFO", f"Excel 文件: {excel_path}")

    # ── 读取分项名称 ──
    log("INFO", "正在读取分项报价表...")
    product_names = read_product_names(excel_path)
    if not product_names:
        log("ERROR", "分项报价表 B 列无数据")
        sys.exit(1)
    log("INFO", f"读取到 {len(product_names)} 个分项名称")

    # ── 准备输出目录 ──
    images_root = os.path.join(working_dir, "_images")

    # ── 遍历产品 ──
    manifest = {"products": [], "summary": {"total_products": 0, "matched": 0, "skipped": 0, "total_images": 0}}

    for name in product_names:
        product_dir = os.path.join(docs_dir, name)
        manifest["summary"]["total_products"] += 1

        if not os.path.isdir(product_dir):
            log("SKIP", f"产品目录不存在: docs/{name}/")
            manifest["summary"]["skipped"] += 1
            continue

        log("INFO", f"━━━ 处理产品: {name} ━━━")
        manifest["summary"]["matched"] += 1

        product_entry = {"name": name, "categories": {}}

        for category in CATEGORIES:
            cat_dir = os.path.join(product_dir, category)
            if not os.path.isdir(cat_dir):
                log("SKIP", f"  {category}/ 目录不存在")
                continue

            files = sorted([
                f for f in os.listdir(cat_dir)
                if not f.startswith(".") and not f.startswith("~")  # 跳过隐藏和临时文件
                and os.path.isfile(os.path.join(cat_dir, f))
            ])

            if not files:
                log("SKIP", f"  {category}/ 目录为空")
                continue

            log("INFO", f"  ── {category} ({len(files)} 个文件) ──")

            output_img_dir = os.path.join(images_root, name, category)
            os.makedirs(output_img_dir, exist_ok=True)

            all_images = []
            for filename in files:
                filepath = os.path.join(cat_dir, filename)
                images = process_file(filepath, output_img_dir)
                all_images.extend(images)

            if all_images:
                product_entry["categories"][category] = all_images
                manifest["summary"]["total_images"] += len(all_images)
                log("INFO", f"  {category}: 生成 {len(all_images)} 张图片")
            else:
                log("WARN", f"  {category}: 无有效图片")

        if product_entry["categories"]:
            manifest["products"].append(product_entry)
        else:
            log("WARN", f"产品 '{name}' 无任何有效图片，不会出现在文档中")

    # ── 输出清单文件 ──
    manifest_path = os.path.join(working_dir, "_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    log("OK", f"清单文件已生成: {manifest_path}")

    # ── 打印摘要 ──
    print()
    print("=" * 60)
    print("  处理摘要")
    print("=" * 60)
    print(f"  分项名称总数:   {manifest['summary']['total_products']}")
    print(f"  匹配到目录:     {manifest['summary']['matched']}")
    print(f"  跳过（无目录）: {manifest['summary']['skipped']}")
    print(f"  总图片数:       {manifest['summary']['total_images']}")
    print(f"  有图片的产品:   {len(manifest['products'])}")
    print()

    if manifest["products"]:
        for p in manifest["products"]:
            cats = ", ".join([f"{k}({len(v)})" for k, v in p["categories"].items()])
            print(f"    {p['name']}: {cats}")
    else:
        print("  ⚠ 没有任何产品生成了有效图片，不会创建 DOCX 文档")

    print()
    print("=" * 60)
    return manifest


# ═══════════════════════════════════════════════════════════════════
# CLI 入口
# ═══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python orchestrate.py <工作目录路径>")
        print("示例: python orchestrate.py D:\\投标项目\\项目A")
        sys.exit(1)

    working_dir = sys.argv[1]
    orchestrate(working_dir)
