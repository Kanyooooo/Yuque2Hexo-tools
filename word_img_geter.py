import os
import re
import shutil
import urllib.parse
from datetime import datetime
import requests
import xml.etree.ElementTree as ET
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
from PIL import Image
from docx.oxml.xmlchemy import OxmlElement

OUTPUT_PY_IMG_DIR = "img/filesimg"
BASE_URL = "img/filesimg"
OUTPUT_HEXO_MD_DIR = r"D:\hexo\source\_posts"
OUTPUT_HEXO_IMG_DIR = r"D:\hexo\themes\butterfly\source\img\filesimg"

TAGS = ["计算机原理"]
CATEGORIES = ["CSAPP - 深入了解计算机系统"]

FORMULA_MARKER = "⚡FORMULA⚡"
FORMULA_END_MARKER = "⚡FORMULA_END⚡"


def rewrite_links(content: str, folder_name: str) -> str:
    def rewrite_yuque_links(md_text: str, folder_name: str) -> str:
        pattern = re.compile(r"\[([^\]]+)\]\((https?://www\.yuque\.com/[^\)]+)\)")

        def repl(m: re.Match) -> str:
            text = m.group(1)
            encoded_folder = urllib.parse.quote(folder_name, safe='')
            return f"[{text}](/docx/{encoded_folder}/)"

        return pattern.sub(repl, md_text)

    def rewrite_local_links(md_text: str) -> str:
        pattern = re.compile(r"\[([^\]]+)\]\((?!https?://)(?!.*?\.(png|jpg|jpeg|gif|bmp))[^)]*\)")

        def repl(m: re.Match) -> str:
            text = m.group(1)
            encoded = urllib.parse.quote(text, safe='')
            return f"[{text}](/docx/{encoded}/)"

        return pattern.sub(repl, md_text)

    content = rewrite_yuque_links(content, folder_name)
    content = rewrite_local_links(content)
    return content


def extract_images_from_word(docx_path: str, folder_name: str) -> list:
    local_img_dir = os.path.join(OUTPUT_PY_IMG_DIR, folder_name)
    os.makedirs(local_img_dir, exist_ok=True)
    hexo_img_dir = os.path.join(OUTPUT_HEXO_IMG_DIR, folder_name)
    os.makedirs(hexo_img_dir, exist_ok=True)

    doc = Document(docx_path)
    image_info = []
    img_counter = 0

    # 新的图片提取方法，使用文档关系查找图片
    for rel_id, part in doc.part.related_parts.items():
        if hasattr(part, 'content_type') and part.content_type.startswith('image/'):
            image_data = part.blob

            try:
                with BytesIO(image_data) as img_stream:
                    img = Image.open(img_stream)
                    width, height = img.size
                    # 公式通常长宽比异常
                    is_formula = (width / height > 5) or (height / width > 5)
            except Exception as e:
                print(f"  解析图片出错: {str(e)}")
                is_formula = False

            prefix = "formula" if is_formula else "image"
            image_name = f"{prefix}_{folder_name}_{img_counter}.png"
            img_counter += 1

            # 保存到本地目录
            local_path = os.path.join(local_img_dir, image_name)
            with open(local_path, 'wb') as f:
                f.write(image_data)

            # 保存到Hexo目录
            hexo_path = os.path.join(hexo_img_dir, image_name)
            shutil.copy(local_path, hexo_path)

            image_info.append((image_name, is_formula))

    return image_info


def download_external_images(content: str, folder_name: str) -> tuple:
    img_pattern = re.compile(r'!\[(.*?)\]\((https?://[^\)]+)\)')
    downloaded_images = []
    img_counter = 0

    local_img_dir = os.path.join(OUTPUT_PY_IMG_DIR, folder_name)
    os.makedirs(local_img_dir, exist_ok=True)
    hexo_img_dir = os.path.join(OUTPUT_HEXO_IMG_DIR, folder_name)
    os.makedirs(hexo_img_dir, exist_ok=True)

    def replace_external_image(match):
        nonlocal img_counter
        alt_text = match.group(1)
        img_url = match.group(2)

        try:
            response = requests.get(img_url, stream=True, timeout=10)
            response.raise_for_status()

            ext = os.path.splitext(img_url)[1].lower()
            if not ext or ext not in ['.png', '.jpg', '.jpeg', '.gif']:
                ext = '.png'

            image_name = f"{folder_name}_external_{img_counter}{ext}"
            img_counter += 1

            local_path = os.path.join(local_img_dir, image_name)
            with open(local_path, 'wb') as f:
                for chunk in response.iter_content(8192):
                    f.write(chunk)

            hexo_path = os.path.join(hexo_img_dir, image_name)
            shutil.copy(local_path, hexo_path)

            downloaded_images.append((image_name, False))

            img_url = f"{BASE_URL}/{folder_name}/{image_name}"
            encoded_url = img_url.replace(' ', '%20')
            return f"![{alt_text}]({encoded_url})"

        except Exception as e:
            print(f"  警告: 无法下载图片 {img_url} ({str(e)})")
            return match.group(0)

    processed_content = img_pattern.sub(replace_external_image, content)
    return processed_content, downloaded_images


def mark_formulas(content: str) -> str:
    # 块级公式：$$...$$
    content = re.sub(r'(\$\$(.*?)\$\$)',
                     lambda m: f"{FORMULA_MARKER}{m.group(0)}{FORMULA_END_MARKER}",
                     content, flags=re.DOTALL)

    # 行内公式：$...$
    content = re.sub(r'(\$(.*?)\$)',
                     lambda m: f"{FORMULA_MARKER}{m.group(0)}{FORMULA_END_MARKER}",
                     content)

    # LaTeX环境公式
    latex_environments = ['equation', 'align', 'gather']
    for env in latex_environments:
        pattern = rf'(\s*\\begin{{{env}}}.*?\\end{{{env}}}\s*)'
        content = re.sub(pattern,
                         lambda m: f"{FORMULA_MARKER}{m.group(0)}{FORMULA_END_MARKER}",
                         content, flags=re.DOTALL)

    return content


def replace_formula_markers(content: str) -> str:
    content = content.replace(FORMULA_MARKER, "")
    return content.replace(FORMULA_END_MARKER, "")


def process_markdown_file(md_path: str, folder_name: str, image_info: list):
    print(f"  开始处理Markdown文件: {md_path}")

    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    print(f"  已读取Markdown内容，长度: {len(content)} 字符")

    # 移除现有的Front-Matter
    fm_pattern = re.compile(r'^---\n(.*?\n)---\n', re.DOTALL)
    match = fm_pattern.search(content)
    if match:
        content = content.replace(match.group(0), '', 1)
        print(f"  已移除现有的Front-Matter")

    # 标记公式
    print("  开始标记公式")
    content = mark_formulas(content)
    print("  完成标记公式")

    # 下载并替换外部图片
    print("  开始下载并替换外部图片")
    content, external_images = download_external_images(content, folder_name)
    print(f"  完成下载并替换外部图片，共下载 {len(external_images)} 张外部图片")

    # 合并图片信息
    all_images = image_info + external_images
    formula_images = [img for img, is_formula in all_images if is_formula]
    non_formula_images = [img for img, is_formula in all_images if not is_formula]
    print(f"  总图片数量: {len(all_images)} (公式图片: {len(formula_images)}, 普通图片: {len(non_formula_images)})")

    # 查找公式范围
    formula_ranges = []
    start_idx = 0
    while start_idx < len(content):
        start_marker = content.find(FORMULA_MARKER, start_idx)
        if start_marker == -1:
            break

        end_marker = content.find(FORMULA_END_MARKER, start_marker)
        if end_marker == -1:
            break

        end_pos = end_marker + len(FORMULA_END_MARKER)
        formula_ranges.append((start_marker, end_pos))
        start_idx = end_pos

    # 替换图片链接
    img_pattern = re.compile(r'!\[(.*?)\]\(([^)]+)\)')
    img_count = 0
    skipped_formulas = 0

    def replace_image(match):
        nonlocal img_count, skipped_formulas
        match_start = match.start()

        # 检查是否在公式范围内
        for start, end in formula_ranges:
            if start <= match_start < end:
                skipped_formulas += 1
                return match.group(0)

        # 替换普通图片
        if img_count < len(non_formula_images):
            alt_text = match.group(1)
            img_name = non_formula_images[img_count]
            img_count += 1
            img_url = f"{BASE_URL}/{folder_name}/{img_name}"
            encoded_url = img_url.replace(' ', '%20')  # 替换空格
            return f"%7Bencoded_url%7D"

        return match.group(0)

    print("  开始替换图片链接")
    content = img_pattern.sub(replace_image, content)
    print(f"  完成替换内嵌图片链接: 处理了 {img_count} 张图片, 跳过了 {skipped_formulas} 个公式位置")

    # 重写链接
    print("  开始重写链接")
    content = rewrite_links(content, folder_name)
    print("  完成重写链接")

    # 清理公式标记
    print("  开始清理公式标记占位符")
    content = replace_formula_markers(content)
    print("  完成清理公式标记占位符")

    # 创建Front-Matter
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    front_matter = [
        "---",
        f'title: "{folder_name}"',
        "tags:"
    ]
    for tag in TAGS:
        front_matter.append(f'    - "{tag}"')
    front_matter.append("categories:")
    for category in CATEGORIES:
        front_matter.append(f'    - "{category}"')
    front_matter.append(f'date: "{now}"')
    front_matter.append("---\n")
    fm = "\n".join(front_matter)
    final_content = fm + content

    # 保存到原始位置
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(final_content)
    print(f"  已保存到原始位置: {md_path}")

    # 保存到Hexo目录
    hexo_md_path = os.path.join(OUTPUT_HEXO_MD_DIR, os.path.basename(md_path))
    os.makedirs(os.path.dirname(hexo_md_path), exist_ok=True)
    with open(hexo_md_path, 'w', encoding='utf-8') as f:
        f.write(final_content)
    print(f"  已保存到Hexo目录: {hexo_md_path}")

    return img_count, len(external_images), skipped_formulas


def batch_process():
    cwd = os.getcwd()
    processed_count = 0

    for file_name in os.listdir(cwd):
        if not file_name.lower().endswith('.docx'):
            continue

        base_name = os.path.splitext(file_name)[0]
        md_file = f"{base_name}.md"
        md_path = os.path.join(cwd, md_file)

        if not os.path.exists(md_path):
            print(f"跳过 {file_name}，未找到对应的Markdown文件")
            continue

        try:
            print(f"处理: {base_name}")
            # 提取图片
            image_info = extract_images_from_word(file_name, base_name)
            print(f"  找到图片: {len(image_info)}")

            # 处理Markdown文件
            img_count, external_count, skipped_formulas = process_markdown_file(
                md_path, base_name, image_info)

            print(
                f"  成功处理: 替换了 {img_count} 张图片, 下载了 {external_count} 张外部图片, "
                f"跳过了 {skipped_formulas} 个公式位置"
            )
            processed_count += 1
        except Exception as e:
            import traceback
            print(f"  处理 {base_name} 时出错: {str(e)}")
            traceback.print_exc()

    print(f"\n处理完成! 共处理 {processed_count} 个文档")


if __name__ == '__main__':
    os.makedirs(OUTPUT_PY_IMG_DIR, exist_ok=True)
    os.makedirs(OUTPUT_HEXO_IMG_DIR, exist_ok=True)
    os.makedirs(OUTPUT_HEXO_MD_DIR, exist_ok=True)
    batch_process()