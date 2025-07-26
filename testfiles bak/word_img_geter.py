import os
import re
import shutil
import urllib.parse
from datetime import datetime
from io import BytesIO
from PIL import Image
from docx import Document
import requests
import xml.etree.ElementTree as ET
from docx.oxml import parse_xml

# ===== 配置参数 =====
OUTPUT_PY_IMG_DIR = "imgshit/img/filesimg"  # 本地图片输出目录
BASE_URL = "/img/filesimg"  # Hexo图片URL前缀
OUTPUT_HEXO_MD_DIR = r"D:\hexo\source\_posts"  # Hexo文章目录
OUTPUT_HEXO_IMG_DIR = r"D:\hexo\themes\butterfly\source\img\filesimg"  # Hexo图片目录

# 全局Front-Matter配置
TAGS = ["计算机原理"]
CATEGORIES = ["CSAPP - 深入了解计算机系统"]


OUTPUT_HEXO_IMG_DIR = OUTPUT_HEXO_MD_DIR = r"D:\hexotest"


# ===== 核心函数 =====
def rewrite_links(content: str, folder_name: str) -> str:
    """重写所有非HTTP链接和特定域名的链接为/docx/编码格式，但排除图片链接"""

    # 重写语雀链接为 /docx/文件夹名/
    def rewrite_yuque_links(md_text: str, folder_name: str) -> str:
        """重写语雀链接为 /docx/文件夹名/ 格式"""
        pattern = re.compile(r"\[([^\]]+)\]\((https?://www\.yuque\.com/[^\)]+)\)")

        def repl(m: re.Match) -> str:
            text = m.group(1)
            # 使用当前文件夹名作为链接路径
            encoded_folder = urllib.parse.quote(folder_name, safe='')
            return f"[{text}](/docx/{encoded_folder}/)"

        return pattern.sub(repl, md_text)

    # 重写普通非HTTP链接
    def rewrite_local_links(md_text: str) -> str:
        """重写所有非HTTP链接为/docx/编码格式"""
        pattern = re.compile(r"\[([^\]]+)\]\((?!https?://)(?!.*?\.(png|jpg|jpeg|gif|bmp))[^)]*\)")

        def repl(m: re.Match) -> str:
            text = m.group(1)
            encoded = urllib.parse.quote(text, safe='')
            return f"[{text}](/docx/{encoded}/)"

        return pattern.sub(repl, md_text)

    # 先处理语雀链接
    content = rewrite_yuque_links(content, folder_name)
    # 再处理普通本地链接
    content = rewrite_local_links(content)
    return content


def extract_images_from_word(docx_path: str, folder_name: str) -> list:
    """从.docx提取图片，确保按照文档中的顺序"""
    # 创建本地图片目录
    local_img_dir = os.path.join(OUTPUT_PY_IMG_DIR, folder_name)
    os.makedirs(local_img_dir, exist_ok=True)

    # 创建Hexo图片目录
    hexo_img_dir = os.path.join(OUTPUT_HEXO_IMG_DIR, folder_name)
    os.makedirs(hexo_img_dir, exist_ok=True)

    doc = Document(docx_path)
    image_filenames = []
    img_counter = 0  # 图片计数器

    # 方法1：使用更全面的命名空间查询
    namespaces = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
    }

    # 获取文档XML结构
    xml_str = doc.part._element.xml
    root = ET.fromstring(xml_str)

    # 方法2：检查文档中所有的blip元素
    blip_elements = []

    # 查找所有可能的blip元素位置
    for element_name in [
        './/wp:blipFill/a:blip',
        './/a:blip',
        './/wps:txbx/w:r/w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip',
        './/pic:blipFill/a:blip'
    ]:
        try:
            elements = root.findall(element_name, namespaces)
            for elem in elements:
                blip_attr = elem.attrib.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if blip_attr:
                    blip_elements.append(blip_attr)
        except SyntaxError:
            # 忽略语法错误，尝试下一种方法
            pass

    # 去重但保持顺序
    pic_refs = []
    seen = set()
    for rId in blip_elements:
        if rId not in seen:
            seen.add(rId)
            pic_refs.append(rId)

    # 按文档中出现的顺序提取图片
    for rId in pic_refs:
        if rId in doc.part.related_parts:
            image_part = doc.part.related_parts[rId]
            if hasattr(image_part, 'blob'):
                image_data = image_part.blob

                # 生成图片文件名（格式：文件夹名_连续序号.png）
                image_name = f"{folder_name}_{img_counter}.png"
                img_counter += 1

                try:
                    # 保存到本地目录
                    local_path = os.path.join(local_img_dir, image_name)
                    with open(local_path, 'wb') as f:
                        f.write(image_data)

                    # 保存到Hexo目录
                    hexo_path = os.path.join(hexo_img_dir, image_name)
                    shutil.copy(local_path, hexo_path)

                    image_filenames.append(image_name)
                    print(f"  已提取图片: {image_name}")
                except Exception as e:
                    print(f"  警告: 保存图片时出错 ({str(e)})")

    # 方法3：如果以上方法都没找到，尝试直接遍历所有关系
    if not image_filenames:
        print("  尝试方法3：直接遍历文档关系")
        for rel_id, part in doc.part.related_parts.items():
            if part.content_type.startswith('image/'):
                image_data = part.blob
                image_name = f"{folder_name}_rel_{rel_id}.png"
                try:
                    # 保存到本地目录
                    local_path = os.path.join(local_img_dir, image_name)
                    with open(local_path, 'wb') as f:
                        f.write(image_data)

                    # 保存到Hexo目录
                    hexo_path = os.path.join(hexo_img_dir, image_name)
                    shutil.copy(local_path, hexo_path)

                    image_filenames.append(image_name)
                    print(f"  已提取图片: {image_name}")
                except Exception as e:
                    print(f"  警告: 保存图片时出错 ({str(e)})")

    if not image_filenames:
        print("  未找到图片内容")

    return image_filenames


def download_external_images(content: str, folder_name: str) -> tuple:
    """
    下载Markdown中的所有外部图片并替换为本地链接
    返回：(处理后的内容, 下载的图片文件名列表)
    """
    # 匹配所有图片链接
    img_pattern = re.compile(r'!\[(.*?)\]\((https?://[^\)]+)\)')
    downloaded_images = []
    img_counter = 0

    # 创建本地图片目录
    local_img_dir = os.path.join(OUTPUT_PY_IMG_DIR, folder_name)
    os.makedirs(local_img_dir, exist_ok=True)

    # 创建Hexo图片目录
    hexo_img_dir = os.path.join(OUTPUT_HEXO_IMG_DIR, folder_name)
    os.makedirs(hexo_img_dir, exist_ok=True)

    def replace_external_image(match):
        nonlocal img_counter
        alt_text = match.group(1)
        img_url = match.group(2)

        try:
            # 下载图片
            response = requests.get(img_url, stream=True)
            response.raise_for_status()

            # 生成唯一的图片文件名
            ext = os.path.splitext(img_url)[1].lower()
            if not ext or ext not in ['.png', '.jpg', '.jpeg', '.gif']:
                ext = '.png'  # 默认使用PNG格式

            image_name = f"{folder_name}_external_{img_counter}{ext}"
            img_counter += 1

            # 保存到本地目录
            local_path = os.path.join(local_img_dir, image_name)
            with open(local_path, 'wb') as f:
                for chunk in response.iter_content(8192):
                    f.write(chunk)

            # 保存到Hexo目录
            hexo_path = os.path.join(hexo_img_dir, image_name)
            shutil.copy(local_path, hexo_path)

            # 添加到下载列表
            downloaded_images.append(image_name)

            # 返回新的Markdown图片链接
            img_url = f"{BASE_URL}/{folder_name}/{image_name}"
            encoded_url = img_url.replace(' ', '%20')
            return f"![{alt_text}]({encoded_url})"

        except Exception as e:
            print(f"  警告: 无法下载图片 {img_url} ({str(e)})")
            return match.group(0)  # 返回原始链接

    # 替换所有外部图片链接
    processed_content = img_pattern.sub(replace_external_image, content)

    return processed_content, downloaded_images


def process_markdown_file(md_path: str, folder_name: str, image_filenames: list):
    """处理Markdown文件并保存到两个位置"""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 移除已存在的Front-Matter（如果存在）
    fm_pattern = re.compile(r'^---\n(.*?\n)---\n', re.DOTALL)
    match = fm_pattern.search(content)
    if match:
        # 移除匹配到的Front-Matter部分
        content = content.replace(match.group(0), '', 1)
        print(f"  已移除现有的Front-Matter")

    # 下载并替换所有外部图片
    content, external_images = download_external_images(content, folder_name)
    all_images = image_filenames + external_images

    # 添加Hexo Front-Matter（所有值添加双引号）
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 使用双引号包装所有值
    front_matter = [
        "---",
        f'title: "{folder_name}"',  # 添加双引号
        "tags:"
    ]

    for tag in TAGS:
        front_matter.append(f'    - "{tag}"')  # 添加双引号

    front_matter.append("categories:")
    for category in CATEGORIES:
        front_matter.append(f'    - "{category}"')  # 添加双引号

    front_matter.append(f'date: "{now}"')  # 添加双引号
    front_matter.append("---\n")
    fm = "\n".join(front_matter)

    # 替换Word内嵌图片链接
    img_pattern = re.compile(r'!\[(.*?)\]\([^)]*\)')
    img_count = 0

    def replace_image_link(match):
        nonlocal img_count
        if img_count < len(all_images):
            alt_text = match.group(1)
            # 构建正确的图片URL格式: /img/filesimg/[filename]/[filename]_[编号].png
            img_url = f"{BASE_URL}/{folder_name}/{all_images[img_count]}"
            # 对URL进行编码（替换空格为%20）
            encoded_url = img_url.replace(' ', '%20')
            new_link = f"![{alt_text}]({encoded_url})"
            img_count += 1
            return new_link
        return match.group(0)

    content = img_pattern.sub(replace_image_link, content)

    # 重写链接（包括语雀链接）
    content = rewrite_links(content, folder_name)

    # 添加Front-Matter
    final_content = fm + content

    # 保存到原始位置（覆盖）
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(final_content)

    # 保存到Hexo目录
    hexo_md_path = os.path.join(OUTPUT_HEXO_MD_DIR, os.path.basename(md_path))
    os.makedirs(os.path.dirname(hexo_md_path), exist_ok=True)
    with open(hexo_md_path, 'w', encoding='utf-8') as f:
        f.write(final_content)

    return img_count, len(external_images)

def batch_process():
    """批量处理当前目录下的所有.docx文件"""
    cwd = os.getcwd()
    processed_count = 0

    for file_name in os.listdir(cwd):
        if not file_name.lower().endswith('.docx'):
            continue

        # 获取基本文件名（不含扩展名）
        base_name = os.path.splitext(file_name)[0]
        md_file = f"{base_name}.md"
        md_path = os.path.join(cwd, md_file)

        if not os.path.exists(md_path):
            print(f"跳过 {file_name}，未找到对应的Markdown文件")
            continue

        try:
            print(f"处理: {base_name}")
            # 提取Word内嵌图片（保持顺序）
            image_filenames = extract_images_from_word(file_name, base_name)

            if not image_filenames:
                print(f"  警告: 未在Word文档中找到图片")

            # 处理Markdown
            image_count, external_count = process_markdown_file(md_path, base_name, image_filenames)

            print(f"  成功处理: 替换了 {image_count} 张图片 (Word内嵌: {len(image_filenames)}, 外部: {external_count})")
            processed_count += 1
        except Exception as e:
            import traceback
            print(f"  处理 {base_name} 时出错: {str(e)}")
            traceback.print_exc()

    print(f"\n处理完成! 共处理 {processed_count} 个文档")


if __name__ == '__main__':
    # 创建必要的目录
    os.makedirs(OUTPUT_PY_IMG_DIR, exist_ok=True)
    os.makedirs(OUTPUT_HEXO_IMG_DIR, exist_ok=True)
    os.makedirs(OUTPUT_HEXO_MD_DIR, exist_ok=True)

    batch_process()