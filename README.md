## 语言选择

- 简体中文 | [English](https://github.com/Kanyooooo/Yuque2Hexo-tools/blob/main/README_en.md)

---

# Yuque2HEXO-tools

Yuque2HEXO-tools 是一个用于将语雀文档（Yuque）转换为 Hexo 博客文章的工具集。该工具主要实现了以下功能：

- 提取 .docx 文件中的图片并自动保存到本地和 Hexo 图片目录。
- 下载并处理 Markdown 文件中的外部图片链接。
- 替换文档中的语雀链接和本地文件链接，重写为 Hexo 兼容的链接格式。
- 自动生成并插入 Hexo 博客文章的 Front-Matter 配置（包括标签、分类和发布日期等）。
- 支持批量处理当前目录下的多个 .docx 文件。

## 配置说明

- `OUTPUT_PY_IMG_DIR`: 本地存储图片的目录。
- `BASE_URL`: Hexo 图片 URL 前缀。
- `OUTPUT_HEXO_MD_DIR`: Hexo 文章目录。
- `OUTPUT_HEXO_IMG_DIR`: Hexo 图片目录。

## 功能特点

- **提取图片**: 支持提取 .docx 文件中的所有图片并将其保存在指定目录，同时上传到 Hexo 主题的图片目录。
- **重写链接**: 将语雀链接、外部图片链接和本地文件链接重写为 Hexo 支持的格式。
- **Markdown 处理**: 自动处理 Markdown 文件中的图片和链接，确保 Hexo 文章格式正确。
- **批量处理**: 可一次性处理多个 .docx 文件，节省手动转换的时间。

## 使用方法

1. 配置参数：根据自己的 Hexo 项目配置相关目录路径。
2. 运行工具：将 .docx 文件与 Markdown 文件放置在同一目录下，运行该工具，它会自动处理并生成 Hexo 格式的文章。

## 后续更新

未来该工具将进一步扩展，增加更多实用的功能，敬请期待！
