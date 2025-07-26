## Language Selection

- [简体中文](https://github.com/Kanyooooo/Yuque2Hexo-tools/blob/main/README.md) | English

---

# Yuque2HEXO-tools

Yuque2HEXO-tools is a set of tools designed to convert Yuque documents into Hexo blog posts. Currently, the tool provides the following features:

- Extract images from .docx files and save them to both local and Hexo image directories.
- Download and process external image links in Markdown files.
- Rewrite Yuque links and local file links to a format compatible with Hexo.
- Automatically generate and insert Hexo blog post Front-Matter (including tags, categories, and date).
- Support batch processing of multiple .docx files in the current directory.

## Configuration Overview

- `OUTPUT_PY_IMG_DIR`: Directory to store images locally.
- `BASE_URL`: Prefix for Hexo image URLs.
- `OUTPUT_HEXO_MD_DIR`: Hexo post directory.
- `OUTPUT_HEXO_IMG_DIR`: Hexo image directory.

## Features

- **Image Extraction**: Supports extracting all images from .docx files and saving them to specified directories, while uploading them to the Hexo theme image directory.
- **Link Rewriting**: Rewrites Yuque links, external image links, and local file links into Hexo-compatible formats.
- **Markdown Processing**: Automatically processes images and links in Markdown files to ensure the correct Hexo post format.
- **Batch Processing**: Allows processing multiple .docx files at once, saving time on manual conversions.

## Usage

1. Configure Parameters: Set the directory paths according to your Hexo project.
2. Run the Tool: Place the .docx files and Markdown files in the same directory, and run the tool. It will automatically process and generate Hexo-compatible posts.

## Future Updates

This tool will be further enhanced in the future with additional features, so stay tuned!
