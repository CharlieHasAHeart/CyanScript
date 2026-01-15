# CyanScript

CyanScript 是一个基于 Markdown + Word 模板的软著说明书生成工具。它通过 `docxtpl` 渲染模板，占位符替换，同时将 Markdown 内容解析为 Word 段落、图片、表格、代码块与行内代码。

## 功能特性

- 交互式 CLI 输入：软件名称、版本号、工作目录、MD 文件
- 模板渲染（通过 `.env` 中的 `CYANSCRIPT_TEMPLATE` 指定）
- Markdown 内容转换：
  - 标题映射（标题 1~4）
  - 正文、列表、图片、引用块
  - 表格（含表注样式与表头/正文样式）
  - 代码块与行内代码
  - 自动更新目录（打开 Word 时提示更新）

## 安装

使用可编辑安装（推荐）：

```bash
pip install -e .
```

安装完成后可用命令：

```bash
cyanscript
```

## 使用

运行后按提示输入：

- 工作目录（可选）
- 软件名称
- 版本号
- MD 文件

输出文件：

```
./软件名_版本号_软件说明书.docx
```

### 模板路径查找规则

仅使用环境变量 `CYANSCRIPT_TEMPLATE` 指定模板路径。脚本会优先读取“脚本目录下的 .env”，再读取“当前工作目录的 .env”，首次生效为准。

## 示例

使用 `test/docs_sample.md` 进行测试，包含图片、表格、列表、引用块与代码块示例。

## 模板样式清单

已整理模板样式与示例展示：`assets/template_styles.md`。

## 工具脚本

- 模板自检：`scripts/template_selfcheck.py`
- 修复页眉/页脚占位符断裂：`scripts/fix_header_placeholders.py`
- 修复 `main_content` 占位符：`scripts/fix_main_content_placeholder.py`
- 修复封面 `{{software_name}}` 占位符断裂：`scripts/fix_cover_title_placeholder.py`

## 许可证

本项目使用 Apache-2.0 许可证。

## 版权

Copyright 2026 米凌轩
