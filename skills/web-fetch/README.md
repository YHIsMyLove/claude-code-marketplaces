# Web Fetch

智能网页内容提取工具，使用 Trafilatura 库将网页内容转换为干净的 Markdown 格式。

## 快速开始

### 1. 安装 UV 包管理器

```bash
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 2. 安装依赖

```bash
cd skills/web-fetch
uv sync
```

### 3. 使用示例

```bash
# 基本用法 - 输出到终端
uv run python scripts/cli.py https://example.com/article

# 保存到文件
uv run python scripts/cli.py https://example.com/article -o article.md

# JSON 格式输出
uv run python scripts/cli.py https://example.com/article --format json -o article.json

# 显示详细信息
uv run python scripts/cli.py https://example.com/article -v
```

## 功能特性

- 智能正文提取，自动过滤广告和导航
- 提取元数据（标题、作者、发布时间等）
- 多种输出格式（Markdown、JSON、HTML、TXT）
- 跨平台支持（Windows、macOS、Linux）

## 命令行参数

| 参数 | 说明 |
|------|------|
| `url` | 要提取的网页 URL（必需） |
| `-o, --output` | 输出文件路径 |
| `-f, --format` | 输出格式：markdown, json, html, txt |
| `--no-metadata` | 不包含元数据 |
| `--include-comments` | 包含文章评论 |
| `--no-tables` | 不包含表格 |
| `--include-images` | 包含图片链接 |
| `--include-links` | 包含链接 |
| `-l, --language` | 目标语言代码（如 zh, en） |
| `-v, --verbose` | 显示详细输出 |

## 输出格式示例

### Markdown 输出

```markdown
# 文章标题

**作者**: 张三
**发布时间**: 2024-01-15
**来源**: 示例网站
**原文链接**: https://example.com/article

---

正文内容...
```

### JSON 输出

```json
{
  "content": "正文内容...",
  "metadata": {
    "title": "文章标题",
    "author": "张三",
    "date": "2024-01-15",
    "url": "https://example.com/article",
    "sitename": "示例网站"
  },
  "url": "https://example.com/article"
}
```

## 常见问题

### Q: 提示找不到 uv 命令？

A: 请确保 UV 已正确安装并添加到系统 PATH。重新打开终端后再试。

### Q: 某些网页无法提取？

A: 本工具仅支持静态 HTML 页面。JavaScript 动态渲染的页面（如 Vue/React 应用）无法提取。

### Q: 如何批量提取多个网页？

A: 可以编写 shell 脚本循环调用：

```bash
while read -r url; do
    uv run python scripts/cli.py "$url" -o "$(date +%s).md"
done < urls.txt
```

## 参考资料

- [Trafilatura 官方文档](https://trafilatura.readthedocs.io/)
- [UV 包管理器文档](https://docs.astral.sh/uv/)

## 许可证

Apache 2.0
