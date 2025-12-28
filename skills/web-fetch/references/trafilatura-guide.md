# Trafilatura 使用指南

## 核心概念

### extract() 函数

`extract()` 是 Trafilatura 的主函数，用于从 HTML 中提取主要内容。

```python
import trafilatura

# 下载网页
downloaded = trafilatura.fetch_url(url)

# 提取内容
result = trafilatura.extract(
    downloaded,
    url=url,
    output_format="markdown",
    with_metadata=True,
)
```

### bare_extraction() 函数

返回包含完整元数据的 Python 字典，适合需要单独处理各字段的场景。

```python
extracted = trafilatura.bare_extraction(
    downloaded,
    url=url,
    output_format="python",
    with_metadata=True,
)

# 访问元数据
title = extracted.title
author = extracted.author
date = extracted.date
```

## 常用参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `url` | str | 网页 URL，用于帮助元数据提取 |
| `output_format` | str | 输出格式：markdown, json, html, txt, xml |
| `with_metadata` | bool | 是否提取元数据 |
| `include_comments` | bool | 是否包含评论 |
| `include_tables` | bool | 是否包含表格 |
| `include_images` | bool | 是否包含图片 |
| `include_links` | bool | 是否保留链接 |
| `target_language` | str | 目标语言代码（如 zh, en） |

## 输出格式

### Markdown

保留格式和结构的文本，适合阅读和编辑。

```python
result = trafilatura.extract(html, output_format="markdown")
```

### JSON

包含完整元数据和内容的结构化数据。

```python
result = trafilatura.extract(html, output_format="json")
```

### XML/TEI

符合 Text Encoding Initiative 标准的学术格式。

```python
result = trafilatura.extract(html, output_format="xmltei")
```

## 元数据字段

使用 `bare_extraction()` 返回的对象包含以下字段：

| 字段 | 说明 |
|------|------|
| `title` | 页面或文章标题 |
| `author` | 内容作者 |
| `date` | 发布日期 |
| `url` | 原始 URL |
| `sitename` | 网站名称 |
| `categories` | 分类标签 |
| `tags` | 文章标签 |
| `text` | 正文内容 |

## 最佳实践

1. **优先使用 `extract()`**: 适合大多数场景
2. **指定 URL**: 帮助元数据提取更准确
3. **设置目标语言**: 过滤不匹配的内容
4. **处理编码问题**: Trafilatura 会自动检测编码

## 示例代码

### 基本提取

```python
import trafilatura

url = "https://example.com/article"
downloaded = trafilatura.fetch_url(url)
result = trafilatura.extract(downloaded, url=url)
print(result)
```

### 获取元数据

```python
downloaded = trafilatura.fetch_url(url)
extracted = trafilatura.bare_extraction(
    downloaded,
    url=url,
    with_metadata=True,
)

print(f"标题: {extracted.title}")
print(f"作者: {extracted.author}")
print(f"日期: {extracted.date}")
```

### Markdown 输出

```python
downloaded = trafilatura.fetch_url(url)
result = trafilatura.extract(
    downloaded,
    url=url,
    output_format="markdown",
    with_metadata=True,
    include_tables=True,
)
```

## 参考资料

- [官方文档](https://trafilatura.readthedocs.io/)
- [GitHub 仓库](https://github.com/adbar/trafilatura)
- [核心函数 API](https://trafilatura.readthedocs.io/en/latest/corefunctions.html)
