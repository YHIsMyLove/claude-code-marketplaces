---
name: web-fetch
description: Extract and convert web page content to clean Markdown format with metadata. Uses Trafilatura library for intelligent content extraction and UV for fast dependency management.
license: Apache 2.0
allowed-tools: ["Read", "Write", "Bash"]
---

# Web Fetch - 网页内容提取器

## 概述

一个智能网页内容提取工具，能够：
- 从网页 URL 抓取主要内容
- 自动提取元数据（标题、作者、发布时间等）
- 输出格式化的 Markdown 文档
- 智能过滤广告、导航栏等无关内容

## 核心功能

### 1. 智能内容提取
- 使用 Trafilatura 库进行网页正文提取
- 支持多种网页格式（新闻文章、博客、技术文档等）
- 自动识别主内容区域，过滤导航和广告

### 2. 元数据提取
- **标题**: 页面或文章标题
- **作者**: 内容作者信息
- **发布时间**: 原创发布日期
- **网站名称**: 来源网站
- **URL**: 原始链接

### 3. 多格式输出
- Markdown（默认，推荐）
- JSON（包含完整元数据）
- HTML（保留结构）
- 纯文本（TXT）

## 使用场景

### 文章归档
```
用户: 帮我抓取 https://example.com/article 这篇文章
Claude: 正在提取网页内容...
✓ 标题: 深度学习入门指南
✓ 作者: 张三
✓ 发布时间: 2024-01-15
✓ 已保存到: article.md
```

### 批量内容收集
```
用户: 抓取这些网页的内容并保存
Claude: 开始批量提取网页内容...
处理 1/3: https://example.com/page1 → page1.md ✓
处理 2/3: https://example.com/page2 → page2.md ✓
处理 3/3: https://example.com/page3 → page3.md ✓
```

### 技术文档下载
```
用户: 下载这个技术文档并转为 Markdown
Claude: 正在提取技术文档...
✓ 保留代码格式
✓ 保留表格结构
✓ 已保存: technical-doc.md
```

## 技术要求

### 运行时要求
- **Python**: 3.8+ (推荐 3.10+)
- **包管理器**: UV（新一代 Python 包管理器）

### 依赖库
- **trafilatura**: 网页内容提取核心库
- **urllib3**: HTTP 请求处理

### 安装步骤

1. **安装 UV 包管理器**:
   ```bash
   # Windows (PowerShell)
   powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

   # macOS/Linux
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

2. **安装技能依赖**:
   ```bash
   cd skills/web-fetch
   uv sync
   ```

## 使用方法

### 基本用法
```bash
# 提取网页内容为 Markdown
uv run python scripts/cli.py https://example.com/article

# 指定输出文件
uv run python scripts/cli.py https://example.com/article --output my-article.md

# 提取并显示元数据
uv run python scripts/cli.py https://example.com/article --metadata
```

### 高级选项
```bash
# JSON 格式输出（包含完整元数据）
uv run python scripts/cli.py https://example.com/article --format json

# 包含评论和表格
uv run python scripts/cli.py https://example.com/article --include-comments --include-tables

# 保留图片链接
uv run python scripts/cli.py https://example.com/article --include-images

# 指定目标语言（过滤不匹配的内容）
uv run python scripts/cli.py https://example.com/article --language zh
```

## 输出格式

### Markdown 输出示例
```markdown
# 深度学习入门指南

**作者**: 张三
**发布时间**: 2024-01-15
**来源**: 示例网站
**原文链接**: https://example.com/article

---

## 正文内容

这里是提取的文章正文内容...
```

### JSON 输出示例
```json
{
  "title": "深度学习入门指南",
  "author": "张三",
  "date": "2024-01-15",
  "url": "https://example.com/article",
  "sitename": "示例网站",
  "text": "正文内容...",
  "comments": null
}
```

## 文件结构

```
web-fetch/
├── SKILL.md              # 技能定义文件
├── pyproject.toml        # UV 项目配置
├── README.md             # 使用说明
├── scripts/
│   ├── fetch.py          # 主提取逻辑
│   ├── cli.py            # 命令行接口
│   └── utils.py          # 工具函数
└── references/
    └── trafilatura-guide.md
```

## 错误处理

技能会处理以下错误情况：
- 无效的 URL 格式
- 网络连接失败
- 网页内容无法提取
- 元数据缺失警告
- 编码问题自动处理

## 特色优势

### 1. 极速依赖管理
使用 UV 包管理器，依赖安装速度提升 10-100 倍

### 2. 智能内容识别
基于机器学习的正文提取，准确率远超传统爬虫

### 3. 元数据丰富
自动提取标题、作者、时间等多种元数据

### 4. 跨平台兼容
支持 Windows、macOS、Linux，自动适配命令行工具

### 5. 多格式支持
Markdown、JSON、HTML、TXT 等多种输出格式

## 注意事项

- 某些网站可能有反爬虫机制，需遵守 robots.txt
- 建议性使用，避免对服务器造成过大压力
- 提取内容仅限个人学习使用，请尊重版权
- **仅支持静态 HTML 页面**：JavaScript 动态渲染的页面无法提取

## 参考资料

- [Trafilatura 官方文档](https://trafilatura.readthedocs.io/)
- [UV 包管理器文档](https://docs.astral.sh/uv/)
- [网页内容提取最佳实践](https://github.com/adbar/trafilatura)
