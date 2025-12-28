#!/usr/bin/env python3
"""
Web Fetch - 核心内容提取模块
使用 Trafilatura 库进行智能网页内容提取
"""

import json
from typing import Optional, Dict, Any
from urllib.parse import urlparse
import trafilatura


class WebFetcher:
    """网页内容提取器"""

    def __init__(
        self,
        output_format: str = "markdown",
        include_metadata: bool = True,
        include_comments: bool = False,
        include_tables: bool = True,
        include_images: bool = False,
        include_links: bool = False,
        target_language: Optional[str] = None,
    ):
        """
        初始化提取器

        Args:
            output_format: 输出格式 (markdown, json, html, txt)
            include_metadata: 是否包含元数据
            include_comments: 是否包含评论
            include_tables: 是否包含表格
            include_images: 是否包含图片
            include_links: 是否包含链接
            target_language: 目标语言代码（如 'zh', 'en'）
        """
        self.output_format = output_format
        self.include_metadata = include_metadata
        self.include_comments = include_comments
        self.include_tables = include_tables
        self.include_images = include_images
        self.include_links = include_links
        self.target_language = target_language

    def validate_url(self, url: str) -> bool:
        """
        验证 URL 格式

        Args:
            url: 待验证的 URL

        Returns:
            是否为有效的 URL
        """
        try:
            result = urlparse(url)
            return all([result.scheme in ["http", "https"], result.netloc])
        except Exception:
            return False

    def fetch(self, url: str) -> Optional[str]:
        """
        下载网页内容

        Args:
            url: 网页 URL

        Returns:
            下载的 HTML 内容，失败返回 None
        """
        try:
            downloaded = trafilatura.fetch_url(url)
            if downloaded:
                return downloaded
            return None
        except Exception as e:
            raise RuntimeError(f"下载网页失败: {e}")

    def extract(
        self,
        url: str,
        html_content: Optional[str] = None
    ) -> Optional[Dict[str, Any]]:
        """
        提取网页内容和元数据

        Args:
            url: 网页 URL
            html_content: 可选的 HTML 内容（如果已下载）

        Returns:
            包含内容和元数据的字典
        """
        # 验证 URL
        if not self.validate_url(url):
            raise ValueError(f"无效的 URL: {url}")

        # 下载网页（如果未提供 HTML）
        if html_content is None:
            html_content = self.fetch(url)
            if html_content is None:
                raise RuntimeError(f"无法下载网页: {url}")

        # 执行提取
        result = trafilatura.extract(
            html_content,
            url=url,
            output_format=self.output_format,
            with_metadata=self.include_metadata,
            include_comments=self.include_comments,
            include_tables=self.include_tables,
            include_images=self.include_images,
            include_links=self.include_links,
            target_language=self.target_language,
        )

        if result is None:
            return None

        # 获取元数据
        metadata = {}
        if self.include_metadata:
            extracted = trafilatura.bare_extraction(
                html_content,
                url=url,
                output_format="python",
                with_metadata=True,
                include_comments=self.include_comments,
                include_tables=self.include_tables,
            )

            if extracted:
                metadata = {
                    "title": extracted.title or "未知标题",
                    "author": extracted.author or "未知作者",
                    "date": extracted.date or "未知日期",
                    "url": extracted.url or url,
                    "sitename": extracted.sitename or "未知网站",
                    "categories": extracted.categories or [],
                    "tags": extracted.tags or [],
                }

        return {
            "content": result,
            "metadata": metadata,
            "url": url,
        }

    def format_markdown(self, data: Dict[str, Any]) -> str:
        """
        将提取结果格式化为 Markdown

        Args:
            data: 提取的数据字典

        Returns:
            Markdown 格式的字符串
        """
        metadata = data.get("metadata", {})
        content = data.get("content", "")

        # 构建 Markdown 头部
        md_parts = []

        # 标题
        title = metadata.get("title", "未知标题")
        md_parts.append(f"# {title}\n")

        # 元数据块
        meta_info = []
        if metadata.get("author"):
            meta_info.append(f"**作者**: {metadata['author']}")
        if metadata.get("date"):
            meta_info.append(f"**发布时间**: {metadata['date']}")
        if metadata.get("sitename"):
            meta_info.append(f"**来源**: {metadata['sitename']}")
        if metadata.get("url"):
            meta_info.append(f"**原文链接**: {metadata['url']}")

        if meta_info:
            md_parts.append("\n".join(meta_info))
            md_parts.append("\n")
            md_parts.append("---\n")

        # 正文内容
        md_parts.append(content)

        return "\n".join(md_parts)

    def format_json(self, data: Dict[str, Any]) -> str:
        """
        将提取结果格式化为 JSON 字符串

        Args:
            data: 提取的数据字典

        Returns:
            JSON 格式的字符串
        """
        return json.dumps(data, ensure_ascii=False, indent=2)
