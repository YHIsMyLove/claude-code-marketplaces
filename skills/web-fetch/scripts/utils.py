#!/usr/bin/env python3
"""
Web Fetch - 工具函数模块
"""

import re
from typing import Optional
from urllib.parse import urlparse


def is_valid_url(url: str) -> bool:
    """
    验证 URL 是否有效

    Args:
        url: 待验证的 URL 字符串

    Returns:
        URL 是否有效
    """
    try:
        result = urlparse(url)
        return all([result.scheme in ["http", "https"], result.netloc])
    except Exception:
        return False


def sanitize_filename(filename: str, max_length: int = 255) -> str:
    """
    清理文件名，移除非法字符

    Args:
        filename: 原始文件名
        max_length: 最大长度

    Returns:
        清理后的文件名
    """
    # 移除或替换非法字符
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    # 移除控制字符
    filename = re.sub(r'[\x00-\x1f\x7f]', '', filename)
    # 限制长度
    if len(filename) > max_length:
        filename = filename[:max_length]
    # 移除首尾空格和点
    filename = filename.strip('. ')
    return filename or "output"


def generate_output_filename(url: str, title: Optional[str] = None, extension: str = "md") -> str:
    """
    根据URL和标题生成输出文件名

    Args:
        url: 网页 URL
        title: 网页标题（可选）
        extension: 文件扩展名

    Returns:
        生成的文件名
    """
    if title:
        base_name = sanitize_filename(title)
    else:
        # 从 URL 提取文件名
        parsed = urlparse(url)
        path = parsed.path.rstrip("/")
        base_name = path.split("/")[-1] or "output"
        base_name = sanitize_filename(base_name)

    return f"{base_name}.{extension}"
