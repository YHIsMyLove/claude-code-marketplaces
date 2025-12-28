#!/usr/bin/env python3
"""
Web Fetch - 命令行接口
"""

import argparse
import sys
from pathlib import Path
from fetch import WebFetcher
from utils import generate_output_filename


def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="Web Fetch - 智能网页内容提取工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 基本用法
  python cli.py https://example.com/article

  # 指定输出文件
  python cli.py https://example.com/article -o article.md

  # JSON 格式输出
  python cli.py https://example.com/article --format json

  # 包含图片和链接
  python cli.py https://example.com/article --include-images --include-links
        """
    )

    parser.add_argument(
        "url",
        help="要提取的网页 URL"
    )

    parser.add_argument(
        "-o", "--output",
        help="输出文件路径（不指定则打印到标准输出）",
        type=str
    )

    parser.add_argument(
        "-f", "--format",
        choices=["markdown", "json", "html", "txt"],
        default="markdown",
        help="输出格式（默认: markdown）"
    )

    parser.add_argument(
        "--no-metadata",
        action="store_true",
        help="不包含元数据"
    )

    parser.add_argument(
        "--include-comments",
        action="store_true",
        help="包含文章评论"
    )

    parser.add_argument(
        "--no-tables",
        action="store_true",
        help="不包含表格"
    )

    parser.add_argument(
        "--include-images",
        action="store_true",
        help="包含图片链接"
    )

    parser.add_argument(
        "--include-links",
        action="store_true",
        help="包含链接"
    )

    parser.add_argument(
        "-l", "--language",
        help="目标语言代码（如 zh, en）"
    )

    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="显示详细输出"
    )

    return parser.parse_args()


def main():
    """主函数"""
    args = parse_arguments()

    try:
        # 创建提取器实例
        fetcher = WebFetcher(
            output_format=args.format,
            include_metadata=not args.no_metadata,
            include_comments=args.include_comments,
            include_tables=not args.no_tables,
            include_images=args.include_images,
            include_links=args.include_links,
            target_language=args.language,
        )

        if args.verbose:
            print(f"[INFO] 正在提取: {args.url}", file=sys.stderr)

        # 执行提取
        result = fetcher.extract(args.url)

        if result is None:
            print(f"[ERROR] 无法提取网页内容", file=sys.stderr)
            sys.exit(1)

        # 格式化输出
        if args.format == "markdown":
            output = fetcher.format_markdown(result)
        elif args.format == "json":
            output = fetcher.format_json(result)
        else:
            output = result["content"]

        # 输出结果
        if args.output:
            output_path = Path(args.output)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(output, encoding="utf-8")
            if args.verbose:
                print(f"[INFO] 已保存到: {args.output}", file=sys.stderr)
        else:
            print(output)

        # 显示元数据摘要
        if args.verbose and result.get("metadata"):
            metadata = result["metadata"]
            print(f"\n[INFO] 提取摘要:", file=sys.stderr)
            print(f"  标题: {metadata.get('title', '未知')}", file=sys.stderr)
            print(f"  作者: {metadata.get('author', '未知')}", file=sys.stderr)
            if metadata.get("date"):
                print(f"  日期: {metadata['date']}", file=sys.stderr)
            print(f"  内容长度: {len(result['content'])} 字符", file=sys.stderr)

        sys.exit(0)

    except ValueError as e:
        print(f"[ERROR] 参数错误: {e}", file=sys.stderr)
        sys.exit(1)
    except RuntimeError as e:
        print(f"[ERROR] 运行时错误: {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print(f"\n[WARN] 用户中断", file=sys.stderr)
        sys.exit(130)
    except Exception as e:
        print(f"[ERROR] 未知错误: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
