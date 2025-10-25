#!/usr/bin/env python3
"""
极简 Markdown 转 Word 转换器
只做两件事：1. 用 mmdc 转换 mermaid 图表，2. 用 pandoc 转换为 docx
"""

import subprocess
import sys
import argparse
import platform
from pathlib import Path


def has_mermaid_diagrams(input_file):
    """检查文件是否包含 mermaid 代码块"""
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
            return '```mermaid' in content
    except Exception as e:
        raise RuntimeError(f"读取文件失败: {e}")


def convert_mermaid(input_file, output_file):
    """第一步：使用 mmdc 转换 mermaid 图表"""
    mmdc_cmd = 'mmdc.cmd' if platform.system() == 'Windows' else 'mmdc'
    cmd = [mmdc_cmd, '-i', str(input_file), '-o', str(output_file), '-e', 'png', '-s', '2']

    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"✓ Mermaid 图表已转换: {output_file}")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"mmdc 转换失败:\n{e.stderr}") from e
    except FileNotFoundError:
        mmdc_name = "mmdc.cmd" if platform.system() == 'Windows' else "mmdc"
        raise RuntimeError(f"{mmdc_name} 命令未找到，请安装 mermaid-cli: npm install -g @mermaid-js/mermaid-cli")


def convert_to_docx(input_file, output_file, template_file=None):
    """第二步：使用 pandoc 转换为 docx"""
    cmd = ['pandoc', str(input_file), '-o', str(output_file), '--toc']

    # 添加模板支持
    if template_file and Path(template_file).exists():
        cmd.extend(['--reference-doc', str(template_file)])
        print(f"✓ 使用模板文件: {template_file}")
    elif Path('assets/template.docx').exists():
        cmd.extend(['--reference-doc', 'assets/template.docx'])
        print("✓ 使用默认模板: assets/template.docx")

    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(f"✓ Word 文档已生成: {output_file}")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"pandoc 转换失败:\n{e.stderr}") from e
    except FileNotFoundError:
        raise RuntimeError("pandoc 命令未找到，请安装 pandoc: https://pandoc.org/installing.html")


def convert(input_file, output_file=None, template_file=None):
    """主转换函数：协调两步转换流程"""
    input_path = Path(input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {input_file}")

    if not input_path.suffix.lower() == '.md':
        raise ValueError("只支持 .md 文件")

    # 生成输出文件名
    if output_file is None:
        output_file = input_path.with_suffix('.docx')

    output_path = Path(output_file)

    print(f"开始转换: {input_path} → {output_path}")

    # 检查是否包含 mermaid 图表
    has_mermaid = has_mermaid_diagrams(input_path)

    if has_mermaid:
        print("检测到 Mermaid 图表，执行两步转换...")
        # 第一步：转换 mermaid 图表
        intermediate_md = input_path.parent / f"{input_path.stem}.mmdc.md"
        convert_mermaid(input_path, intermediate_md)

        # 第二步：转换为 docx
        convert_to_docx(intermediate_md, output_path, template_file)

        print(f"📄 中间文件已保留: {intermediate_md}")
    else:
        print("未检测到 Mermaid 图表，直接转换...")
        # 直接转换
        convert_to_docx(input_path, output_path, template_file)

    print(f"🎉 转换完成! 输出文件: {output_path}")
    return str(output_path)


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description="极简 Markdown 转 Word 转换器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python convert.py README.md
  python convert.py README.md output.docx
        """
    )

    parser.add_argument('input_file', help='输入的 Markdown 文件 (.md)')
    parser.add_argument('output_file', nargs='?', help='输出的 Word 文件 (.docx，可选)')
    parser.add_argument('--template', help='Word 模板文件 (.docx，可选)')

    args = parser.parse_args()

    try:
        convert(args.input_file, args.output_file, args.template)
    except Exception as e:
        print(f"❌ 错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()