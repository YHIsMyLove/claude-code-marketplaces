#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Chapter Manager for Novel Projects
小说章节管理工具

Author: Claude Code
License: Apache 2.0
Version: 1.0.0
"""

import os
import sys
import json
import argparse
import re
from pathlib import Path
from datetime import datetime

class ChapterManager:
    """章节管理器"""

    def __init__(self, project_path=None):
        """初始化章节管理器

        Args:
            project_path: 小说项目路径，默认为当前目录下的项目
        """
        self.current_dir = Path.cwd()
        self.project_path = project_path or self._find_project_path()
        self.chapters_dir = self.project_path / "chapters"
        self.database_dir = self.project_path / "database"

    def _find_project_path(self):
        """查找小说项目路径

        Returns:
            Path: 项目路径
        """
        # 检查当前目录是否是小说项目
        if self._is_novel_project(self.current_dir):
            return self.current_dir

        # 查找当前目录下的小说项目
        for item in self.current_dir.iterdir():
            if item.is_dir() and self._is_novel_project(item):
                return item

        raise FileNotFoundError("未找到小说项目目录")

    def _is_novel_project(self, path):
        """检查是否是小说项目目录

        Args:
            path: 路径

        Returns:
            bool: 是否是小说项目
        """
        required_dirs = ["chapters", "database"]
        return all((path / dir_name).exists() for dir_name in required_dirs)

    def list_chapters(self):
        """列出所有章节

        Returns:
            list: 章节文件列表
        """
        if not self.chapters_dir.exists():
            print(f"错误: 章节目录不存在: {self.chapters_dir}")
            return []

        chapters = []
        pattern = re.compile(r'^(\d+)-(.*?)(?:\.md)?$')

        for file_path in sorted(self.chapters_dir.glob("*.md")):
            match = pattern.match(file_path.stem)
            if match:
                chapter_num = int(match.group(1))
                chapter_title = match.group(2)
                file_size = file_path.stat().st_size
                modified_time = datetime.fromtimestamp(file_path.stat().st_mtime)

                chapters.append({
                    'number': chapter_num,
                    'title': chapter_title,
                    'filename': file_path.name,
                    'path': file_path,
                    'size': file_size,
                    'modified': modified_time
                })

        return sorted(chapters, key=lambda x: x['number'])

    def display_chapters(self):
        """显示章节列表"""
        chapters = self.list_chapters()

        if not chapters:
            print("未找到任何章节文件")
            return

        print(f"\n小说项目: {self.project_path.name}")
        print("=" * 60)
        print(f"{'章节号':<8} {'标题':<30} {'文件大小':<10} {'修改时间':<20}")
        print("-" * 60)

        for chapter in chapters:
            size_str = self._format_file_size(chapter['size'])
            time_str = chapter['modified'].strftime("%Y-%m-%d %H:%M")
            print(f"{chapter['number']:<8} {chapter['title']:<30} {size_str:<10} {time_str:<20}")

        print("-" * 60)
        print(f"总计: {len(chapters)} 章")

    def create_chapter(self, chapter_num, title=None):
        """创建新章节

        Args:
            chapter_num: 章节号
            title: 章节标题

        Returns:
            bool: 创建是否成功
        """
        # 验证章节号
        if not isinstance(chapter_num, int) or chapter_num < 1:
            print("错误: 章节号必须是大于0的整数")
            return False

        # 检查章节是否已存在
        chapter_file = self.chapters_dir / f"{chapter_num:02d}-{title or '章节标题'}.md"
        if chapter_file.exists():
            print(f"错误: 章节 {chapter_num} 已存在: {chapter_file}")
            return False

        # 确保章节号连续
        existing_chapters = self.list_chapters()
        if existing_chapters and chapter_num != existing_chapters[-1]['number'] + 1:
            print(f"警告: 章节号不连续。建议使用章节号 {existing_chapters[-1]['number'] + 1}")

        # 获取小说名称
        novel_name = self._get_novel_name()
        chapter_title = title or f"第{chapter_num}章"

        try:
            # 创建章节目录
            self.chapters_dir.mkdir(parents=True, exist_ok=True)

            # 生成章节内容
            content = self._generate_chapter_content(novel_name, chapter_num, chapter_title)

            # 写入文件
            with open(chapter_file, 'w', encoding='utf-8') as f:
                f.write(content)

            print(f"✓ 成功创建章节: {chapter_file.name}")

            # 更新进度文件
            self._update_progress_file()

            return True

        except Exception as e:
            print(f"错误: 创建章节失败: {e}")
            return False

    def rename_chapter(self, chapter_num, new_title):
        """重命名章节

        Args:
            chapter_num: 章节号
            new_title: 新标题

        Returns:
            bool: 重命名是否成功
        """
        # 查找章节文件
        chapter_file = self._find_chapter_file(chapter_num)
        if not chapter_file:
            print(f"错误: 未找到章节 {chapter_num}")
            return False

        # 生成新文件名
        new_filename = f"{chapter_num:02d}-{new_title}.md"
        new_file_path = self.chapters_dir / new_filename

        # 检查新文件名是否已存在
        if new_file_path.exists():
            print(f"错误: 目标文件名已存在: {new_filename}")
            return False

        try:
            # 重命名文件
            chapter_file.rename(new_file_path)
            print(f"✓ 章节 {chapter_num} 已重命名为: {new_title}")

            # 更新进度文件
            self._update_progress_file()

            return True

        except Exception as e:
            print(f"错误: 重命名失败: {e}")
            return False

    def delete_chapter(self, chapter_num):
        """删除章节

        Args:
            chapter_num: 章节号

        Returns:
            bool: 删除是否成功
        """
        chapter_file = self._find_chapter_file(chapter_num)
        if not chapter_file:
            print(f"错误: 未找到章节 {chapter_num}")
            return False

        # 确认删除
        confirm = input(f"确定要删除章节 {chapter_num} ({chapter_file.stem}) 吗? (y/N): ")
        if confirm.lower() not in ['y', 'yes']:
            print("操作已取消")
            return False

        try:
            chapter_file.unlink()
            print(f"✓ 已删除章节 {chapter_num}")

            # 重新编号后续章节
            self._renumber_chapters()

            # 更新进度文件
            self._update_progress_file()

            return True

        except Exception as e:
            print(f"错误: 删除失败: {e}")
            return False

    def reorder_chapters(self, new_order):
        """重新排序章节

        Args:
            new_order: 新的章节顺序 [2, 1, 3] 表示把第2章排到第1位

        Returns:
            bool: 重排是否成功
        """
        chapters = self.list_chapters()
        chapter_map = {ch['number']: ch for ch in chapters}

        # 验证新顺序
        if set(new_order) != set(chapter_map.keys()):
            print("错误: 新顺序必须包含所有现有章节号")
            return False

        try:
            # 创建临时目录
            temp_dir = self.chapters_dir / "temp_reorder"
            temp_dir.mkdir(exist_ok=True)

            # 按新顺序移动文件到临时目录
            for i, old_num in enumerate(new_order, 1):
                old_chapter = chapter_map[old_num]
                new_filename = f"{i:02d}-{old_chapter['title']}.md"
                new_path = temp_dir / new_filename

                # 读取原文件内容并更新章节号
                content = old_chapter['path'].read_text(encoding='utf-8')
                content = self._update_chapter_number(content, i)

                # 写入新文件
                new_path.write_text(content, encoding='utf-8')

            # 删除原文件并移动新文件
            for chapter in chapters:
                chapter['path'].unlink()

            for temp_file in sorted(temp_dir.glob("*.md")):
                temp_file.rename(self.chapters_dir / temp_file.name)

            # 清理临时目录
            temp_dir.rmdir()

            print(f"✓ 章节已重新排序")

            # 更新进度文件
            self._update_progress_file()

            return True

        except Exception as e:
            print(f"错误: 重排失败: {e}")
            return False

    def _find_chapter_file(self, chapter_num):
        """查找章节文件

        Args:
            chapter_num: 章节号

        Returns:
            Path: 章节文件路径，未找到返回None
        """
        pattern = re.compile(f"^{chapter_num:02d}-(.*?)(?:\\.md)?$")

        for file_path in self.chapters_dir.glob("*.md"):
            if pattern.match(file_path.stem):
                return file_path

        return None

    def _get_novel_name(self):
        """获取小说名称

        Returns:
            str: 小说名称
        """
        # 从项目名称推断小说名称
        project_name = self.project_path.name
        if project_name.endswith("_project"):
            return project_name[:-7]  # 移除 "_project" 后缀

        # 从README文件获取
        readme_path = self.project_path / "README.md"
        if readme_path.exists():
            try:
                content = readme_path.read_text(encoding='utf-8')
                match = re.search(r"# (.+)", content)
                if match:
                    return match.group(1)
            except:
                pass

        return "未命名小说"

    def _generate_chapter_content(self, novel_name, chapter_num, title):
        """生成章节内容

        Args:
            novel_name: 小说名称
            chapter_num: 章节号
            title: 章节标题

        Returns:
            str: 章节内容
        """
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        return f"""# 第{chapter_num}章 - {title}

> 小说: {novel_name}
> 创建时间: {current_time}
> 字数: 0

## 章节大纲

### 主要情节
1.
2.
3.

### 场景设置
- **时间**:
- **地点**:
- **人物**:
- **氛围**:

### 对话要点
-
-
-

## 正文

### 开场
[在此处开始写作...]


### 发展
[情节展开...]


### 高潮
[本章高潮部分...]


### 结尾
[本章结尾，为下一章做铺垫...]

## 章节检查

- [ ] 情节发展合理
- [ ] 人物行为符合性格
- [ ] 对话自然流畅
- [ ] 场景描写生动
- [ ] 与前后章节能衔接

## 作者备注

[在此处添加写作过程中的想法、问题或备注]

---

**写作提示**:
1. 保持与前一章的连贯性
2. 注意控制章节长度
3. 确保每个章节都有其存在的意义
4. 适当设置悬念，吸引读者继续阅读
"""

    def _update_chapter_number(self, content, new_num):
        """更新章节内容中的章节号

        Args:
            content: 章节内容
            new_num: 新章节号

        Returns:
            str: 更新后的内容
        """
        # 更新标题中的章节号
        content = re.sub(r"# 第\d+章 -", f"# 第{new_num}章 -", content)

        # 更新正文中的章节号引用
        content = re.sub(r"第\d+章", f"第{new_num}章", content)

        return content

    def _renumber_chapters(self):
        """重新编号章节"""
        chapters = []
        pattern = re.compile(r'^(\d+)-(.*?)(?:\.md)?$')

        # 收集所有章节
        for file_path in sorted(self.chapters_dir.glob("*.md")):
            match = pattern.match(file_path.stem)
            if match:
                chapters.append({
                    'path': file_path,
                    'title': match.group(2),
                    'content': file_path.read_text(encoding='utf-8')
                })

        # 重新编号
        for i, chapter in enumerate(chapters, 1):
            new_filename = f"{i:02d}-{chapter['title']}.md"
            new_path = self.chapters_dir / new_filename

            # 更新内容中的章节号
            updated_content = self._update_chapter_number(chapter['content'], i)

            # 写入新文件
            new_path.write_text(updated_content, encoding='utf-8')

            # 删除原文件（如果文件名有变化）
            if new_path != chapter['path']:
                chapter['path'].unlink()

    def _update_progress_file(self):
        """更新进度文件"""
        progress_file = self.database_dir / "主线-支线-进度.md"
        if not progress_file.exists():
            return

        try:
            content = progress_file.read_text(encoding='utf-8')
            chapters = self.list_chapters()

            # 查找章节进度表格
            table_pattern = r"(\| 章节 \| 标题 \| 状态 \| 字数 \| 完成时间 \|\n\|.*?\n)"

            new_table_rows = ["| 章节 | 标题 | 状态 | 字数 | 完成时间 |",
                             "|------|------|------|------|----------|"]

            for chapter in chapters:
                status = "✅ 已完成" if chapter['size'] > 1000 else "📝 写作中"
                word_count = self._estimate_word_count(chapter['path'])
                completed_time = "✅ 已完成" if chapter['size'] > 1000 else "-"

                new_table_rows.append(
                    f"| {chapter['number']} | {chapter['title']} | {status} | {word_count} | {completed_time} |"
                )

            new_table = "\n".join(new_table_rows) + "\n"

            # 替换原有表格
            new_content = re.sub(table_pattern, new_table, content, flags=re.DOTALL)

            # 如果没有找到表格，在文件末尾添加
            if new_content == content:
                new_content += f"\n## 章节进度\n{new_table}\n"

            progress_file.write_text(new_content, encoding='utf-8')

        except Exception as e:
            print(f"警告: 无法更新进度文件: {e}")

    def _estimate_word_count(self, file_path):
        """估算文件字数

        Args:
            file_path: 文件路径

        Returns:
            str: 字数统计
        """
        try:
            content = file_path.read_text(encoding='utf-8')
            # 移除Markdown标记，只计算正文内容
            text = re.sub(r'[#*`\[\]()_-]', '', content)
            text = re.sub(r'\n+', ' ', text)
            return str(len(text.replace(' ', '')))
        except:
            return "未知"

    def _format_file_size(self, size):
        """格式化文件大小

        Args:
            size: 文件大小（字节）

        Returns:
            str: 格式化后的大小
        """
        if size < 1024:
            return f"{size}B"
        elif size < 1024 * 1024:
            return f"{size//1024}KB"
        else:
            return f"{size//(1024*1024)}MB"

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="小说章节管理工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python chapter_manager.py list                    # 列出所有章节
  python chapter_manager.py create 4 "新章节标题"    # 创建第4章
  python chapter_manager.py rename 1 "新标题"        # 重命名第1章
  python chapter_manager.py delete 2                 # 删除第2章
        """
    )

    parser.add_argument(
        "action",
        choices=["list", "create", "rename", "delete", "reorder"],
        help="操作类型"
    )

    parser.add_argument(
        "chapter_num",
        type=int,
        nargs="?",
        help="章节号"
    )

    parser.add_argument(
        "title",
        nargs="?",
        help="章节标题或新标题"
    )

    parser.add_argument(
        "--project",
        help="指定项目路径"
    )

    parser.add_argument(
        "--version",
        action="version",
        version="Chapter Manager v1.0.0"
    )

    args = parser.parse_args()

    try:
        manager = ChapterManager(args.project)

        if args.action == "list":
            manager.display_chapters()

        elif args.action == "create":
            if not args.chapter_num:
                print("错误: 创建章节需要指定章节号")
                sys.exit(1)
            success = manager.create_chapter(args.chapter_num, args.title)
            sys.exit(0 if success else 1)

        elif args.action == "rename":
            if not args.chapter_num or not args.title:
                print("错误: 重命名章节需要指定章节号和新标题")
                sys.exit(1)
            success = manager.rename_chapter(args.chapter_num, args.title)
            sys.exit(0 if success else 1)

        elif args.action == "delete":
            if not args.chapter_num:
                print("错误: 删除章节需要指定章节号")
                sys.exit(1)
            success = manager.delete_chapter(args.chapter_num)
            sys.exit(0 if success else 1)

        elif args.action == "reorder":
            print("重排功能需要交互式操作，暂未实现")
            sys.exit(1)

    except KeyboardInterrupt:
        print("\n操作被用户取消")
        sys.exit(1)
    except Exception as e:
        print(f"程序执行出错: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()