#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Novel Project Creator
创建小说工程目录结构和基础模板

Author: Claude Code
License: Apache 2.0
Version: 1.0.0
"""

import os
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime
import re

class NovelProjectCreator:
    """小说项目创建器"""

    def __init__(self, config_path=None):
        """初始化创建器

        Args:
            config_path: 配置文件路径，默认使用 assets/config.json
        """
        self.script_dir = Path(__file__).parent
        self.config_path = config_path or self.script_dir.parent / "assets" / "config.json"
        self.templates_dir = self.script_dir.parent / "assets" / "templates"
        self.config = self._load_config()

    def _load_config(self):
        """加载配置文件"""
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return self._get_default_config()
        except Exception as e:
            print(f"警告: 无法加载配置文件，使用默认配置: {e}")
            return self._get_default_config()

    def _get_default_config(self):
        """获取默认配置"""
        return {
            "project_suffix": "_project",
            "database_folders": [
                "参考资料",
                "背景设定",
                "人物设定",
                "写作风格",
                "主线-支线-进度"
            ],
            "chapter_count": 3,
            "chapter_prefix": "",
            "encoding": "utf-8",
            "create_readme": True,
            "create_assets": True
        }

    def validate_novel_name(self, name):
        """验证小说名称的合法性

        Args:
            name: 小说名称

        Returns:
            bool: 名称是否有效
        """
        if not name or not name.strip():
            return False

        # 检查是否包含无效字符
        invalid_chars = ['<', '>', ':', '"', '|', '?', '*', '/', '\\']
        for char in invalid_chars:
            if char in name:
                return False

        # 检查长度
        if len(name.strip()) > 50:
            return False

        return True

    def generate_project_name(self, novel_name):
        """生成项目文件夹名称

        Args:
            novel_name: 小说名称

        Returns:
            str: 项目文件夹名称
        """
        suffix = self.config.get("project_suffix", "_project")
        clean_name = novel_name.strip()
        return f"{clean_name}{suffix}"

    def create_directory_structure(self, project_path):
        """创建项目目录结构

        Args:
            project_path: 项目根目录路径
        """
        directories = [
            project_path / "database",
            project_path / "chapters",
            project_path / "assets"
        ]

        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            print(f"✓ 创建目录: {directory}")

    def create_template_files(self, project_path, novel_name):
        """创建模板文件

        Args:
            project_path: 项目根目录路径
            novel_name: 小说名称
        """
        database_path = project_path / "database"

        # 创建数据库文件夹中的文件
        database_files = self.config.get("database_folders", [])
        for folder_name in database_files:
            template_file = database_path / f"{folder_name}.md"
            if not template_file.exists():
                content = self._generate_database_template(folder_name, novel_name)
                self._write_file(template_file, content)
                print(f"✓ 创建模板文件: {template_file}")

    def create_chapter_files(self, project_path, novel_name):
        """创建章节文件

        Args:
            project_path: 项目根目录路径
            novel_name: 小说名称
        """
        chapters_path = project_path / "chapters"
        chapter_count = self.config.get("chapter_count", 3)

        for i in range(1, chapter_count + 1):
            chapter_num = str(i).zfill(2)
            chapter_file = chapters_path / f"{chapter_num}-章节标题.md"
            if not chapter_file.exists():
                content = self._generate_chapter_template(novel_name, i)
                self._write_file(chapter_file, content)
                print(f"✓ 创建章节文件: {chapter_file}")

    def create_readme(self, project_path, novel_name):
        """创建项目说明文件

        Args:
            project_path: 项目根目录路径
            novel_name: 小说名称
        """
        if not self.config.get("create_readme", True):
            return

        readme_path = project_path / "README.md"
        content = self._generate_readme_content(novel_name)
        self._write_file(readme_path, content)
        print(f"✓ 创建项目说明: {readme_path}")

    def _generate_database_template(self, folder_name, novel_name):
        """生成数据库模板内容

        Args:
            folder_name: 文件夹名称
            novel_name: 小说名称

        Returns:
            str: 模板内容
        """
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        templates = {
            "参考资料": f"""# 参考资料 - {novel_name}

> 创建时间: {current_time}

## 历史背景

- 在这里记录与小说相关的历史背景资料
- 包括时代背景、社会环境、文化习俗等

## 文学作品

- 列出参考的文学作品和相关资料
- 记录灵感和借鉴来源

## 研究资料

- 专业书籍、学术论文、研究报告等
- 为小说创作提供事实依据的材料

## 图片资源

- 收集的图片、插画、地图等视觉资料
- 有助于场景构建和人物设计的素材

## 网络资源

- 有用的网站链接、在线资料等
- 便于随时查阅的网络资源

---

**提示**: 请在创作过程中持续更新和补充参考资料，确保内容的准确性和丰富性。
""",
            "背景设定": f"""# 背景设定 - {novel_name}

> 创建时间: {current_time}

## 世界观设定

### 时代背景
- **时代**: [填写时代背景]
- **地理位置**: [填写主要发生的地点]
- **社会结构**: [描述社会阶层和政治结构]

### 地理环境
- **主要地点**:
  - [地点1]: 简要描述
  - [地点2]: 简要描述
- **气候特点**:
- **自然景观**:

### 文化习俗
- **宗教信仰**:
- **节日庆典**:
- **生活习惯**:
- **价值观念**:

## 技术水平

- **科技发展**:
- **交通工具**:
- **通讯方式**:
- **武器装备**:

## 经济体系

- **主要产业**:
- **贸易方式**:
- **货币制度**:
- **社会分配**:

---

**提示**: 背景设定是小说的基础，请详细构思，确保内部逻辑的一致性。
""",
            "人物设定": f"""# 人物设定 - {novel_name}

> 创建时间: {current_time}

## 主要人物

### 主角
**姓名**: [主角姓名]
**年龄**:
**性别**:
**职业**:
**外貌特征**:
**性格特点**:
**背景故事**:
**目标动机**:
**能力特长**:
**缺点弱点**:

### 配角
**姓名**: [配角姓名]
**年龄**:
**性别**:
**与主角关系**:
**性格特点**:
**作用**:

## 次要人物

- **人物1**: 简要描述
- **人物2**: 简要描述
- **人物3**: 简要描述

## 人物关系图

- **主角与配角1**: [关系描述]
- **主角与配角2**: [关系描述]
- **配角之间的关系**: [关系描述]

## 人物发展轨迹

### 主角成长线
1. **初始状态**:
2. **转折点1**:
3. **转折点2**:
4. **最终状态**:

---

**提示**: 人物设定要立体饱满，避免扁平化。每个人物都应该有自己的动机和成长轨迹。
""",
            "写作风格": f"""# 写作风格 - {novel_name}

> 创建时间: {current_time}

## 文学风格

### 叙事视角
- **主要视角**: [第一人称/第三人称/全知视角]
- **视角转换**:
- **叙事口吻**:

### 语言风格
- **正式程度**: [非常正式/较正式/口语化/非常口语化]
- **修辞手法**:
  - 比喻:
  - 拟人:
  - 排比:
  - 其他:
- **句式特点**:

### 节奏控制
- **整体节奏**: [快节奏/中等节奏/慢节奏]
- **章节节奏**:
- **场景转换**:

## 情感基调

### 主要情感
- **整体基调**: [轻松愉快/沉重严肃/悬疑紧张/温馨感人]
- **情感变化**:

### 氛围营造
- **主要场景氛围**:
- **季节氛围**:
- **时间氛围**:

## 主题思想

### 核心主题
1. **主题1**: [描述]
2. **主题2**: [描述]
3. **主题3**: [描述]

### 价值观导向
- **提倡的价值观**:
- **反对的价值观**:

---

**提示**: 写作风格要贯穿全文，保持一致性。在创作过程中可以适当调整，但不要频繁改变。
""",
            "主线-支线-进度": f"""# 主线-支线-进度 - {novel_name}

> 创建时间: {current_time}
> 最后更新: {current_time}

## 主线剧情

### 核心冲突
**主要矛盾**: [描述故事的核心冲突]

### 情节大纲

#### 第一幕：开端
- **起因**:
- **背景介绍**:
- **冲突建立**:

#### 第二幕：发展
- **矛盾激化**:
- **关键转折**:
- **高潮前奏**:

#### 第三幕：高潮与结局
- **高潮**:
- **结局**:
- **主题升华**:

## 支线剧情

### 支线1: [支线名称]
- **起因**:
- **发展**:
- **结局**:
- **与主线关系**:

### 支线2: [支线名称]
- **起因**:
- **发展**:
- **结局**:
- **与主线关系**:

## 创作进度

### 章节进度
| 章节 | 标题 | 状态 | 字数 | 完成时间 |
|------|------|------|------|----------|
| 第一章 | 章节标题 | 📝 计划中 | - | - |
| 第二章 | 章节标题 | 📝 计划中 | - | - |
| 第三章 | 章节标题 | 📝 计划中 | - | - |

### 总体进度
- **总字数目标**:
- **已完成字数**:
- **完成百分比**: 0%
- **预计完成时间**:

### 里程碑
- [ ] 大纲定稿
- [ ] 第一章完成
- [ ] 前五章完成
- [ ] 第一稿完成
- [ ] 修改完成
- [ ] 最终定稿

## 修改记录

| 日期 | 修改内容 | 版本 |
|------|----------|------|
| {current_time} | 创建项目 | v1.0 |

---

**提示**: 定期更新进度记录，有助于保持创作动力和时间管理。
"""
        }

        return templates.get(folder_name, f"""# {folder_name} - {novel_name}

> 创建时间: {current_time}

请在此处添加{folder_name}相关的内容。

---

**提示**: 根据需要自定义内容结构。
""")

    def _generate_chapter_template(self, novel_name, chapter_num):
        """生成章节模板内容

        Args:
            novel_name: 小说名称
            chapter_num: 章节编号

        Returns:
            str: 模板内容
        """
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        return f"""# 第{chapter_num}章 - 章节标题

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

    def _generate_readme_content(self, novel_name):
        """生成README内容

        Args:
            novel_name: 小说名称

        Returns:
            str: README内容
        """
        current_time = datetime.now().strftime("%Y-%m-%d")
        project_name = self.generate_project_name(novel_name)

        return f"""# {novel_name}

> 创建时间: {current_time}
> 项目类型: 小说创作工程

## 项目简介

这是一个使用 Novel Project 技能创建的小说工程，旨在提供结构化的写作环境和完整的创作模板。

## 目录结构

```
{project_name}/
├── database/                   # 数据库目录
│   ├── 参考资料.md            # 参考资料
│   ├── 背景设定.md            # 世界观和背景设定
│   ├── 人物设定.md            # 主要人物设定
│   ├── 写作风格.md            # 写作风格指导
│   └── 主线-支线-进度.md       # 情节大纲和进度跟踪
├── chapters/                  # 章节目录
│   ├── 01-章节标题.md         # 第一章
│   ├── 02-章节标题.md         # 第二章
│   └── 03-章节标题.md         # 第三章
├── assets/                    # 资源文件目录
└── README.md                  # 项目说明文件
```

## 使用指南

### 1. 准备工作
- 在 `database/人物设定.md` 中设定主要角色
- 在 `database/背景设定.md` 中构建世界观
- 在 `database/写作风格.md` 中确定写作风格

### 2. 章节写作
- 在 `chapters/` 目录下找到对应的章节文件
- 按照模板格式进行写作
- 定期更新 `database/主线-支线-进度.md` 中的进度

### 3. 进度管理
- 使用 `主线-支线-进度.md` 跟踪写作进度
- 定期备份重要的写作内容
- 可以使用 Git 等版本控制工具管理项目

## 写作建议

1. **先规划后写作**: 在开始写作前，先完善人物设定和背景设定
2. **保持一致性**: 确保人物性格、世界观设定在全书中保持一致
3. **定期回顾**: 定期检查和更新进度，保持创作动力
4. **多写多改**: 初稿完成后，要进行多次修改和完善

## 备份建议

- 定期备份整个项目文件夹
- 可以考虑使用云存储或版本控制系统
- 重要章节建议单独备份

## 技术支持

本项目由 Claude Code 的 novel-project 技能创建，如需帮助请参考技能文档。

---

**祝你写作顺利！** 📚✍️
"""

    def _write_file(self, file_path, content):
        """写入文件

        Args:
            file_path: 文件路径
            content: 文件内容
        """
        try:
            encoding = self.config.get("encoding", "utf-8")
            with open(file_path, 'w', encoding=encoding) as f:
                f.write(content)
        except Exception as e:
            print(f"错误: 无法写入文件 {file_path}: {e}")
            raise

    def create_project(self, novel_name):
        """创建小说项目

        Args:
            novel_name: 小说名称

        Returns:
            bool: 创建是否成功
        """
        # 验证小说名称
        if not self.validate_novel_name(novel_name):
            print("错误: 小说名称无效。请使用字母、数字和中文，避免特殊字符。")
            return False

        # 生成项目名称
        project_name = self.generate_project_name(novel_name)
        project_path = Path.cwd() / project_name

        # 检查项目是否已存在
        if project_path.exists():
            print(f"错误: 项目目录 '{project_name}' 已存在。")
            print("请选择不同的小说名称或删除现有目录。")
            return False

        try:
            print(f"正在创建小说项目: {novel_name}")
            print(f"项目路径: {project_path}")
            print("-" * 50)

            # 创建目录结构
            self.create_directory_structure(project_path)

            # 创建模板文件
            self.create_template_files(project_path, novel_name)

            # 创建章节文件
            self.create_chapter_files(project_path, novel_name)

            # 创建README
            self.create_readme(project_path, novel_name)

            print("-" * 50)
            print(f"✓ 小说项目 '{novel_name}' 创建成功！")
            print(f"✓ 项目位置: {project_path}")
            print(f"✓ 请查看 {project_path}/README.md 了解使用方法")

            return True

        except Exception as e:
            print(f"错误: 创建项目失败: {e}")
            # 清理可能创建的不完整文件
            try:
                if project_path.exists():
                    import shutil
                    shutil.rmtree(project_path)
                    print(f"已清理不完整的项目目录: {project_path}")
            except Exception as cleanup_error:
                print(f"清理失败: {cleanup_error}")
            return False

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="创建小说工程目录结构和基础模板",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python create_novel.py "三国演义"
  python create_novel.py "红楼梦" --config custom_config.json
        """
    )

    parser.add_argument(
        "novel_name",
        help="小说名称"
    )

    parser.add_argument(
        "--config",
        help="配置文件路径",
        default=None
    )

    parser.add_argument(
        "--version",
        action="version",
        version="Novel Project Creator v1.0.0"
    )

    args = parser.parse_args()

    try:
        creator = NovelProjectCreator(args.config)
        success = creator.create_project(args.novel_name)
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n操作被用户取消")
        sys.exit(1)
    except Exception as e:
        print(f"程序执行出错: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()