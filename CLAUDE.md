# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个个人维护的 Claude Code 插件市场，为 Claude Code 用户提供额外的技能和功能扩展。项目结构为技能仓库形式，支持通过 Claude Code 的插件系统安装和使用。

## 项目架构

### 核心结构
```
claude-code-marketplaces/
├── marketplace.json      # 插件市场配置文件
├── README.md            # 项目说明文档
├── skills/              # 技能目录
│   └── [skill-name]/    # 单个技能文件夹
│       ├── SKILL.md     # 技能定义文件（必需）
│       ├── scripts/     # 执行脚本（可选）
│       ├── assets/      # 资源文件（可选）
│       └── references/  # 参考文档（可选）
└── docs/                # 项目文档
```

### 技能规范
每个技能必须包含 `SKILL.md` 文件，定义技能的：
- 名称和描述
- 许可证信息
- 允许使用的工具列表
- 使用说明和前置条件
- 文件结构

## 常用命令

### 插件安装
```bash
# 添加整个插件市场
/plugin marketplace add YHIsMyLove/claude-code-marketplaces

# 直接安装特定技能
/plugin install markdown-word-converter@YHIsMyLove/claude-code-marketplaces
/plugin install novel-project@YHIsMyLove/claude-code-marketplaces
```

### 技能测试和开发
```bash
# 测试 markdown-word-converter 技能
python skills/markdown-word-converter/scripts/convert.py test_document.md

# 安装技能依赖
python skills/markdown-word-converter/scripts/install_dependencies.py

# 测试 novel-project 技能
python skills/novel-project/scripts/create_novel.py "测试小说"

# 章节管理测试
python skills/novel-project/scripts/chapter_manager.py --help
```

## 开发指南

### 添加新技能
1. 在 `skills/` 目录下创建新文件夹
2. 编写符合规范的 `SKILL.md` 文件
3. 添加必要的脚本和资源文件
4. 更新 `marketplace.json` 配置
5. 测试技能功能

### 技能开发要求
- 技能名称使用小写字母和连字符
- 确保脚本跨平台兼容（Windows/macOS/Linux）
- 提供完整的依赖检查和错误处理
- 包含详细的使用示例和故障排除文档

## 现有技能

### markdown-word-converter
- **功能**: 将 Markdown 转换为 Word 文档，支持 Mermaid 图表
- **依赖**: Python 3.6+, Pandoc, Mermaid CLI
- **主要脚本**: `scripts/convert.py`, `scripts/install_dependencies.py`
- **模板文件**: `assets/template.docx`
- **核心流程**: 使用 mmdc 转换 Mermaid 图表 → 使用 pandoc 转换文档

### novel-project
- **功能**: 创建小说工程目录结构和基础模板，支持章节管理
- **依赖**: Python 3.6+
- **主要脚本**: `scripts/create_novel.py`, `scripts/chapter_manager.py`
- **配置文件**: `assets/config.json`
- **模板目录**: `assets/templates/`
- **生成结构**: 包含 database/、chapters/、assets/ 的完整小说项目

## 关键架构理解

### 技能加载机制
- 通过 `marketplace.json` 定义技能清单
- 每个技能的 `SKILL.md` 文件包含 Claude Code 解析的元数据
- `allowed-tools` 字段限制技能可使用的工具范围

### 技能开发模式
- **极简主义**: markdown-word-converter 采用两步转换流程
- **模板驱动**: novel-project 使用配置文件和模板系统
- **错误处理**: 所有技能都包含完整的依赖检查和用户友好的错误信息
- **跨平台**: 脚本自动检测操作系统并适配命令行工具

## 注意事项

- 本插件市场为社区维护，非 Anthropic 官方支持
- 使用技能前请确保依赖正确安装
- 技能遵循各自的许可证条款
- 开发时请参考 Claude Code 官方技能文档规范
- Python 脚本需要 shebang 行支持跨平台执行
- 依赖安装脚本应提供详细的安装指导