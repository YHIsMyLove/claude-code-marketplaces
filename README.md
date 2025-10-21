# Claude Code Plugin Marketplace

这是个人维护的 Claude Code 插件市场，包含自定义技能和工具。

## 插件市场介绍

本插件市场为 Claude Code 用户提供额外的技能和功能扩展，旨在提升开发效率和用户体验。

## 可用技能

### markdown-word-converter

**功能**: 将 Markdown 文件转换为 Word 文档，支持 Mermaid 图表、自定义模板和目录生成。

**主要特性**:
- 跨平台支持 (Windows/macOS/Linux)
- Mermaid 图表自动转换为 PNG 图片
- 支持自定义 Word 模板
- 自动生成目录
- 完整的依赖检查和安装指导

**使用场景**:
- 技术文档转换为 Word 格式
- 报告生成和格式化
- 包含图表的文档处理
- 企业级文档标准化输出

## 安装方法

### 方法一：通过 Claude Code 添加插件市场

```bash
/plugin marketplace add YHIsMyLove/claude-code-marketplaces
```

然后选择要安装的技能。

### 方法二：直接安装技能

```bash
/plugin install markdown-word-converter@YHIsMyLove/claude-code-marketplaces
```

## 使用指南

安装技能后，您可以直接在 Claude Code 中使用相关功能。例如：

"请使用 markdown-word-converter 技能将我的技术文档转换为 Word 格式"

## 技能要求

### markdown-word-converter 依赖

- **Python 3.6+**
- **Pandoc** - 文档转换工具
- **Mermaid CLI** - 图表转换工具

安装依赖的方法请参考各技能的详细文档。

## 贡献指南

欢迎贡献新的技能和改进！请遵循以下步骤：

1. Fork 本仓库
2. 创建技能分支
3. 添加您的技能到 `skills/` 目录
4. 确保技能符合 Claude Code 技能规范
5. 提交 Pull Request

## 技能开发规范

每个技能应包含：
- `SKILL.md` - 技能定义和说明
- `scripts/` - 执行脚本（可选）
- `assets/` - 资源文件（可选）
- `references/` - 参考文档（可选）

详细规范请参考 [Claude Code 技能文档](https://docs.claude.com/en/articles/12512198-creating-custom-skills)。

## 许可证

本插件市场中的技能遵循各自的许可证条款。请在使用前查看每个技能的许可证信息。

## 支持

如果您遇到问题或有建议，请：
1. 查看技能的详细文档
2. 在 GitHub 上提交 Issue
3. 检查依赖是否正确安装

## 更新日志

### v1.0.0 (2025-10-21)
- 初始版本发布
- 包含 markdown-word-converter 技能
- 完整的文档和使用指南

---

**注意**: 本插件市场由社区维护，不是 Anthropic 官方支持的产品。使用时请自行评估风险。