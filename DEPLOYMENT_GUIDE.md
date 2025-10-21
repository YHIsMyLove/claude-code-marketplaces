# 部署指导

## 已完成的工作

✅ **技能创建完成**
- 创建了完整的 `markdown-word-converter` 技能
- 实现了跨平台的 Python 转换脚本
- 包含完整的依赖检查和安装指导
- 提供了详细的使用文档和示例

✅ **插件市场结构完成**
- 创建了标准的插件市场目录结构
- 包含所有必要的法律文档（LICENSE.txt, THIRD_PARTY_NOTICES.md）
- 提供了完整的 README.md 文档

✅ **Git 仓库设置完成**
- 代码已成功推送到 GitHub：https://github.com/YHIsMyLove/claude-code-marketplaces.git
- 仓库包含完整的提交历史和文档

## 下一步操作

### 1. 添加插件市场到 Claude Code

在 Claude Code 中执行以下命令：

```bash
/plugin marketplace add YHIsMyLove/claude-code-marketplaces
```

### 2. 安装技能

添加插件市场后，您可以：

**方法一：通过界面选择**
1. 选择 "Browse and install plugins"
2. 选择 "YHIsMyLove/claude-code-marketplaces"
3. 选择 "markdown-word-converter"
4. 选择 "Install now"

**方法二：直接安装**
```bash
/plugin install markdown-word-converter@YHIsMyLove/claude-code-marketplaces
```

### 3. 验证安装

安装完成后，您可以通过以下方式验证：

1. **检查技能列表**：技能应该出现在可用技能列表中
2. **测试转换功能**：
   ```
   请使用 markdown-word-converter 技能将我的文档转换为 Word 格式
   ```
3. **检查依赖**：技能会自动检查所需依赖（pandoc, mmdc）

## 技能使用指南

### 基本使用

安装技能后，您可以直接使用：

```
请使用 markdown-word-converter 技能将 document.md 转换为 Word 文档
```

### 高级功能

```
请使用 markdown-word-converter 技能：
1. 转换包含 Mermaid 图表的技术文档
2. 使用自定义模板
3. 生成目录结构
```

## 依赖安装

如果系统提示缺少依赖，请安装：

### Pandoc
- **Windows**: 从 https://pandoc.org/installing.html 下载安装程序
- **macOS**: `brew install pandoc`
- **Linux**: `sudo apt install pandoc`

### Mermaid CLI
```bash
npm install -g @mermaid-js/mermaid-cli
```

## 故障排除

### 常见问题

1. **技能安装失败**
   - 检查网络连接
   - 确认 GitHub 仓库地址正确

2. **依赖检查失败**
   - 确保已安装 Python 3.6+
   - 安装所需的系统工具（pandoc, mmdc）

3. **转换失败**
   - 检查 Markdown 文件语法
   - 确认 Mermaid 图表语法正确
   - 验证模板文件存在

### 获取帮助

如果遇到问题：
1. 查看技能的详细文档 (`references/usage_examples.md`)
2. 检查依赖安装脚本 (`scripts/install_dependencies.py`)
3. 在 GitHub 仓库提交 Issue

## 项目结构

```
claude-code-marketplaces/
├── README.md                    # 插件市场介绍
├── LICENSE.txt                  # Apache 2.0 许可证
├── THIRD_PARTY_NOTICES.md       # 第三方组件声明
├── DEPLOYMENT_GUIDE.md          # 本部署指导
└── skills/
    └── markdown-word-converter/
        ├── SKILL.md             # 技能定义
        ├── scripts/
        │   ├── convert.py       # 主转换脚本
        │   └── install_dependencies.py # 依赖检查
        ├── assets/
        │   └── template.docx    # Word 模板
        ├── references/
        │   └── usage_examples.md # 使用示例
        ├── test_document.md     # 测试文档
        └── test_document.docx   # 转换结果示例
```

## 成功指标

部署成功的标志：
- ✅ 插件市场可以成功添加到 Claude Code
- ✅ 技能可以正常安装
- ✅ 依赖检查功能正常工作
- ✅ Markdown 转 Word 转换功能正常
- ✅ Mermaid 图表可以正确处理（如果 mmdc 已安装）

## 后续维护

1. **更新技能**：直接推送到 GitHub 仓库
2. **添加新技能**：在 `skills/` 目录下添加新技能
3. **文档更新**：保持 README.md 和相关文档的更新

---

**恭喜！您的 Markdown 转 Word 技能现在已经成功部署并可以在 Claude Code 中使用了。**