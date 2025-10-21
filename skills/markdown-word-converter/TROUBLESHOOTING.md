# Mermaid CLI 故障排除指南

## 问题诊断

如果您遇到 "⚠️ 由于缺少 mermaid-cli，文档中的 Mermaid 图表未能转换为图片格式" 错误，即使您已经安装了 mmdc，本指南将帮助您诊断和解决问题。

## 快速诊断

使用新的测试功能来诊断 mmdc 的具体问题：

```bash
python scripts/convert.py --test-mmdc
```

这个命令会执行以下检查：
1. 验证 mmdc 命令是否在 PATH 中
2. 检查 mmdc 版本信息
3. 执行实际的功能测试，转换一个简单的图表

## 常见问题和解决方案

### 1. mmdc 已安装但检测不到

**症状**: `mmdc --version` 在终端中工作，但脚本报告找不到 mmdc

**解决方案**:
- Windows: 重启命令提示符或 PowerShell
- macOS/Linux: 重启终端或运行 `source ~/.bashrc` 或 `source ~/.zshrc`
- 确认 npm 全局安装路径在系统 PATH 中

### 2. Puppeteer/Chrome 问题

**症状**: mmdc 找到但转换失败，提到 Puppeteer 或 Chrome

**解决方案**:
```bash
# 重新安装 mermaid-cli（包含 puppeteer）
npm install -g @mermaid-js/mermaid-cli

# 或者单独安装 puppeteer
npm install -g puppeteer
```

### 3. 权限问题

**症状**: 拒绝访问或权限相关错误

**解决方案**:
- Windows: 以管理员身份运行命令提示符
- macOS/Linux: 使用 `sudo npm install -g @mermaid-js/mermaid-cli`

### 4. 网络问题

**症状**: 安装过程中网络相关错误

**解决方案**:
```bash
# 使用国内镜像（中国用户）
npm config set registry https://registry.npmmirror.com/
npm install -g @mermaid-js/mermaid-cli

# 或者使用淘宝镜像
npm config set registry https://registry.npm.taobao.org/
npm install -g @mermaid-js/mermaid-cli
```

## 验证安装

使用以下命令验证安装：

```bash
# 检查 mmdc 版本
mmdc --version

# 检查 npm 全局包
npm list -g @mermaid-js/mermaid-cli

# 运行功能测试
python scripts/convert.py --test-mmdc

# 检查所有依赖
python scripts/convert.py --check-deps
```

## 转换测试

验证 mmdc 可以正常转换后，测试完整的转换流程：

```bash
# 使用项目提供的测试文档
python scripts/convert.py test_document.md

# 或者转换自己的文档
python scripts/convert.py your_document.md
```

## 改进内容

本次更新包含以下改进：

1. **更可靠的依赖检测**
   - 使用 `subprocess.run` 直接执行 `mmdc --version`
   - 添加超时和详细错误报告
   - 提供具体的版本信息

2. **优化的 mmdc 调用**
   - 移除可能不兼容的 `-s` 参数
   - 增加转换超时时间
   - 改进错误处理和日志记录

3. **增强的错误报告**
   - 详细的错误信息和诊断
   - 针对常见问题的具体建议
   - 清晰的进度指示和状态反馈

4. **新增测试功能**
   - `--test-mmdc` 参数用于功能测试
   - 实际转换测试而非仅检查命令存在
   - 全面的故障排除指导

## 手动 Mermaid 转换

如果自动转换仍有问题，您可以手动转换 Mermaid 图表：

1. 从 Markdown 文件中提取 Mermaid 代码块
2. 保存为 `.mmd` 文件
3. 使用 mmdc 手动转换：
   ```bash
   mmdc -i diagram.mmd -o diagram.png -e png
   ```
4. 在 Markdown 中引用生成的 PNG 图片

## 获取帮助

如果问题仍然存在：

1. 查看 mmdc 官方文档：https://github.com/mermaid-js/mermaid-cli
2. 检查 Node.js 和 npm 版本是否过旧
3. 确认系统有足够的磁盘空间和内存
4. 在 GitHub 仓库提交 Issue 并提供详细的错误信息

---

**注意**: mmdc 需要下载 Chromium，首次使用可能需要较长时间。