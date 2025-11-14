# Novel Project 技能故障排除指南

本文档提供了 novel-project 技能常见问题的解决方案和故障排除方法。

## 快速诊断

### 环境检查
在使用技能前，请确保以下环境要求已满足：

```bash
# 检查 Python 版本
python --version

# 检查文件权限
ls -la

# 检查磁盘空间
df -h
```

### 基本故障排查步骤
1. 检查错误信息
2. 查看日志输出
3. 验证输入参数
4. 检查文件权限
5. 确认环境配置

## 安装和配置问题

### 问题1：Python 版本不兼容
**错误信息**: `SyntaxError: invalid syntax` 或 `需要 Python 3.6 或更高版本`

**原因分析**:
- 使用了过旧版本的 Python
- 脚本使用了新版本的语法特性

**解决方案**:
```bash
# 安装或升级 Python
# Windows
# 从 python.org 下载最新版本

# macOS
brew install python3

# Linux Ubuntu/Debian
sudo apt update
sudo apt install python3 python3-pip

# 验证安装
python3 --version
```

**预防措施**:
- 在脚本开头添加版本检查
- 使用虚拟环境隔离依赖

### 问题2：模块导入失败
**错误信息**: `ModuleNotFoundError: No module named 'xxx'`

**常见场景**:
```bash
错误: ImportError: cannot import name 'Path' from 'pathlib'
错误: ModuleNotFoundError: No module named 'argparse'
```

**解决方案**:
```bash
# 安装缺失的模块
pip install module_name

# 检查模块是否在标准库中
python3 -c "import pathlib; print('pathlib available')"

# 使用完整的 Python 路径
/usr/bin/python3 script_name.py
```

**系统特定解决方案**:
```bash
# Windows
py -3 script_name.py

# macOS/Linux
python3 script_name.py

# 或者使用 shebang
#!/usr/bin/env python3
```

## 项目创建问题

### 问题3：目录权限不足
**错误信息**: `PermissionError: [Errno 13] Permission denied`

**原因分析**:
- 当前用户没有写入权限
- 目录被其他程序占用
- 磁盘空间不足

**解决方案**:
```bash
# 检查目录权限
ls -la /path/to/directory

# 修改权限（谨慎使用）
chmod 755 /path/to/directory

# 切换到有权限的目录
cd ~/Documents
python create_novel.py "小说名称"

# 检查磁盘空间
df -h
```

**Windows 解决方案**:
```cmd
# 以管理员身份运行命令提示符
# 或者切换到用户目录
cd %USERPROFILE%\Documents
python create_novel.py "小说名称"
```

### 问题4：项目名称冲突
**错误信息**: `项目目录已存在`

**场景示例**:
```
错误: 项目目录 '三国演义_project' 已存在。
请选择不同的小说名称或删除现有目录。
```

**解决方案**:
```bash
# 方案1：使用不同的名称
python create_novel.py "三国演义新编"

# 方案2：删除现有项目
rm -rf 三国演义_project

# 方案3：重命名现有项目
mv 三国演义_project 三国演义_backup

# 方案4：添加时间戳
python create_novel.py "三国演义_$(date +%Y%m%d)"
```

### 问题5：无效字符输入
**错误信息**: `小说名称包含无效字符`

**无效字符**:
```
< > : " | ? * / \
```

**解决方案**:
```bash
# 使用有效的字符
python create_novel.py "三国演义"        # ✅ 正确
python create_novel.py "三国演义：新章"  # ❌ 错误，包含冒号
python create_novel.py "三国演义_新章"   # ✅ 正确，使用下划线
```

**名称清理函数**:
```python
import re

def clean_novel_name(name):
    # 移除或替换无效字符
    cleaned = re.sub(r'[<>:"|?*\\/]', '_', name)
    return cleaned.strip()
```

## 章节管理问题

### 问题6：章节文件找不到
**错误信息**: `未找到章节 X` 或 `章节文件不存在`

**原因分析**:
- 章节文件被误删
- 文件名格式不正确
- 项目目录结构错误

**解决方案**:
```bash
# 检查章节目录
ls -la chapters/

# 查找特定章节
find . -name "*01*"

# 检查文件命名格式
ls chapters/*.md

# 重新创建丢失的章节
python chapter_manager.py create 1 "新章节标题"
```

### 问题7：章节编号混乱
**错误现象**:
- 章节编号不连续
- 文件名格式不一致
- 章节顺序错误

**解决方案**:
```bash
# 列出所有章节
python chapter_manager.py list

# 重新编号章节
python chapter_manager.py renumber

# 手动重命名章节
mv chapters/01-第一章.md chapters/01-乱世开始.md
mv chapters/02-第二章.md chapters/02-英雄出世.md
```

**重新编号脚本**:
```python
import re
from pathlib import Path

def renumber_chapters(chapters_dir):
    chapters = []
    pattern = re.compile(r'^(\d+)-(.*?)(?:\.md)?$')

    # 收集所有章节文件
    for file_path in sorted(chapters_dir.glob("*.md")):
        match = pattern.match(file_path.stem)
        if match:
            chapters.append({
                'path': file_path,
                'title': match.group(2),
                'content': file_path.read_text(encoding='utf-8')
            })

    # 重新编号
    for i, chapter in enumerate(chapters, 1):
        new_name = f"{i:02d}-{chapter['title']}.md"
        new_path = chapters_dir / new_name
        new_path.write_text(chapter['content'], encoding='utf-8')

        # 删除原文件
        if new_path != chapter['path']:
            chapter['path'].unlink()

    print(f"重新编号完成，共处理 {len(chapters)} 个章节")

# 使用
renumber_chapters(Path("chapters"))
```

## 文件内容问题

### 问题8：文件编码错误
**错误信息**: `UnicodeDecodeError: 'utf-8' codec can't decode`

**原因分析**:
- 文件使用了非 UTF-8 编码
- 文件损坏
- 二进制文件混入

**解决方案**:
```bash
# 检查文件编码
file -bi filename.md
chardet filename.md

# 转换编码
iconv -f gbk -t utf-8 source.txt > target.md

# 使用 Python 处理编码问题
with open('filename.md', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()
```

**编码检测函数**:
```python
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read(1024)  # 读取前1024字节
        result = chardet.detect(raw_data)
        return result['encoding']

# 使用
encoding = detect_encoding('filename.md')
with open('filename.md', 'r', encoding=encoding) as f:
    content = f.read()
```

### 问题9：模板文件损坏
**错误现象**:
- 生成的文件内容不完整
- 格式错乱
- 包含乱码

**解决方案**:
```bash
# 重新生成模板
rm -rf assets/templates/
python -c "
import sys
sys.path.append('scripts')
from create_novel import NovelProjectCreator
creator = NovelProjectCreator()
creator._create_template_files()
"

# 从备份恢复
cp -r backup/templates/ assets/templates/

# 手动修复模板文件
```

## 性能问题

### 问题10：大文件处理缓慢
**现象**:
- 项目包含大量章节时响应缓慢
- 文件打开和保存速度慢

**解决方案**:
```python
# 使用流式处理大文件
def process_large_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        for line_num, line in enumerate(f, 1):
            if line_num % 1000 == 0:
                print(f"已处理 {line_num} 行")
            # 处理每一行
            process_line(line)

# 使用生成器减少内存占用
def read_chapter_content(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        yield from f
```

### 问题11：内存占用过高
**原因分析**:
- 同时打开太多文件
- 未及时释放资源
- 内存泄漏

**解决方案**:
```python
# 使用上下文管理器
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 及时关闭文件
f = open(file_path, 'r', encoding='utf-8')
content = f.read()
f.close()

# 使用生成器减少内存占用
def get_all_chapters():
    for chapter_file in chapters_dir.glob("*.md"):
        yield chapter_file.name, chapter_file.read_text(encoding='utf-8')
```

## 系统兼容性问题

### 问题12：Windows 路径问题
**错误现象**:
- 路径分隔符错误
- 文件路径包含中文乱码
- 长路径名问题

**解决方案**:
```python
from pathlib import Path
import os

# 使用 Path 处理跨平台路径
def create_project_path(novel_name):
    project_name = f"{novel_name}_project"
    return Path.cwd() / project_name

# 处理中文路径
def handle_chinese_path(path):
    if os.name == 'nt':  # Windows
        return str(path).encode('utf-8').decode('utf-8')
    return path

# 处理长路径
def enable_long_paths():
    if os.name == 'nt':  # Windows
        import ctypes
        ctypes.windll.kernel32.SetFileAttributesW('\\\\?\\', 2)
```

### 问题13：MacOS 权限问题
**错误信息**: `Operation not permitted` 或 `无法创建目录`

**解决方案**:
```bash
# 检查安全设置
sudo spctl --status

# 给予终端完整磁盘访问权限
# 系统偏好设置 > 安全性与隐私 > 隐私 > 完整磁盘访问

# 使用用户目录
cd ~/Documents
python create_novel.py "小说名称"

# 检查文件权限
ls -la ~/Documents
```

## 调试技巧

### 启用详细日志
```python
import logging

# 配置日志
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='novel_project.log'
)

logger = logging.getLogger(__name__)

# 在关键位置添加日志
def create_project(novel_name):
    logger.info(f"开始创建项目: {novel_name}")

    try:
        project_path = generate_project_path(novel_name)
        logger.debug(f"项目路径: {project_path}")

        # 创建目录
        create_directory_structure(project_path)
        logger.info("目录结构创建成功")

    except Exception as e:
        logger.error(f"创建项目失败: {e}")
        raise
```

### 使用调试模式
```bash
# 启用详细输出
python create_novel.py "测试小说" --verbose

# 检查脚本语法
python -m py_compile scripts/create_novel.py

# 运行测试
python -m pytest tests/ -v
```

### 环境诊断脚本
```python
#!/usr/bin/env python3
import sys
import os
from pathlib import Path

def diagnose_environment():
    print("=== 环境诊断 ===")
    print(f"Python 版本: {sys.version}")
    print(f"当前工作目录: {os.getcwd()}")
    print(f"操作系统: {os.name}")

    # 检查必要模块
    required_modules = ['pathlib', 'json', 'argparse', 'datetime', 're']
    for module in required_modules:
        try:
            __import__(module)
            print(f"✓ {module}: 可用")
        except ImportError:
            print(f"✗ {module}: 不可用")

    # 检查文件权限
    test_file = Path.cwd() / "test_permission.txt"
    try:
        test_file.write_text("test", encoding='utf-8')
        test_file.unlink()
        print("✓ 文件写入权限: 正常")
    except Exception as e:
        print(f"✗ 文件写入权限: {e}")

    # 检查磁盘空间
    import shutil
    total, used, free = shutil.disk_usage(Path.cwd())
    print(f"磁盘空间: {free // (1024**3)} GB 可用")

if __name__ == "__main__":
    diagnose_environment()
```

## 寻求帮助

### 收集诊断信息
在报告问题时，请提供以下信息：

1. **系统信息**:
   ```bash
   python --version
   uname -a  # Linux/Mac
   ver       # Windows
   ```

2. **错误信息**: 完整的错误堆栈信息

3. **重现步骤**: 详细的操作步骤

4. **环境设置**: 项目目录、文件权限等

### 联系方式
- 查看项目文档获取最新联系方式
- 提交 Issue 到项目仓库
- 联系技术支持团队

---

**提示**: 大多数问题都可以通过检查环境配置和文件权限来解决。如果问题持续存在，请尝试在干净的环境中重新安装技能。