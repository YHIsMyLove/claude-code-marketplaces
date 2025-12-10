#!/usr/bin/env bun
/**
 * Novel Project Structure Generator
 * 用于生成规范化小说工程文件夹结构
 */

import { existsSync, mkdirSync, writeFileSync, readdirSync } from 'fs';
import { join } from 'path';

interface NovelProjectConfig {
  novelName: string;
  chapterCount?: number;
  author?: string;
  genre?: string;
  description?: string;
  targetDir?: string;
}

const DEFAULT_CHAPTER_COUNT = 10;
const PROGRESS_TEMPLATE = `# {novelName} - 写作进度

## 基本信息
- **小说名称**: {novelName}
- **作者**: {author}
- **类型**: {genre}
- **总章节数**: {totalChapters}
- **创建时间**: {createdAt}

## 章节进度

### 已完成章节
- [ ] 第1章 - 尚未开始

### 进行中章节
暂无

### 待写作章节
- 第2章
- 第3章
- ...（后续章节）

## 统计信息
- **总字数**: 0
- **已完成章节**: 0/{totalChapters}
- **完成进度**: 0%

## 笔记与想法
- 在这里记录创作灵感、角色设定、情节构思等
-
-

## 更新日志
- {createdAt}: 创建项目
`;

const CHAPTER_TEMPLATE = `# 第{chapterNumber}章

## 本章概要
（在此处填写本章的主要情节概要）

## 场景设置
- **时间**:
- **地点**:
- **出场人物**:

## 正文

---

## 本章笔记
- 关键情节:
- 人物发展:
- 伏笔线索:
- 字数统计: 0
`;

/**
 * 安全地创建目录
 */
function ensureDir(dirPath: string): void {
  if (!existsSync(dirPath)) {
    mkdirSync(dirPath, { recursive: true });
    console.log(`✅ 创建目录: ${dirPath}`);
  }
}

/**
 * 格式化章节名称
 */
function formatChapterName(chapterNumber: number): string {
  return `第${chapterNumber}章`;
}

/**
 * 生成进度文件内容
 */
function generateProgressContent(config: NovelProjectConfig): string {
  const now = new Date().toLocaleString('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });

  return PROGRESS_TEMPLATE
    .replace(/{novelName}/g, config.novelName)
    .replace(/{author}/g, config.author || '未知')
    .replace(/{genre}/g, config.genre || '未分类')
    .replace(/{totalChapters}/g, String(config.chapterCount || DEFAULT_CHAPTER_COUNT))
    .replace(/{createdAt}/g, now);
}

/**
 * 生成章节文件内容
 */
function generateChapterContent(chapterNumber: number): string {
  return CHAPTER_TEMPLATE.replace(/{chapterNumber}/g, String(chapterNumber));
}

/**
 * 创建数据库子目录结构
 */
function createDatabaseStructure(dbPath: string): void {
  const subdirs = [
    '人物设定',
    '世界观设定',
    '情节大纲',
    '素材收集',
    '参考资料'
  ];

  subdirs.forEach(dir => {
    ensureDir(join(dbPath, dir));
  });

  // 创建人物设定模板文件
  const characterTemplate = `# 人物设定模板

## 基本信息
- **姓名**:
- **年龄**:
- **性别**:
- **职业**:
- **外貌特征**:
- **性格特点**:

## 背景故事
（描述人物的成长经历、家庭背景等）

## 人物关系
- 与主角的关系:
- 与其他角色的关系:

## 角色发展
- 故事开始时的状态:
- 故事结束时的状态:
- 关键转折点:

## 经典台词
-
-
`;

  writeFileSync(
    join(dbPath, '人物设定', '人物设定模板.md'),
    characterTemplate,
    'utf8'
  );
  console.log(`✅ 创建人物设定模板`);
}

/**
 * 生成小说工程文件夹结构
 */
function generateNovelProject(config: NovelProjectConfig): void {
  const targetDir = config.targetDir || process.cwd();
  const projectPath = join(targetDir, config.novelName);

  console.log(`🚀 开始生成小说工程: ${config.novelName}`);
  console.log(`📍 目标路径: ${projectPath}`);

  // 创建主项目目录
  ensureDir(projectPath);

  // 创建主要子目录
  const dbPath = join(projectPath, '数据库');
  const chaptersPath = join(projectPath, '章节');

  ensureDir(dbPath);
  ensureDir(chaptersPath);

  // 创建数据库子结构
  createDatabaseStructure(dbPath);

  // 创建章节目录和文件
  const chapterCount = config.chapterCount || DEFAULT_CHAPTER_COUNT;
  for (let i = 1; i <= chapterCount; i++) {
    const chapterPath = join(chaptersPath, formatChapterName(i));
    ensureDir(chapterPath);

    const chapterFile = join(chapterPath, `${formatChapterName(i)}.md`);
    writeFileSync(chapterFile, generateChapterContent(i), 'utf8');
    console.log(`✅ 创建章节文件: ${formatChapterName(i)}.md`);
  }

  // 创建进度文件
  const progressFile = join(projectPath, '进度.md');
  writeFileSync(progressFile, generateProgressContent(config), 'utf8');
  console.log(`✅ 创建进度文件: 进度.md`);

  // 创建项目说明文件
  const readmeFile = join(projectPath, 'README.md');
  const readmeContent = `# ${config.novelName}

## 项目说明
这是一个使用小说工程生成器创建的小说写作项目。

## 目录结构
\`\`\`
${config.novelName}/
├── 数据库/           # 存放设定、素材等参考资料
│   ├── 人物设定/      # 人物角色设定
│   ├── 世界观设定/    # 世界观和背景设定
│   ├── 情节大纲/      # 故事情节大纲
│   ├── 素材收集/      # 写作素材收集
│   └── 参考资料/      # 参考资料和文献
├── 章节/             # 章节内容存放
│   ├── 第1章/        # 各章节独立目录
│   │   └── 第1章.md   # 章节正文文件
│   └── ...
├── 进度.md           # 写作进度跟踪
└── README.md         # 项目说明文件
\`\`\`

## 工具推荐
- 使用 \`bun\` 作为运行时环境
- 支持Markdown编辑器进行写作
- 建议使用版本控制工具（如Git）管理草稿

## 开始写作
1. 查看 \`进度.md\` 了解项目概况
2. 编辑 \`数据库/\` 中的设定文件
3. 开始在 \`章节/\` 中进行写作
4. 定期更新 \`进度.md\` 跟踪写作进展

祝创作愉快！📚✍️
`;
  writeFileSync(readmeFile, readmeContent, 'utf8');
  console.log(`✅ 创建项目说明: README.md`);

  console.log(`\n🎉 小说工程 "${config.novelName}" 创建完成！`);
  console.log(`📁 项目路径: ${projectPath}`);
  console.log(`📖 开始写作: cd "${projectPath}"`);
}

/**
 * 命令行参数解析
 */
function parseArguments(): NovelProjectConfig | null {
  const args = process.argv.slice(2);

  // 检查帮助参数
  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    console.log(`
📚 小说工程生成器

用法:
  bun run scripts/generate-novel-project.ts <小说名称> [选项]

参数:
  小说名称                    必需，小说项目的名称

选项:
  --chapters <数量>          章节数量 (默认: ${DEFAULT_CHAPTER_COUNT})
  --author <作者>            作者名称
  --genre <类型>             小说类型
  --description <描述>       小说描述
  --dir <目录>               目标目录 (默认: 当前目录)
  --help, -h                 显示此帮助信息

示例:
  bun run scripts/generate-novel-project.ts "我的小说"
  bun run scripts/generate-novel-project.ts "奇幻冒险" --chapters 20 --author "张三" --genre "奇幻"
`);
    return null;
  }

  const config: NovelProjectConfig = {
    novelName: args[0],
  };

  // 解析可选参数
  for (let i = 1; i < args.length; i += 2) {
    const flag = args[i];
    const value = args[i + 1];

    if (!value || !flag.startsWith('--')) {
      console.error(`❌ 无效的参数格式: ${flag}`);
      return null;
    }

    switch (flag) {
      case '--chapters':
        config.chapterCount = parseInt(value, 10);
        if (isNaN(config.chapterCount) || config.chapterCount <= 0) {
          console.error(`❌ 无效的章节数量: ${value}`);
          return null;
        }
        break;
      case '--author':
        config.author = value;
        break;
      case '--genre':
        config.genre = value;
        break;
      case '--description':
        config.description = value;
        break;
      case '--dir':
        config.targetDir = value;
        break;
      default:
        console.error(`❌ 未知选项: ${flag}`);
        return null;
    }
  }

  return config;
}

/**
 * 主函数
 */
function main(): void {
  const config = parseArguments();

  if (!config) {
    process.exit(1);
  }

  try {
    generateNovelProject(config);
  } catch (error) {
    console.error(`❌ 生成项目时出错: ${error}`);
    process.exit(1);
  }
}

// 如果直接运行此脚本，执行主函数
if (import.meta.main) {
  main();
}

export { generateNovelProject, NovelProjectConfig };