# keynote-ppt-to-markdown

[中文](README.md) | [English](README_en.md)

一个将 Keynote/PPT 演示文稿转换为 Markdown 格式的命令行工具，支持导出演讲者注释和幻灯片图片。

## 功能特点

- 支持 Keynote 和 PowerPoint 演示文稿转换
- 保留演讲者注释
- 导出幻灯片图片
- 生成结构化的 Markdown 文档

## 使用方法

直接运行:

```bash
npx keynote-ppt-to-markdown <演示文稿路径> [输出目录]
```

或者，先安装再运行:

```bash
npm install -g keynote-ppt-to-markdown
ppt2md <演示文稿路径> [输出目录]
```

### 参数说明

- `<演示文稿路径>`: 必需，Keynote 或 PowerPoint 文件的路径
- `[输出目录]`: 可选，指定输出目录，默认为当前目录

### 示例

```bash
# 转换 Keynote 文件
npx keynote-ppt-to-markdown presentation.key

# 转换 PowerPoint 文件并指定输出目录
npx keynote-ppt-to-markdown presentation.pptx ./output
```

## 输出内容

转换后的输出包含：

- `presentation.md`: 包含幻灯片每页图片和演讲者注释的 Markdown 文件
- `images/`: 包含所有幻灯片图片的目录

## 开发

### 安装依赖

```bash
npm install
```

### 构建

```bash
npm run build
```

### 运行

```bash
npm run dev
npx keynote-ppt-to-markdown <演示文稿路径> [输出目录]
```

## 许可证

MIT

## 贡献

欢迎提交 Issue 和 Pull Request！

## 致谢

Cursor 帮我完成了几乎所有代码，感谢！
