# 📝 Markdown 转 Word 在线工具

一款功能丰富的 Markdown 转 Word 文档工具，针对不同用户群体提供三个专业版本。

![preview](https://img.shields.io/badge/Platform-Windows-blue) ![license](https://img.shields.io/badge/License-MIT-green)

## ✨ 三大版本

### 🎓 学术版
专为论文写作、学术报告设计

- LaTeX 数学公式（行内 `$E=mc^2$`，块级 `$$...$$`）
- 自动生成目录
- 行号显示
- 学术论文模板（学位论文、期刊论文、研究报告）
- 参考文献格式支持

### 💻 开发者版
专为技术文档、代码笔记设计

- 深色主题界面
- 代码块语法高亮
- 自动保存草稿到本地
- 快捷键支持（Ctrl+B/I/S/K 等）
- 光标位置显示（Ln/Col）
- 技术文档模板

### � 速办公版
专为工作报告、商务文档设计

- 可视化工具栏（点击即插入）
- 图片拖拽上传
- 撤销/重做功能
- 双格式导出（Word + PDF）
- 页眉页脚设置
- 阅读时间估算

## 📦 通用功能

| 功能 | 支持 |
|------|------|
| 标题 (H1-H6) | ✅ |
| 粗体 / 斜体 / 删除线 | ✅ |
| 有序 / 无序列表 | ✅ |
| 代码块（语法高亮） | ✅ |
| 表格 | ✅ |
| 引用块 | ✅ |
| 分割线 | ✅ |
| 自定义文件名导出 | ✅ |
| 自定义样式模板 | ✅ |

## 🚀 快速开始

### 在线使用

直接用浏览器打开 `index.html`，选择对应版本即可使用（需联网加载 CDN 资源）

### 本地开发

```bash
# 启动本地服务
npx serve -p 3000

# 访问 http://localhost:3000
```

### 打包 exe

```bash
# 安装依赖
npm install

# 打包 Windows 便携版
npm run build

# 生成文件位于 dist/MD转Word工具 1.0.0.exe
```

## 🛠️ 技术栈

- **前端**: HTML + CSS + JavaScript
- **Markdown 解析**: [marked.js](https://marked.js.org/)
- **代码高亮**: [highlight.js](https://highlightjs.org/)
- **数学公式**: [KaTeX](https://katex.org/)
- **Word 生成**: [docx.js](https://docx.js.org/)
- **PDF 导出**: [html2pdf.js](https://ekoopmans.github.io/html2pdf.js/)
- **桌面打包**: [Electron](https://www.electronjs.org/)

## 📁 项目结构

```
├── index.html          # 首页（版本选择）
├── academic.html/js    # 学术版
├── developer.html/js   # 开发者版
├── office.html/js      # 办公版
├── shared.css          # 共享样式
├── main.js             # Electron 入口
├── package.json        # 项目配置
└── dist/               # 打包输出目录
```

## ⌨️ 快捷键（开发者版）

| 快捷键 | 功能 |
|--------|------|
| Ctrl + B | 粗体 |
| Ctrl + I | 斜体 |
| Ctrl + ` | 行内代码 |
| Ctrl + K | 链接 |
| Ctrl + Shift + K | 代码块 |
| Ctrl + S | 保存草稿 |
| Ctrl + H | 标题 |

## 📄 License

MIT
