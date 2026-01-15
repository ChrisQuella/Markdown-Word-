// 引用 docx 库
const { Document, Paragraph, TextRun, Table, TableRow, TableCell, Packer, WidthType, BorderStyle, LevelFormat, AlignmentType } = window.docx;

// 样式配置
let styleConfig = {
    bodyFont: '宋体',
    bodySize: 12,
    headingFont: '黑体',
    lineSpacing: 1.5,
    firstIndent: 2,
    pageMargin: 'normal'
};

// 预设模板
const templates = {
    default: {
        bodyFont: '宋体',
        bodySize: 12,
        headingFont: '黑体',
        lineSpacing: 1.5,
        firstIndent: 2,
        pageMargin: 'normal'
    },
    academic: {
        bodyFont: '宋体',
        bodySize: 12,
        headingFont: '黑体',
        lineSpacing: 2,
        firstIndent: 2,
        pageMargin: 'normal'
    },
    business: {
        bodyFont: '微软雅黑',
        bodySize: 10.5,
        headingFont: '微软雅黑',
        lineSpacing: 1.5,
        firstIndent: 0,
        pageMargin: 'normal'
    },
    minimal: {
        bodyFont: 'Arial',
        bodySize: 11,
        headingFont: 'Arial',
        lineSpacing: 1.5,
        firstIndent: 0,
        pageMargin: 'narrow'
    }
};

// 示例 Markdown
const sampleMarkdown = `# Markdown 转 Word 工具使用指南

## 简介

这是一个功能强大的 **Markdown 转 Word** 在线工具，支持多种格式转换。

## 主要功能

### 1. 格式支持

- 标题（H1-H6）
- **粗体** 和 *斜体*
- ~~删除线~~
- 有序和无序列表
- 代码块和行内代码
- 表格
- 引用块
- 分割线

### 2. 代码示例

\`\`\`javascript
function hello() {
    console.log("Hello, World!");
    return true;
}
\`\`\`

### 3. 表格示例

| 功能 | 支持状态 | 备注 |
|------|----------|------|
| 标题 | ✅ | H1-H6 |
| 列表 | ✅ | 有序/无序 |
| 代码 | ✅ | 高亮显示 |
| 表格 | ✅ | 完整支持 |

### 4. 引用示例

> 这是一段引用文字。
> 可以包含多行内容。

## 使用步骤

1. 在左侧编辑器输入 Markdown 内容
2. 右侧实时预览效果
3. 点击"样式设置"自定义格式
4. 点击"下载 Word"获取文档

---

**感谢使用！** 🎉
`;

// DOM 元素
const markdownInput = document.getElementById('markdownInput');
const preview = document.getElementById('preview');
const charCount = document.getElementById('charCount');
const fileInput = document.getElementById('fileInput');
const clearBtn = document.getElementById('clearBtn');
const sampleBtn = document.getElementById('sampleBtn');
const styleBtn = document.getElementById('styleBtn');
const downloadBtn = document.getElementById('downloadBtn');
const styleModal = document.getElementById('styleModal');
const closeModal = document.getElementById('closeModal');
const resetStyle = document.getElementById('resetStyle');
const applyStyle = document.getElementById('applyStyle');

// 配置 marked
marked.setOptions({
    highlight: function(code, lang) {
        if (lang && hljs.getLanguage(lang)) {
            return hljs.highlight(code, { language: lang }).value;
        }
        return hljs.highlightAuto(code).value;
    },
    breaks: true,
    gfm: true
});

// 实时预览
function updatePreview() {
    const markdown = markdownInput.value;
    preview.innerHTML = marked.parse(markdown);
    charCount.textContent = `${markdown.length} 字符`;
}

// 事件监听
markdownInput.addEventListener('input', updatePreview);

// 文件上传
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
            markdownInput.value = event.target.result;
            updatePreview();
        };
        reader.readAsText(file);
    }
});

// 清空
clearBtn.addEventListener('click', () => {
    markdownInput.value = '';
    updatePreview();
});

// 示例
sampleBtn.addEventListener('click', () => {
    markdownInput.value = sampleMarkdown;
    updatePreview();
});

// 样式弹窗
styleBtn.addEventListener('click', () => {
    styleModal.classList.add('active');
    loadStyleToForm();
});

closeModal.addEventListener('click', () => {
    styleModal.classList.remove('active');
});

styleModal.addEventListener('click', (e) => {
    if (e.target === styleModal) {
        styleModal.classList.remove('active');
    }
});

// 模板选择
document.querySelectorAll('.template-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.template-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const template = templates[btn.dataset.template];
        Object.assign(styleConfig, template);
        loadStyleToForm();
    });
});

// 加载样式到表单
function loadStyleToForm() {
    document.getElementById('bodyFont').value = styleConfig.bodyFont;
    document.getElementById('bodySize').value = styleConfig.bodySize;
    document.getElementById('headingFont').value = styleConfig.headingFont;
    document.getElementById('lineSpacing').value = styleConfig.lineSpacing;
    document.getElementById('firstIndent').value = styleConfig.firstIndent;
    document.getElementById('pageMargin').value = styleConfig.pageMargin;
}

// 从表单读取样式
function readStyleFromForm() {
    styleConfig.bodyFont = document.getElementById('bodyFont').value;
    styleConfig.bodySize = parseFloat(document.getElementById('bodySize').value);
    styleConfig.headingFont = document.getElementById('headingFont').value;
    styleConfig.lineSpacing = parseFloat(document.getElementById('lineSpacing').value);
    styleConfig.firstIndent = parseInt(document.getElementById('firstIndent').value);
    styleConfig.pageMargin = document.getElementById('pageMargin').value;
}

// 重置样式
resetStyle.addEventListener('click', () => {
    Object.assign(styleConfig, templates.default);
    loadStyleToForm();
    document.querySelectorAll('.template-btn').forEach(b => b.classList.remove('active'));
    document.querySelector('[data-template="default"]').classList.add('active');
});

// 应用样式
applyStyle.addEventListener('click', () => {
    readStyleFromForm();
    styleModal.classList.remove('active');
});

// 解析 Markdown 为结构化数据
function parseMarkdown(markdown) {
    const lines = markdown.split('\n');
    const elements = [];
    let inCodeBlock = false;
    let codeContent = '';
    let codeLang = '';
    let inTable = false;
    let tableRows = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // 代码块
        if (line.startsWith('```')) {
            if (!inCodeBlock) {
                inCodeBlock = true;
                codeLang = line.slice(3).trim();
                codeContent = '';
            } else {
                elements.push({ type: 'code', content: codeContent.trim(), lang: codeLang });
                inCodeBlock = false;
            }
            continue;
        }

        if (inCodeBlock) {
            codeContent += line + '\n';
            continue;
        }

        // 表格
        if (line.includes('|') && line.trim().startsWith('|')) {
            if (!inTable) {
                inTable = true;
                tableRows = [];
            }
            if (!line.match(/^\|[\s-:|]+\|$/)) {
                tableRows.push(line.split('|').filter(cell => cell.trim()).map(cell => cell.trim()));
            }
            continue;
        } else if (inTable) {
            elements.push({ type: 'table', rows: tableRows });
            inTable = false;
            tableRows = [];
        }

        // 空行
        if (line.trim() === '') {
            continue;
        }

        // 标题
        const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) {
            elements.push({ type: 'heading', level: headingMatch[1].length, content: headingMatch[2] });
            continue;
        }

        // 分割线
        if (line.match(/^[-*_]{3,}$/)) {
            elements.push({ type: 'hr' });
            continue;
        }

        // 引用
        if (line.startsWith('>')) {
            elements.push({ type: 'quote', content: line.replace(/^>\s*/, '') });
            continue;
        }

        // 无序列表
        if (line.match(/^[\s]*[-*+]\s+/)) {
            const indent = line.match(/^(\s*)/)[1].length;
            const content = line.replace(/^[\s]*[-*+]\s+/, '');
            elements.push({ type: 'bullet', content, indent: Math.floor(indent / 2) });
            continue;
        }

        // 有序列表
        if (line.match(/^[\s]*\d+\.\s+/)) {
            const indent = line.match(/^(\s*)/)[1].length;
            const content = line.replace(/^[\s]*\d+\.\s+/, '');
            elements.push({ type: 'number', content, indent: Math.floor(indent / 2) });
            continue;
        }

        // 普通段落
        elements.push({ type: 'paragraph', content: line });
    }

    // 处理未结束的表格
    if (inTable && tableRows.length > 0) {
        elements.push({ type: 'table', rows: tableRows });
    }

    return elements;
}

// 解析行内格式 - 返回配置对象数组
function parseInlineFormatting(text, baseFont, baseSize) {
    const runs = [];
    let remaining = text;

    while (remaining.length > 0) {
        // 粗体
        let match = remaining.match(/\*\*(.+?)\*\*/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], bold: true, font: baseFont, size: baseSize });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        // 斜体
        match = remaining.match(/\*(.+?)\*/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], italics: true, font: baseFont, size: baseSize });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        // 删除线
        match = remaining.match(/~~(.+?)~~/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], strike: true, font: baseFont, size: baseSize });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        // 行内代码
        match = remaining.match(/`(.+?)`/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], font: 'Consolas', size: baseSize });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        // 查找下一个特殊字符
        const nextSpecial = remaining.search(/\*\*|\*|~~|`/);
        if (nextSpecial > 0) {
            runs.push({ text: remaining.slice(0, nextSpecial), font: baseFont, size: baseSize });
            remaining = remaining.slice(nextSpecial);
        } else if (nextSpecial === -1) {
            runs.push({ text: remaining, font: baseFont, size: baseSize });
            break;
        } else {
            runs.push({ text: remaining[0], font: baseFont, size: baseSize });
            remaining = remaining.slice(1);
        }
    }

    return runs.length > 0 ? runs : [{ text, font: baseFont, size: baseSize }];
}


// 生成 Word 文档
async function generateWord() {
    try {
        const markdown = markdownInput.value;
        if (!markdown.trim()) {
            alert('请先输入 Markdown 内容');
            return;
        }

        const elements = parseMarkdown(markdown);
        const children = [];

        // 页边距配置
        const margins = {
            normal: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
            narrow: { top: 720, right: 720, bottom: 720, left: 720 },
            wide: { top: 1800, right: 1800, bottom: 1800, left: 1800 }
        };

        // 标题大小映射 (half-points)
        const headingSizes = {
            1: 64,
            2: 52,
            3: 44,
            4: 36,
            5: 32,
            6: 28
        };

        // 正文大小 (half-points)
        const bodySize = styleConfig.bodySize * 2;

        // 行间距转换 (twips)
        const lineSpacingValue = Math.round(styleConfig.lineSpacing * 240);

        for (const el of elements) {
            switch (el.type) {
                case 'heading':
                    children.push(new Paragraph({
                        children: [new TextRun({
                            text: el.content,
                            bold: true,
                            size: headingSizes[el.level],
                            font: styleConfig.headingFont
                        })],
                        spacing: { before: 240, after: 120, line: lineSpacingValue }
                    }));
                    break;

                case 'paragraph':
                    const pRuns = parseInlineFormatting(el.content, styleConfig.bodyFont, bodySize);
                    children.push(new Paragraph({
                        children: pRuns.map(r => new TextRun(r)),
                        spacing: { after: 120, line: lineSpacingValue },
                        indent: styleConfig.firstIndent > 0 ? { firstLine: styleConfig.firstIndent * 240 } : undefined
                    }));
                    break;

                case 'bullet':
                    const bRuns = parseInlineFormatting(el.content, styleConfig.bodyFont, bodySize);
                    children.push(new Paragraph({
                        children: bRuns.map(r => new TextRun(r)),
                        bullet: { level: el.indent },
                        spacing: { after: 60, line: lineSpacingValue }
                    }));
                    break;

                case 'number':
                    const nRuns = parseInlineFormatting(el.content, styleConfig.bodyFont, bodySize);
                    children.push(new Paragraph({
                        children: nRuns.map(r => new TextRun(r)),
                        numbering: { reference: 'default-numbering', level: el.indent },
                        spacing: { after: 60, line: lineSpacingValue }
                    }));
                    break;

                case 'quote':
                    children.push(new Paragraph({
                        children: [new TextRun({
                            text: el.content,
                            italics: true,
                            color: '666666',
                            size: bodySize,
                            font: styleConfig.bodyFont
                        })],
                        indent: { left: 720 },
                        border: {
                            left: { style: BorderStyle.SINGLE, size: 24, color: '667eea' }
                        },
                        spacing: { after: 120, line: lineSpacingValue }
                    }));
                    break;

                case 'code':
                    const codeLines = el.content.split('\n');
                    for (const codeLine of codeLines) {
                        children.push(new Paragraph({
                            children: [new TextRun({
                                text: codeLine || ' ',
                                font: 'Consolas',
                                size: 20
                            })],
                            shading: { fill: 'f4f4f4' },
                            spacing: { after: 0, line: 240 }
                        }));
                    }
                    children.push(new Paragraph({ children: [] }));
                    break;

                case 'table':
                    if (el.rows.length > 0) {
                        const tableRows = el.rows.map((row, rowIndex) => {
                            return new TableRow({
                                children: row.map(cell => {
                                    return new TableCell({
                                        children: [new Paragraph({
                                            children: [new TextRun({
                                                text: cell,
                                                bold: rowIndex === 0,
                                                size: bodySize,
                                                font: styleConfig.bodyFont
                                            })]
                                        })],
                                        shading: rowIndex === 0 ? { fill: 'f8f9fa' } : undefined
                                    });
                                })
                            });
                        });

                        children.push(new Table({
                            rows: tableRows,
                            width: { size: 100, type: WidthType.PERCENTAGE }
                        }));
                        children.push(new Paragraph({ children: [] }));
                    }
                    break;

                case 'hr':
                    children.push(new Paragraph({
                        children: [],
                        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'cccccc' } },
                        spacing: { before: 240, after: 240 }
                    }));
                    break;
            }
        }

        // 创建文档
        const doc = new Document({
            numbering: {
                config: [{
                    reference: 'default-numbering',
                    levels: [
                        { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START },
                        { level: 1, format: LevelFormat.DECIMAL, text: '%1.%2.', alignment: AlignmentType.START },
                        { level: 2, format: LevelFormat.DECIMAL, text: '%1.%2.%3.', alignment: AlignmentType.START }
                    ]
                }]
            },
            sections: [{
                properties: {
                    page: {
                        margin: margins[styleConfig.pageMargin]
                    }
                },
                children: children
            }]
        });

        // 生成并下载
        const blob = await Packer.toBlob(doc);
        saveAs(blob, 'document.docx');
    } catch (error) {
        console.error('生成文档失败:', error);
        alert('生成文档失败: ' + error.message);
    }
}

// 下载按钮
downloadBtn.addEventListener('click', generateWord);

// 初始化
updatePreview();
