// 学术版 - Markdown 转 Word
const { Document, Paragraph, TextRun, Table, TableRow, TableCell, Packer, WidthType, BorderStyle, LevelFormat, AlignmentType } = window.docx;

// 样式配置
let styleConfig = {
    bodyFont: '宋体',
    engFont: 'Times New Roman',
    bodySize: 12,
    headingFont: '黑体',
    lineSpacing: 2,
    firstIndent: 2,
    pageMargin: 'normal'
};

// DOM 元素
const markdownInput = document.getElementById('markdownInput');
const preview = document.getElementById('preview');
const charCount = document.getElementById('charCount');
const wordCount = document.getElementById('wordCount');
const lineNumbers = document.getElementById('lineNumbers');
const fileInput = document.getElementById('fileInput');
const tocBtn = document.getElementById('tocBtn');
const styleBtn = document.getElementById('styleBtn');
const downloadBtn = document.getElementById('downloadBtn');
const styleModal = document.getElementById('styleModal');
const closeModal = document.getElementById('closeModal');
const showToc = document.getElementById('showToc');

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

// 渲染 LaTeX 公式
function renderMath(html) {
    // 块级公式 $$...$$
    html = html.replace(/\$\$([^$]+)\$\$/g, (match, formula) => {
        try {
            return katex.renderToString(formula.trim(), { displayMode: true });
        } catch (e) {
            return `<span style="color:red">[公式错误: ${e.message}]</span>`;
        }
    });
    // 行内公式 $...$
    html = html.replace(/\$([^$\n]+)\$/g, (match, formula) => {
        try {
            return katex.renderToString(formula.trim(), { displayMode: false });
        } catch (e) {
            return `<span style="color:red">[公式错误]</span>`;
        }
    });
    return html;
}

// 生成目录
function generateTOC(markdown) {
    const headings = [];
    const lines = markdown.split('\n');
    lines.forEach(line => {
        const match = line.match(/^(#{2,4})\s+(.+)$/);
        if (match) {
            headings.push({
                level: match[1].length,
                text: match[2],
                id: match[2].toLowerCase().replace(/\s+/g, '-').replace(/[^\w\u4e00-\u9fa5-]/g, '')
            });
        }
    });
    
    if (headings.length === 0) return '';
    
    let toc = '<div class="toc"><div class="toc-title">📑 目录</div><ul class="toc-list">';
    headings.forEach(h => {
        toc += `<li class="toc-h${h.level}"><a href="#${h.id}">${h.text}</a></li>`;
    });
    toc += '</ul></div>';
    return toc;
}

// 为标题添加 ID
function addHeadingIds(html) {
    return html.replace(/<h([2-4])>(.+?)<\/h[2-4]>/g, (match, level, text) => {
        const id = text.toLowerCase().replace(/\s+/g, '-').replace(/[^\w\u4e00-\u9fa5-]/g, '').replace(/<[^>]+>/g, '');
        return `<h${level} id="${id}">${text}</h${level}>`;
    });
}

// 更新行号
function updateLineNumbers() {
    const lines = markdownInput.value.split('\n').length;
    lineNumbers.innerHTML = Array.from({ length: lines }, (_, i) => i + 1).join('<br>');
}

// 实时预览
function updatePreview() {
    const markdown = markdownInput.value;
    let html = marked.parse(markdown);
    html = renderMath(html);
    html = addHeadingIds(html);
    
    if (showToc.checked) {
        const toc = generateTOC(markdown);
        html = toc + html;
    }
    
    preview.innerHTML = html;
    
    // 统计
    charCount.textContent = `${markdown.length} 字符`;
    const words = markdown.trim().split(/\s+/).filter(w => w).length;
    wordCount.textContent = `${words} 词`;
    
    updateLineNumbers();
}

// 同步滚动行号
markdownInput.addEventListener('scroll', () => {
    lineNumbers.scrollTop = markdownInput.scrollTop;
});

// 事件监听
markdownInput.addEventListener('input', updatePreview);
showToc.addEventListener('change', updatePreview);

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

// 工具栏按钮
document.querySelectorAll('.toolbar-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const action = btn.dataset.action;
        const start = markdownInput.selectionStart;
        const end = markdownInput.selectionEnd;
        const text = markdownInput.value;
        const selected = text.substring(start, end);
        
        let insert = '';
        let cursorOffset = 0;
        
        switch (action) {
            case 'heading':
                insert = `## ${selected || '标题'}`;
                cursorOffset = selected ? insert.length : 3;
                break;
            case 'bold':
                insert = `**${selected || '粗体文本'}**`;
                cursorOffset = selected ? insert.length : 2;
                break;
            case 'italic':
                insert = `*${selected || '斜体文本'}*`;
                cursorOffset = selected ? insert.length : 1;
                break;
            case 'formula':
                insert = `$${selected || 'E=mc^2'}$`;
                cursorOffset = selected ? insert.length : 1;
                break;
            case 'formula-block':
                insert = `\n$$\n${selected || '\\int_{a}^{b} f(x) dx'}\n$$\n`;
                cursorOffset = 3;
                break;
            case 'image':
                insert = `![${selected || '图片描述'}](图片URL)`;
                cursorOffset = 2;
                break;
            case 'table':
                insert = `\n| 列1 | 列2 | 列3 |\n|-----|-----|-----|\n| 内容 | 内容 | 内容 |\n`;
                cursorOffset = insert.length;
                break;
            case 'quote':
                insert = `> ${selected || '引用内容'}`;
                cursorOffset = selected ? insert.length : 2;
                break;
            case 'ref':
                insert = `[^${selected || '1'}]`;
                cursorOffset = 2;
                break;
        }
        
        markdownInput.value = text.substring(0, start) + insert + text.substring(end);
        markdownInput.focus();
        markdownInput.setSelectionRange(start + cursorOffset, start + cursorOffset);
        updatePreview();
    });
});

// 生成目录按钮
tocBtn.addEventListener('click', () => {
    const markdown = markdownInput.value;
    const headings = [];
    const lines = markdown.split('\n');
    lines.forEach(line => {
        const match = line.match(/^(#{1,4})\s+(.+)$/);
        if (match) {
            headings.push({ level: match[1].length, text: match[2] });
        }
    });
    
    if (headings.length === 0) {
        alert('未找到标题，无法生成目录');
        return;
    }
    
    let toc = '## 目录\n\n';
    headings.forEach(h => {
        const indent = '  '.repeat(h.level - 1);
        toc += `${indent}- ${h.text}\n`;
    });
    toc += '\n---\n\n';
    
    markdownInput.value = toc + markdown;
    updatePreview();
});

// 样式弹窗
styleBtn.addEventListener('click', () => styleModal.classList.add('active'));
closeModal.addEventListener('click', () => styleModal.classList.remove('active'));
styleModal.addEventListener('click', (e) => {
    if (e.target === styleModal) styleModal.classList.remove('active');
});

// 模板选择
document.querySelectorAll('.template-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.template-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
    });
});

// 应用样式
document.getElementById('applyStyle').addEventListener('click', () => {
    styleConfig.bodyFont = document.getElementById('bodyFont').value;
    styleConfig.engFont = document.getElementById('engFont').value;
    styleConfig.bodySize = parseFloat(document.getElementById('bodySize').value);
    styleConfig.lineSpacing = parseFloat(document.getElementById('lineSpacing').value);
    styleConfig.firstIndent = parseInt(document.getElementById('firstIndent').value);
    styleModal.classList.remove('active');
});

// 解析 Markdown
function parseMarkdown(markdown) {
    const lines = markdown.split('\n');
    const elements = [];
    let inCodeBlock = false;
    let codeContent = '';
    let inTable = false;
    let tableRows = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        if (line.startsWith('```')) {
            if (!inCodeBlock) {
                inCodeBlock = true;
                codeContent = '';
            } else {
                elements.push({ type: 'code', content: codeContent.trim() });
                inCodeBlock = false;
            }
            continue;
        }

        if (inCodeBlock) {
            codeContent += line + '\n';
            continue;
        }

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
        }

        if (line.trim() === '') continue;

        const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) {
            elements.push({ type: 'heading', level: headingMatch[1].length, content: headingMatch[2] });
            continue;
        }

        if (line.match(/^[-*_]{3,}$/)) {
            elements.push({ type: 'hr' });
            continue;
        }

        if (line.startsWith('>')) {
            elements.push({ type: 'quote', content: line.replace(/^>\s*/, '') });
            continue;
        }

        if (line.match(/^[\s]*[-*+]\s+/)) {
            const content = line.replace(/^[\s]*[-*+]\s+/, '');
            elements.push({ type: 'bullet', content });
            continue;
        }

        if (line.match(/^[\s]*\d+\.\s+/)) {
            const content = line.replace(/^[\s]*\d+\.\s+/, '');
            elements.push({ type: 'number', content });
            continue;
        }

        elements.push({ type: 'paragraph', content: line });
    }

    if (inTable && tableRows.length > 0) {
        elements.push({ type: 'table', rows: tableRows });
    }

    return elements;
}

// 解析行内格式
function parseInlineFormatting(text, baseFont, baseSize) {
    const runs = [];
    let remaining = text;

    while (remaining.length > 0) {
        let match = remaining.match(/\*\*(.+?)\*\*/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], bold: true, font: baseFont, size: baseSize });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        match = remaining.match(/\*(.+?)\*/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], italics: true, font: baseFont, size: baseSize });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        match = remaining.match(/`(.+?)`/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], font: 'Consolas', size: baseSize });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        // 公式标记（在 Word 中显示为文本）
        match = remaining.match(/\$([^$]+)\$/);
        if (match && match.index === 0) {
            runs.push({ text: match[1], font: 'Cambria Math', size: baseSize, italics: true });
            remaining = remaining.slice(match[0].length);
            continue;
        }

        const nextSpecial = remaining.search(/\*\*|\*|`|\$/);
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

// 生成 Word
async function generateWord() {
    try {
        const markdown = markdownInput.value;
        if (!markdown.trim()) {
            alert('请先输入内容');
            return;
        }

        const elements = parseMarkdown(markdown);
        const children = [];
        const bodySize = styleConfig.bodySize * 2;
        const lineSpacingValue = Math.round(styleConfig.lineSpacing * 240);

        const headingSizes = { 1: 44, 2: 36, 3: 32, 4: 28, 5: 26, 6: 24 };

        for (const el of elements) {
            switch (el.type) {
                case 'heading':
                    children.push(new Paragraph({
                        children: [new TextRun({
                            text: el.content,
                            bold: true,
                            size: headingSizes[el.level],
                            font: '黑体'
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
                        bullet: { level: 0 },
                        spacing: { after: 60, line: lineSpacingValue }
                    }));
                    break;

                case 'number':
                    const nRuns = parseInlineFormatting(el.content, styleConfig.bodyFont, bodySize);
                    children.push(new Paragraph({
                        children: nRuns.map(r => new TextRun(r)),
                        numbering: { reference: 'default-numbering', level: 0 },
                        spacing: { after: 60, line: lineSpacingValue }
                    }));
                    break;

                case 'quote':
                    children.push(new Paragraph({
                        children: [new TextRun({
                            text: el.content,
                            italics: true,
                            size: bodySize,
                            font: styleConfig.bodyFont
                        })],
                        indent: { left: 720 },
                        spacing: { after: 120, line: lineSpacingValue }
                    }));
                    break;

                case 'code':
                    el.content.split('\n').forEach(line => {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: line || ' ', font: 'Consolas', size: 20 })],
                            shading: { fill: 'f4f4f4' },
                            spacing: { after: 0, line: 240 }
                        }));
                    });
                    children.push(new Paragraph({ children: [] }));
                    break;

                case 'table':
                    if (el.rows.length > 0) {
                        const tableRows = el.rows.map((row, idx) => new TableRow({
                            children: row.map(cell => new TableCell({
                                children: [new Paragraph({
                                    children: [new TextRun({ text: cell, bold: idx === 0, size: bodySize, font: styleConfig.bodyFont })]
                                })],
                                shading: idx === 0 ? { fill: 'f0f0f0' } : undefined
                            }))
                        }));
                        children.push(new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
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

        const doc = new Document({
            numbering: {
                config: [{
                    reference: 'default-numbering',
                    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START }]
                }]
            },
            sections: [{
                properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
                children
            }]
        });

        const blob = await Packer.toBlob(doc);
        const filename = prompt('请输入文件名（不含扩展名）：', '学术文档') || '学术文档';
        saveAs(blob, filename + '.docx');
    } catch (error) {
        console.error(error);
        alert('生成失败: ' + error.message);
    }
}

downloadBtn.addEventListener('click', generateWord);
updatePreview();
