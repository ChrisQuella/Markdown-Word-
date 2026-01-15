// 开发者版 - Markdown 转 Word
const { Document, Paragraph, TextRun, Table, TableRow, TableCell, Packer, WidthType, BorderStyle, LevelFormat, AlignmentType } = window.docx;

// 配置
let styleConfig = {
    bodyFont: '微软雅黑',
    bodySize: 11,
    codeFont: 'Consolas',
    codeSize: 11,
    lineSpacing: 1.5
};

let autoSaveInterval = 30;
let autoSaveTimer = null;
let isDirty = false;

// DOM
const markdownInput = document.getElementById('markdownInput');
const preview = document.getElementById('preview');
const charCount = document.getElementById('charCount');
const wordCount = document.getElementById('wordCount');
const lineCount = document.getElementById('lineCount');
const lineNumbers = document.getElementById('lineNumbers');
const saveIndicator = document.getElementById('saveIndicator');
const saveStatus = document.getElementById('saveStatus');
const fileInput = document.getElementById('fileInput');
const styleBtn = document.getElementById('styleBtn');
const downloadBtn = document.getElementById('downloadBtn');
const styleModal = document.getElementById('styleModal');
const closeModal = document.getElementById('closeModal');
const previewTheme = document.getElementById('previewTheme');

// 本地存储 key
const STORAGE_KEY = 'md-developer-draft';

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

// 更新行号
function updateLineNumbers() {
    const lines = markdownInput.value.split('\n').length;
    lineNumbers.innerHTML = Array.from({ length: lines }, (_, i) => i + 1).join('<br>');
}

// 更新光标位置
function updateCursorPosition() {
    const text = markdownInput.value;
    const pos = markdownInput.selectionStart;
    const lines = text.substring(0, pos).split('\n');
    const ln = lines.length;
    const col = lines[lines.length - 1].length + 1;
    charCount.textContent = `Ln ${ln}, Col ${col}`;
}

// 实时预览
function updatePreview() {
    const markdown = markdownInput.value;
    preview.innerHTML = marked.parse(markdown);
    
    // 统计
    const words = markdown.trim().split(/\s+/).filter(w => w).length;
    wordCount.textContent = `${words} words`;
    lineCount.textContent = `${markdown.split('\n').length} lines`;
    
    updateLineNumbers();
    markDirty();
}

// 标记未保存
function markDirty() {
    isDirty = true;
    saveIndicator.classList.add('unsaved');
    saveStatus.textContent = '未保存';
}

// 保存草稿
function saveDraft() {
    localStorage.setItem(STORAGE_KEY, markdownInput.value);
    isDirty = false;
    saveIndicator.classList.remove('unsaved');
    saveStatus.textContent = '已保存';
}

// 加载草稿
function loadDraft() {
    const draft = localStorage.getItem(STORAGE_KEY);
    if (draft) {
        markdownInput.value = draft;
        updatePreview();
    }
}

// 自动保存
function setupAutoSave() {
    if (autoSaveTimer) clearInterval(autoSaveTimer);
    if (autoSaveInterval > 0) {
        autoSaveTimer = setInterval(() => {
            if (isDirty) saveDraft();
        }, autoSaveInterval * 1000);
    }
}

// 同步滚动
markdownInput.addEventListener('scroll', () => {
    lineNumbers.scrollTop = markdownInput.scrollTop;
});

// 事件监听
markdownInput.addEventListener('input', updatePreview);
markdownInput.addEventListener('click', updateCursorPosition);
markdownInput.addEventListener('keyup', updateCursorPosition);

// 快捷键
markdownInput.addEventListener('keydown', (e) => {
    if (e.ctrlKey) {
        let handled = true;
        const start = markdownInput.selectionStart;
        const end = markdownInput.selectionEnd;
        const text = markdownInput.value;
        const selected = text.substring(start, end);
        
        switch (e.key.toLowerCase()) {
            case 'b': // 粗体
                insertText(`**${selected || '粗体'}**`, selected ? 0 : 2);
                break;
            case 'i': // 斜体
                insertText(`*${selected || '斜体'}*`, selected ? 0 : 1);
                break;
            case '`': // 行内代码
                insertText(`\`${selected || '代码'}\``, selected ? 0 : 1);
                break;
            case 'k': // 链接
                if (e.shiftKey) {
                    insertText('\n```\n' + (selected || '// code') + '\n```\n', 4);
                } else {
                    insertText(`[${selected || '链接文本'}](url)`, selected ? 0 : 1);
                }
                break;
            case 's': // 保存
                e.preventDefault();
                saveDraft();
                break;
            case 'h': // 标题
                insertText(`# ${selected || '标题'}`, selected ? 0 : 2);
                break;
            default:
                handled = false;
        }
        
        if (handled && e.key !== 's') {
            e.preventDefault();
            updatePreview();
        }
    }
});

function insertText(insert, cursorOffset) {
    const start = markdownInput.selectionStart;
    const end = markdownInput.selectionEnd;
    const text = markdownInput.value;
    markdownInput.value = text.substring(0, start) + insert + text.substring(end);
    markdownInput.focus();
    const newPos = start + insert.length - cursorOffset;
    markdownInput.setSelectionRange(newPos, newPos);
}

// 工具栏
document.querySelectorAll('.toolbar-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const action = btn.dataset.action;
        const start = markdownInput.selectionStart;
        const end = markdownInput.selectionEnd;
        const text = markdownInput.value;
        const selected = text.substring(start, end);
        
        let insert = '';
        let offset = 0;
        
        switch (action) {
            case 'heading': insert = `# ${selected || '标题'}`; offset = selected ? 0 : 2; break;
            case 'bold': insert = `**${selected || '粗体'}**`; offset = selected ? 0 : 2; break;
            case 'italic': insert = `*${selected || '斜体'}*`; offset = selected ? 0 : 1; break;
            case 'code': insert = `\`${selected || '代码'}\``; offset = selected ? 0 : 1; break;
            case 'codeblock': insert = '\n```javascript\n' + (selected || '// code') + '\n```\n'; offset = 4; break;
            case 'link': insert = `[${selected || '链接'}](url)`; offset = selected ? 0 : 1; break;
            case 'table': insert = '\n| 列1 | 列2 |\n|-----|-----|\n| 值1 | 值2 |\n'; offset = 0; break;
            case 'list': insert = `\n- ${selected || '列表项'}\n`; offset = selected ? 0 : 2; break;
        }
        
        markdownInput.value = text.substring(0, start) + insert + text.substring(end);
        markdownInput.focus();
        markdownInput.setSelectionRange(start + insert.length - offset, start + insert.length - offset);
        updatePreview();
    });
});

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

// 预览主题切换
previewTheme.addEventListener('change', () => {
    const theme = previewTheme.value;
    const previewPanel = document.querySelector('.preview-panel');
    if (theme === 'light') {
        preview.style.background = '#fff';
        preview.style.color = '#333';
    } else {
        preview.style.background = '#1e1e1e';
        preview.style.color = '#ccc';
    }
});

// 弹窗
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

// 应用设置
document.getElementById('applyStyle').addEventListener('click', () => {
    styleConfig.bodyFont = document.getElementById('bodyFont').value;
    styleConfig.bodySize = parseFloat(document.getElementById('bodySize').value);
    styleConfig.codeFont = document.getElementById('codeFont').value;
    styleConfig.codeSize = parseFloat(document.getElementById('codeSize').value);
    autoSaveInterval = parseInt(document.getElementById('autoSaveInterval').value);
    setupAutoSave();
    styleModal.classList.remove('active');
});

// 清除草稿
document.getElementById('clearDraft').addEventListener('click', () => {
    if (confirm('确定清除本地草稿？')) {
        localStorage.removeItem(STORAGE_KEY);
        markdownInput.value = '';
        updatePreview();
    }
});

// 解析 Markdown
function parseMarkdown(markdown) {
    const lines = markdown.split('\n');
    const elements = [];
    let inCodeBlock = false;
    let codeContent = '';
    let codeLang = '';
    let inTable = false;
    let tableRows = [];

    for (const line of lines) {
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

        if (inCodeBlock) { codeContent += line + '\n'; continue; }

        if (line.includes('|') && line.trim().startsWith('|')) {
            if (!inTable) { inTable = true; tableRows = []; }
            if (!line.match(/^\|[\s-:|]+\|$/)) {
                tableRows.push(line.split('|').filter(c => c.trim()).map(c => c.trim()));
            }
            continue;
        } else if (inTable) {
            elements.push({ type: 'table', rows: tableRows });
            inTable = false;
        }

        if (line.trim() === '') continue;

        const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) { elements.push({ type: 'heading', level: headingMatch[1].length, content: headingMatch[2] }); continue; }
        if (line.match(/^[-*_]{3,}$/)) { elements.push({ type: 'hr' }); continue; }
        if (line.startsWith('>')) { elements.push({ type: 'quote', content: line.replace(/^>\s*/, '') }); continue; }
        if (line.match(/^[\s]*[-*+]\s+/)) { elements.push({ type: 'bullet', content: line.replace(/^[\s]*[-*+]\s+/, '') }); continue; }
        if (line.match(/^[\s]*\d+\.\s+/)) { elements.push({ type: 'number', content: line.replace(/^[\s]*\d+\.\s+/, '') }); continue; }
        elements.push({ type: 'paragraph', content: line });
    }

    if (inTable && tableRows.length > 0) elements.push({ type: 'table', rows: tableRows });
    return elements;
}

// 解析行内格式
function parseInline(text, font, size) {
    const runs = [];
    let remaining = text;

    while (remaining.length > 0) {
        let m = remaining.match(/\*\*(.+?)\*\*/);
        if (m && m.index === 0) { runs.push({ text: m[1], bold: true, font, size }); remaining = remaining.slice(m[0].length); continue; }
        
        m = remaining.match(/\*(.+?)\*/);
        if (m && m.index === 0) { runs.push({ text: m[1], italics: true, font, size }); remaining = remaining.slice(m[0].length); continue; }
        
        m = remaining.match(/`(.+?)`/);
        if (m && m.index === 0) { runs.push({ text: m[1], font: styleConfig.codeFont, size }); remaining = remaining.slice(m[0].length); continue; }

        const next = remaining.search(/\*\*|\*|`/);
        if (next > 0) { runs.push({ text: remaining.slice(0, next), font, size }); remaining = remaining.slice(next); }
        else if (next === -1) { runs.push({ text: remaining, font, size }); break; }
        else { runs.push({ text: remaining[0], font, size }); remaining = remaining.slice(1); }
    }

    return runs.length > 0 ? runs : [{ text, font, size }];
}

// 生成 Word
async function doGenerateWord(filename) {
    try {
        const markdown = markdownInput.value;
        if (!markdown.trim()) { alert('请先输入内容'); return; }

        const elements = parseMarkdown(markdown);
        const children = [];
        const bodySize = styleConfig.bodySize * 2;
        const codeSize = styleConfig.codeSize * 2;
        const lineSpacing = Math.round(styleConfig.lineSpacing * 240);
        const headingSizes = { 1: 48, 2: 40, 3: 32, 4: 28, 5: 24, 6: 22 };

        for (const el of elements) {
            switch (el.type) {
                case 'heading':
                    children.push(new Paragraph({
                        children: [new TextRun({ text: el.content, bold: true, size: headingSizes[el.level], font: styleConfig.bodyFont })],
                        spacing: { before: 200, after: 100, line: lineSpacing }
                    }));
                    break;
                case 'paragraph':
                    children.push(new Paragraph({
                        children: parseInline(el.content, styleConfig.bodyFont, bodySize).map(r => new TextRun(r)),
                        spacing: { after: 100, line: lineSpacing }
                    }));
                    break;
                case 'bullet':
                    children.push(new Paragraph({
                        children: parseInline(el.content, styleConfig.bodyFont, bodySize).map(r => new TextRun(r)),
                        bullet: { level: 0 },
                        spacing: { after: 60, line: lineSpacing }
                    }));
                    break;
                case 'number':
                    children.push(new Paragraph({
                        children: parseInline(el.content, styleConfig.bodyFont, bodySize).map(r => new TextRun(r)),
                        numbering: { reference: 'default-numbering', level: 0 },
                        spacing: { after: 60, line: lineSpacing }
                    }));
                    break;
                case 'quote':
                    children.push(new Paragraph({
                        children: [new TextRun({ text: el.content, italics: true, size: bodySize, font: styleConfig.bodyFont, color: '666666' })],
                        indent: { left: 480 },
                        spacing: { after: 100, line: lineSpacing }
                    }));
                    break;
                case 'code':
                    // 添加语言标签
                    if (el.lang) {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: el.lang, size: 18, font: styleConfig.codeFont, color: '888888' })],
                            shading: { fill: '2d2d2d' }
                        }));
                    }
                    el.content.split('\n').forEach(line => {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: line || ' ', font: styleConfig.codeFont, size: codeSize, color: 'abb2bf' })],
                            shading: { fill: '282c34' },
                            spacing: { after: 0, line: 240 }
                        }));
                    });
                    children.push(new Paragraph({ children: [] }));
                    break;
                case 'table':
                    if (el.rows.length > 0) {
                        children.push(new Table({
                            rows: el.rows.map((row, idx) => new TableRow({
                                children: row.map(cell => new TableCell({
                                    children: [new Paragraph({
                                        children: [new TextRun({ text: cell, bold: idx === 0, size: bodySize, font: styleConfig.bodyFont })]
                                    })],
                                    shading: idx === 0 ? { fill: '2d2d2d' } : undefined
                                }))
                            })),
                            width: { size: 100, type: WidthType.PERCENTAGE }
                        }));
                        children.push(new Paragraph({ children: [] }));
                    }
                    break;
                case 'hr':
                    children.push(new Paragraph({
                        children: [],
                        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '444444' } },
                        spacing: { before: 200, after: 200 }
                    }));
                    break;
            }
        }

        const doc = new Document({
            numbering: { config: [{ reference: 'default-numbering', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START }] }] },
            sections: [{ properties: { page: { margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } }, children }]
        });

        const blob = await Packer.toBlob(doc);
        saveAs(blob, filename + '.docx');
    } catch (error) {
        console.error(error);
        alert('生成失败: ' + error.message);
    }
}

// 文件名弹窗
const filenameModal = document.getElementById('filenameModal');
const filenameInput = document.getElementById('filenameInput');

document.getElementById('closeFilenameModal').addEventListener('click', () => filenameModal.classList.remove('active'));
document.getElementById('cancelFilename').addEventListener('click', () => filenameModal.classList.remove('active'));

function showFilenameModal() {
    filenameInput.value = '';
    filenameModal.classList.add('active');
    filenameInput.focus();
}

document.getElementById('confirmFilename').addEventListener('click', () => {
    const filename = filenameInput.value.trim() || '技术文档';
    filenameModal.classList.remove('active');
    doGenerateWord(filename);
});

filenameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') document.getElementById('confirmFilename').click();
});

function generateWord() {
    if (!markdownInput.value.trim()) { alert('请先输入内容'); return; }
    showFilenameModal();
}

downloadBtn.addEventListener('click', generateWord);

// 初始化
loadDraft();
setupAutoSave();
updatePreview();
updateCursorPosition();
