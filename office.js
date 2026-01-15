// 办公版 - Markdown 转 Word
const { Document, Paragraph, TextRun, Table, TableRow, TableCell, Packer, WidthType, BorderStyle, LevelFormat, AlignmentType, Header, Footer, PageNumber } = window.docx;

// 配置
let styleConfig = {
    bodyFont: '微软雅黑',
    bodySize: 12,
    lineSpacing: 1.5,
    pageMargin: 'normal',
    headerText: '',
    showPageNum: true
};

// 历史记录（撤销/重做）
let history = [];
let historyIndex = -1;
const MAX_HISTORY = 50;

// DOM
const markdownInput = document.getElementById('markdownInput');
const preview = document.getElementById('preview');
const charCount = document.getElementById('charCount');
const wordCount = document.getElementById('wordCount');
const readTime = document.getElementById('readTime');
const fileInput = document.getElementById('fileInput');
const styleBtn = document.getElementById('styleBtn');
const downloadWord = document.getElementById('downloadWord');
const downloadPdf = document.getElementById('downloadPdf');
const styleModal = document.getElementById('styleModal');
const closeModal = document.getElementById('closeModal');
const imageModal = document.getElementById('imageModal');
const closeImageModal = document.getElementById('closeImageModal');
const dropZone = document.getElementById('dropZone');
const imageInput = document.getElementById('imageInput');
const templateSelect = document.getElementById('templateSelect');

// marked 配置
marked.setOptions({
    highlight: (code, lang) => lang && hljs.getLanguage(lang) ? hljs.highlight(code, { language: lang }).value : hljs.highlightAuto(code).value,
    breaks: true,
    gfm: true
});

// 保存历史
function saveHistory() {
    const content = markdownInput.value;
    if (history[historyIndex] === content) return;
    
    history = history.slice(0, historyIndex + 1);
    history.push(content);
    if (history.length > MAX_HISTORY) history.shift();
    historyIndex = history.length - 1;
}

// 撤销
function undo() {
    if (historyIndex > 0) {
        historyIndex--;
        markdownInput.value = history[historyIndex];
        updatePreview();
    }
}

// 重做
function redo() {
    if (historyIndex < history.length - 1) {
        historyIndex++;
        markdownInput.value = history[historyIndex];
        updatePreview();
    }
}

// 更新预览
function updatePreview() {
    const markdown = markdownInput.value;
    preview.innerHTML = marked.parse(markdown);
    
    // 统计
    const chars = markdown.length;
    const words = markdown.trim().split(/\s+/).filter(w => w).length;
    const minutes = Math.ceil(chars / 500);
    
    charCount.textContent = `${chars} 字`;
    wordCount.textContent = `${words} 词`;
    readTime.textContent = `阅读约 ${minutes} 分钟`;
}

// 插入文本
function insertText(before, after = '', placeholder = '') {
    const start = markdownInput.selectionStart;
    const end = markdownInput.selectionEnd;
    const text = markdownInput.value;
    const selected = text.substring(start, end) || placeholder;
    
    const insert = before + selected + after;
    markdownInput.value = text.substring(0, start) + insert + text.substring(end);
    
    markdownInput.focus();
    const newPos = start + before.length + selected.length;
    markdownInput.setSelectionRange(start + before.length, newPos);
    
    saveHistory();
    updatePreview();
}

// 工具栏按钮
document.querySelectorAll('.visual-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        const action = btn.dataset.action;
        
        switch (action) {
            case 'h1': insertText('# ', '', '一级标题'); break;
            case 'h2': insertText('## ', '', '二级标题'); break;
            case 'h3': insertText('### ', '', '三级标题'); break;
            case 'bold': insertText('**', '**', '粗体文本'); break;
            case 'italic': insertText('*', '*', '斜体文本'); break;
            case 'strike': insertText('~~', '~~', '删除线'); break;
            case 'ul': insertText('- ', '', '列表项'); break;
            case 'ol': insertText('1. ', '', '列表项'); break;
            case 'quote': insertText('> ', '', '引用内容'); break;
            case 'table': insertText('\n| 列1 | 列2 | 列3 |\n|-----|-----|-----|\n| 内容 | 内容 | 内容 |\n'); break;
            case 'link': insertText('[', '](url)', '链接文本'); break;
            case 'hr': insertText('\n---\n'); break;
            case 'image': imageModal.classList.add('active'); break;
            case 'undo': undo(); break;
            case 'redo': redo(); break;
        }
    });
});

// 事件监听
markdownInput.addEventListener('input', () => {
    saveHistory();
    updatePreview();
});

// 文件上传
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
            markdownInput.value = event.target.result;
            saveHistory();
            updatePreview();
        };
        reader.readAsText(file);
    }
});

// 图片弹窗
closeImageModal.addEventListener('click', () => imageModal.classList.remove('active'));
document.getElementById('cancelImage').addEventListener('click', () => imageModal.classList.remove('active'));

// 图片拖拽
dropZone.addEventListener('click', () => imageInput.click());
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith('image/')) {
        handleImageFile(file);
    }
});

imageInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleImageFile(file);
});

let currentImageData = '';

function handleImageFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        currentImageData = e.target.result;
        dropZone.innerHTML = `<img src="${currentImageData}" style="max-width:100%;max-height:150px;border-radius:8px;">`;
    };
    reader.readAsDataURL(file);
}

document.getElementById('insertImage').addEventListener('click', () => {
    const alt = document.getElementById('imageAlt').value || '图片';
    if (currentImageData) {
        insertText(`![${alt}](${currentImageData})`);
    }
    imageModal.classList.remove('active');
    currentImageData = '';
    dropZone.innerHTML = '<div class="drop-zone-text">拖拽图片到此处<br>或点击选择文件</div>';
    document.getElementById('imageAlt').value = '';
});

// 编辑器拖拽图片
markdownInput.addEventListener('dragover', (e) => e.preventDefault());
markdownInput.addEventListener('drop', (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith('image/')) {
        const reader = new FileReader();
        reader.onload = (ev) => {
            insertText(`![图片](${ev.target.result})`);
        };
        reader.readAsDataURL(file);
    }
});

// 样式弹窗
styleBtn.addEventListener('click', () => styleModal.classList.add('active'));
closeModal.addEventListener('click', () => styleModal.classList.remove('active'));
styleModal.addEventListener('click', (e) => { if (e.target === styleModal) styleModal.classList.remove('active'); });

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
    styleConfig.bodySize = parseFloat(document.getElementById('bodySize').value);
    styleConfig.lineSpacing = parseFloat(document.getElementById('lineSpacing').value);
    styleConfig.pageMargin = document.getElementById('pageMargin').value;
    styleConfig.headerText = document.getElementById('headerText').value;
    styleConfig.showPageNum = document.getElementById('showPageNum').value === 'yes';
    styleModal.classList.remove('active');
});

// 重置样式
document.getElementById('resetStyle').addEventListener('click', () => {
    document.getElementById('bodyFont').value = '微软雅黑';
    document.getElementById('bodySize').value = '12';
    document.getElementById('lineSpacing').value = '1.5';
    document.getElementById('pageMargin').value = 'normal';
    document.getElementById('headerText').value = '';
    document.getElementById('showPageNum').value = 'yes';
});

// 文件名弹窗
const filenameModal = document.getElementById('filenameModal');
const filenameInput = document.getElementById('filenameInput');
let currentExportType = 'word';

document.getElementById('closeFilenameModal').addEventListener('click', () => filenameModal.classList.remove('active'));
document.getElementById('cancelFilename').addEventListener('click', () => filenameModal.classList.remove('active'));

function showFilenameModal(type, defaultName) {
    currentExportType = type;
    filenameInput.value = '';
    filenameInput.placeholder = defaultName;
    filenameModal.classList.add('active');
    filenameInput.focus();
}

document.getElementById('confirmFilename').addEventListener('click', () => {
    const filename = filenameInput.value.trim() || filenameInput.placeholder;
    filenameModal.classList.remove('active');
    if (currentExportType === 'word') {
        doGenerateWord(filename);
    } else {
        doGeneratePdf(filename);
    }
});

// 回车确认
filenameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        document.getElementById('confirmFilename').click();
    }
});

// 解析 Markdown
function parseMarkdown(markdown) {
    const lines = markdown.split('\n');
    const elements = [];
    let inCodeBlock = false, codeContent = '', inTable = false, tableRows = [];

    for (const line of lines) {
        if (line.startsWith('```')) {
            if (!inCodeBlock) { inCodeBlock = true; codeContent = ''; }
            else { elements.push({ type: 'code', content: codeContent.trim() }); inCodeBlock = false; }
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
        
        const h = line.match(/^(#{1,6})\s+(.+)$/);
        if (h) { elements.push({ type: 'heading', level: h[1].length, content: h[2] }); continue; }
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
        
        m = remaining.match(/~~(.+?)~~/);
        if (m && m.index === 0) { runs.push({ text: m[1], strike: true, font, size }); remaining = remaining.slice(m[0].length); continue; }
        
        m = remaining.match(/`(.+?)`/);
        if (m && m.index === 0) { runs.push({ text: m[1], font: 'Consolas', size }); remaining = remaining.slice(m[0].length); continue; }

        const next = remaining.search(/\*\*|\*|~~|`/);
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
        const lineSpacing = Math.round(styleConfig.lineSpacing * 240);
        const margins = {
            normal: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
            narrow: { top: 720, right: 720, bottom: 720, left: 720 },
            wide: { top: 1800, right: 1800, bottom: 1800, left: 1800 }
        };
        const headingSizes = { 1: 48, 2: 40, 3: 32, 4: 28, 5: 24, 6: 22 };

        for (const el of elements) {
            switch (el.type) {
                case 'heading':
                    children.push(new Paragraph({
                        children: [new TextRun({ text: el.content, bold: true, size: headingSizes[el.level], font: '黑体' })],
                        spacing: { before: 240, after: 120, line: lineSpacing }
                    }));
                    break;
                case 'paragraph':
                    children.push(new Paragraph({
                        children: parseInline(el.content, styleConfig.bodyFont, bodySize).map(r => new TextRun(r)),
                        spacing: { after: 120, line: lineSpacing },
                        indent: { firstLine: 480 }
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
                        indent: { left: 720 },
                        border: { left: { style: BorderStyle.SINGLE, size: 24, color: '38d9a9' } },
                        spacing: { after: 120, line: lineSpacing }
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
                        children.push(new Table({
                            rows: el.rows.map((row, idx) => new TableRow({
                                children: row.map(cell => new TableCell({
                                    children: [new Paragraph({
                                        children: [new TextRun({ text: cell, bold: idx === 0, size: bodySize, font: styleConfig.bodyFont })]
                                    })],
                                    shading: idx === 0 ? { fill: 'e6fcf5' } : undefined
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
                        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'cccccc' } },
                        spacing: { before: 240, after: 240 }
                    }));
                    break;
            }
        }

        // 页眉页脚配置
        const sectionProps = {
            page: { margin: margins[styleConfig.pageMargin] }
        };

        if (styleConfig.headerText) {
            sectionProps.headers = {
                default: new Header({
                    children: [new Paragraph({
                        children: [new TextRun({ text: styleConfig.headerText, size: 20, color: '888888' })],
                        alignment: AlignmentType.CENTER
                    })]
                })
            };
        }

        if (styleConfig.showPageNum) {
            sectionProps.footers = {
                default: new Footer({
                    children: [new Paragraph({
                        children: [new TextRun({ children: [PageNumber.CURRENT], size: 20 })],
                        alignment: AlignmentType.CENTER
                    })]
                })
            };
        }

        const doc = new Document({
            numbering: { config: [{ reference: 'default-numbering', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START }] }] },
            sections: [{ properties: sectionProps, children }]
        });

        const blob = await Packer.toBlob(doc);
        saveAs(blob, filename + '.docx');
    } catch (error) {
        console.error(error);
        alert('生成失败: ' + error.message);
    }
}

function generateWord() {
    if (!markdownInput.value.trim()) { alert('请先输入内容'); return; }
    showFilenameModal('word', '办公文档');
}

// 导出 PDF
function doGeneratePdf(filename) {
    const content = preview.cloneNode(true);
    const opt = {
        margin: 10,
        filename: filename + '.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };
    html2pdf().set(opt).from(content).save();
}

function generatePdf() {
    if (!markdownInput.value.trim()) { alert('请先输入内容'); return; }
    showFilenameModal('pdf', '办公文档');
}

downloadWord.addEventListener('click', generateWord);
downloadPdf.addEventListener('click', generatePdf);

// 初始化
saveHistory();
updatePreview();
