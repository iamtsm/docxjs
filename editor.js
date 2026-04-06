/**
 * editor.js
 * Rich text editor logic: toolbar commands, file open/export, state management.
 *
 * All editing operations are exposed via the global `window.DocxEditor` object.
 * External code can call methods like:
 *   DocxEditor.bold()
 *   DocxEditor.setFontSize('14pt')
 *   DocxEditor.exportDocx('my-file.docx')
 *   DocxEditor.getState()
 */

window.DocxEditor = (function () {
    'use strict';

    // ========== DOM References ==========
    const editorPage = document.getElementById('editor-page');
    const fileInput = document.getElementById('file-input');
    const btnNew = document.getElementById('btn-new');
    const btnExport = document.getElementById('btn-export');
    const fileName = document.getElementById('file-name');
    const statusText = document.getElementById('status-text');
    const loadingOverlay = document.getElementById('loading-overlay');

    // Toolbar controls
    const btnBold = document.getElementById('btn-bold');
    const btnItalic = document.getElementById('btn-italic');
    const btnUnderline = document.getElementById('btn-underline');
    const btnStrike = document.getElementById('btn-strike');
    const selectFontSize = document.getElementById('select-font-size');
    const inputColor = document.getElementById('input-color');
    const selectHeading = document.getElementById('select-heading');
    const btnAlignLeft = document.getElementById('btn-align-left');
    const btnAlignCenter = document.getElementById('btn-align-center');
    const btnAlignRight = document.getElementById('btn-align-right');
    const btnOrderedList = document.getElementById('btn-ordered-list');
    const btnUnorderedList = document.getElementById('btn-unordered-list');

    let currentFileName = '';

    // ========== Internal Utilities ==========
    function showLoading() { if (loadingOverlay) loadingOverlay.classList.remove('hidden'); }
    function hideLoading() { if (loadingOverlay) loadingOverlay.classList.add('hidden'); }

    function _setStatus(text) {
        if (statusText) statusText.textContent = text;
    }

    function _setFileName(name) {
        currentFileName = name;
        if (fileName) fileName.textContent = name || '未命名文档';
    }

    function focusEditor() {
        if (editorPage) editorPage.focus();
    }

    function execCmd(command, value) {
        focusEditor();
        document.execCommand(command, false, value || null);
        updateToolbarState();
    }

    function rgbToHex(color) {
        if (!color) return '#000000';
        if (color.startsWith('#')) return color;
        const m = color.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
        if (m) {
            return '#' + [m[1], m[2], m[3]].map(x => parseInt(x).toString(16).padStart(2, '0')).join('');
        }
        return '#000000';
    }

    // Unwrap docx-rendered HTML to make it more editable
    function unwrapDocxContent() {
        const wrapper = editorPage.querySelector('.docx-wrapper');
        if (!wrapper) return;

        const fragment = document.createDocumentFragment();
        const sections = wrapper.querySelectorAll('section.docx > article');
        sections.forEach(article => {
            while (article.firstChild) {
                fragment.appendChild(article.firstChild);
            }
        });

        const styles = editorPage.querySelectorAll('style');
        const styleFragment = document.createDocumentFragment();
        styles.forEach(s => styleFragment.appendChild(s.cloneNode(true)));

        editorPage.innerHTML = '';
        editorPage.appendChild(styleFragment);
        editorPage.appendChild(fragment);
    }

    // ========== Public API: Formatting ==========

    /** Toggle bold on the current selection */
    function bold() {
        execCmd('bold');
    }

    /** Toggle italic on the current selection */
    function italic() {
        execCmd('italic');
    }

    /** Toggle underline on the current selection */
    function underline() {
        execCmd('underline');
    }

    /** Toggle strikethrough on the current selection */
    function strikeThrough() {
        execCmd('strikeThrough');
    }

    /**
     * Set font size on the current selection.
     * @param {string} size - CSS font-size value, e.g. "14pt", "20px"
     */
    function setFontSize(size) {
        if (!size) return;
        focusEditor();

        const sel = window.getSelection();
        if (!sel.rangeCount) return;

        // Use execCommand fontSize with a marker value, then replace <font> with <span>
        document.execCommand('fontSize', false, '7');

        const fonts = editorPage.querySelectorAll('font[size="7"]');
        fonts.forEach(font => {
            const span = document.createElement('span');
            span.style.fontSize = size;
            span.innerHTML = font.innerHTML;
            font.parentNode.replaceChild(span, font);
        });

        updateToolbarState();
    }

    /**
     * Set font color on the current selection.
     * @param {string} color - CSS color value, e.g. "#FF0000", "red"
     */
    function setFontColor(color) {
        if (!color) return;
        execCmd('foreColor', color);
        if (inputColor) inputColor.value = rgbToHex(color);
    }

    /**
     * Set paragraph/heading style on the current block.
     * @param {string} level - "p", "h1", "h2", "h3", "h4", "h5", "h6"
     */
    function setHeading(level) {
        if (!level) return;
        const tag = level.toLowerCase();
        if (tag === 'p') {
            execCmd('formatBlock', '<p>');
        } else if (/^h[1-6]$/.test(tag)) {
            execCmd('formatBlock', `<${tag}>`);
        }
    }

    /** Align current block to the left */
    function alignLeft() {
        execCmd('justifyLeft');
    }

    /** Align current block to the center */
    function alignCenter() {
        execCmd('justifyCenter');
    }

    /** Align current block to the right */
    function alignRight() {
        execCmd('justifyRight');
    }

    /** Toggle ordered (numbered) list */
    function orderedList() {
        execCmd('insertOrderedList');
    }

    /** Toggle unordered (bullet) list */
    function unorderedList() {
        execCmd('insertUnorderedList');
    }

    // ========== Public API: Document Operations ==========

    /** Create a new blank document (clears editor without confirmation) */
    function newDocument() {
        editorPage.innerHTML = '';
        _setFileName('');
        _setStatus('新文档已创建');
        focusEditor();
    }

    /**
     * Open a .docx file programmatically.
     * @param {Blob|File|ArrayBuffer|Uint8Array} fileBlob - The docx file data
     * @param {string} [name] - Optional display name for the file
     * @returns {Promise<void>}
     */
    async function openFile(fileBlob, name) {
        if (!fileBlob) throw new Error('fileBlob is required');

        showLoading();
        _setStatus('正在打开文件...');

        try {
            if (typeof docx === 'undefined') {
                throw new Error('docx-preview library not loaded');
            }

            editorPage.innerHTML = '';

            await docx.renderAsync(fileBlob, editorPage, null, {
                className: 'docx',
                inWrapper: true,
                ignoreWidth: true,
                ignoreHeight: true,
                ignoreFonts: false,
                breakPages: false,
                debug: false,
                experimental: false,
                renderHeaders: false,
                renderFooters: false,
                renderFootnotes: true,
                renderEndnotes: true,
                renderComments: false,
            });

            unwrapDocxContent();

            const displayName = name || (fileBlob.name ? fileBlob.name : '');
            _setFileName(displayName);
            _setStatus('文件已打开' + (displayName ? ': ' + displayName : ''));
        } catch (err) {
            console.error('Error opening file:', err);
            _setStatus('打开文件失败');
            throw err;
        } finally {
            hideLoading();
        }
    }

    /**
     * Export the current editor content as a .docx file and trigger download.
     * @param {string} [filename] - Output filename (default: current name or "document.docx")
     * @returns {Promise<Blob>} The generated docx Blob
     */
    async function exportDocx(filename) {
        showLoading();
        _setStatus('正在导出...');

        try {
            const exportName = filename || currentFileName || 'document.docx';
            const finalName = exportName.endsWith('.docx') ? exportName : exportName + '.docx';

            const blob = await DocxExporter.exportToDocx(editorPage, finalName);
            _setStatus('导出成功: ' + finalName);
            return blob;
        } catch (err) {
            console.error('Export error:', err);
            _setStatus('导出失败');
            throw err;
        } finally {
            hideLoading();
        }
    }

    /**
     * Get the current HTML content of the editor.
     * @returns {string} HTML string
     */
    function getContent() {
        return editorPage.innerHTML;
    }

    /**
     * Set the editor's HTML content.
     * @param {string} html - HTML string to set
     */
    function setContent(html) {
        editorPage.innerHTML = html || '';
        _setStatus('内容已更新');
    }

    /**
     * Get the formatting state at the current cursor position.
     * @returns {object} State object with boolean/string fields
     */
    function getState() {
        const state = {
            bold: false,
            italic: false,
            underline: false,
            strikeThrough: false,
            orderedList: false,
            unorderedList: false,
            justifyLeft: false,
            justifyCenter: false,
            justifyRight: false,
            formatBlock: '',
            foreColor: '#000000',
        };
        try {
            state.bold = document.queryCommandState('bold');
            state.italic = document.queryCommandState('italic');
            state.underline = document.queryCommandState('underline');
            state.strikeThrough = document.queryCommandState('strikeThrough');
            state.orderedList = document.queryCommandState('insertOrderedList');
            state.unorderedList = document.queryCommandState('insertUnorderedList');
            state.justifyLeft = document.queryCommandState('justifyLeft');
            state.justifyCenter = document.queryCommandState('justifyCenter');
            state.justifyRight = document.queryCommandState('justifyRight');
            state.formatBlock = document.queryCommandValue('formatBlock') || '';
            state.foreColor = rgbToHex(document.queryCommandValue('foreColor'));
        } catch (e) {
            // queryCommandState may throw in some browsers
        }
        return state;
    }

    /**
     * Get the editor DOM element.
     * @returns {HTMLElement}
     */
    function getEditorElement() {
        return editorPage;
    }

    // ========== Toolbar State Sync ==========
    function updateToolbarState() {
        try {
            const s = getState();
            if (btnBold) btnBold.classList.toggle('active', s.bold);
            if (btnItalic) btnItalic.classList.toggle('active', s.italic);
            if (btnUnderline) btnUnderline.classList.toggle('active', s.underline);
            if (btnStrike) btnStrike.classList.toggle('active', s.strikeThrough);
            if (btnAlignLeft) btnAlignLeft.classList.toggle('active', s.justifyLeft);
            if (btnAlignCenter) btnAlignCenter.classList.toggle('active', s.justifyCenter);
            if (btnAlignRight) btnAlignRight.classList.toggle('active', s.justifyRight);
            if (btnOrderedList) btnOrderedList.classList.toggle('active', s.orderedList);
            if (btnUnorderedList) btnUnorderedList.classList.toggle('active', s.unorderedList);
            if (inputColor && s.foreColor) inputColor.value = s.foreColor;
        } catch (e) {
            // safe fallback
        }
    }

    // ========== Wire Toolbar UI to API ==========
    if (btnBold) btnBold.addEventListener('click', () => bold());
    if (btnItalic) btnItalic.addEventListener('click', () => italic());
    if (btnUnderline) btnUnderline.addEventListener('click', () => underline());
    if (btnStrike) btnStrike.addEventListener('click', () => strikeThrough());

    if (selectFontSize) selectFontSize.addEventListener('change', () => {
        const val = selectFontSize.value;
        if (val) {
            setFontSize(val);
            selectFontSize.value = '';
        }
    });

    if (inputColor) inputColor.addEventListener('input', () => {
        setFontColor(inputColor.value);
    });

    if (selectHeading) selectHeading.addEventListener('change', () => {
        const val = selectHeading.value;
        if (val) {
            setHeading(val);
            selectHeading.value = '';
        }
    });

    if (btnAlignLeft) btnAlignLeft.addEventListener('click', () => alignLeft());
    if (btnAlignCenter) btnAlignCenter.addEventListener('click', () => alignCenter());
    if (btnAlignRight) btnAlignRight.addEventListener('click', () => alignRight());
    if (btnOrderedList) btnOrderedList.addEventListener('click', () => orderedList());
    if (btnUnorderedList) btnUnorderedList.addEventListener('click', () => unorderedList());

    if (btnExport) btnExport.addEventListener('click', () => exportDocx());

    if (btnNew) btnNew.addEventListener('click', () => {
        if (editorPage.innerHTML.trim() && !confirm('创建新文档将清除当前内容，确定继续？')) {
            return;
        }
        newDocument();
    });

    if (fileInput) fileInput.addEventListener('change', async () => {
        const file = fileInput.files[0];
        if (!file) return;
        try {
            await openFile(file, file.name);
        } catch (err) {
            alert('打开文件失败: ' + err.message);
        }
        fileInput.value = '';
    });

    // Listen for selection changes to update toolbar
    document.addEventListener('selectionchange', updateToolbarState);
    if (editorPage) {
        editorPage.addEventListener('keyup', updateToolbarState);
        editorPage.addEventListener('mouseup', updateToolbarState);
    }

    // ========== Keyboard Shortcuts ==========
    if (editorPage) editorPage.addEventListener('keydown', (e) => {
        if (e.ctrlKey || e.metaKey) {
            switch (e.key.toLowerCase()) {
                case 'b':
                    e.preventDefault();
                    bold();
                    break;
                case 'i':
                    e.preventDefault();
                    italic();
                    break;
                case 'u':
                    e.preventDefault();
                    underline();
                    break;
                case 's':
                    e.preventDefault();
                    exportDocx();
                    break;
            }
        }
    });

    // ========== Ensure editor always has at least a paragraph ==========
    if (editorPage) editorPage.addEventListener('input', () => {
        if (!editorPage.innerHTML.trim()) {
            editorPage.innerHTML = '<p><br></p>';
            const p = editorPage.querySelector('p');
            if (p) {
                const range = document.createRange();
                range.setStart(p, 0);
                range.collapse(true);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
            }
        }
    });

    // ========== Paste handler - clean paste ==========
    if (editorPage) editorPage.addEventListener('paste', (e) => {
        // Allow rich paste (keep formatting)
        // If you want plain text only, uncomment below:
        // e.preventDefault();
        // const text = e.clipboardData.getData('text/plain');
        // document.execCommand('insertText', false, text);
    });

    // ========== Initialize ==========
    _setFileName('');
    _setStatus('就绪');
    hideLoading();

    if (editorPage) editorPage.setAttribute('contenteditable', 'true');

    // ========== Return Public API ==========
    return {
        // Formatting
        bold: bold,
        italic: italic,
        underline: underline,
        strikeThrough: strikeThrough,
        setFontSize: setFontSize,
        setFontColor: setFontColor,
        setHeading: setHeading,
        alignLeft: alignLeft,
        alignCenter: alignCenter,
        alignRight: alignRight,
        orderedList: orderedList,
        unorderedList: unorderedList,

        // Document operations
        newDocument: newDocument,
        openFile: openFile,
        exportDocx: exportDocx,

        // Content access
        getContent: getContent,
        setContent: setContent,
        getState: getState,
        getEditorElement: getEditorElement,
    };
})();
