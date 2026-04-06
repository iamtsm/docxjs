/**
 * docx-editor.ts
 * Rich text editor logic: formatting commands, file open/export, state management.
 * Provides a DocxEditor class that can be attached to any container element
 * to enable editing of rendered docx content.
 */

import { exportToDocx, exportAndDownload } from './docx-exporter';

export interface EditorState {
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strikeThrough: boolean;
    orderedList: boolean;
    unorderedList: boolean;
    justifyLeft: boolean;
    justifyCenter: boolean;
    justifyRight: boolean;
    formatBlock: string;
    foreColor: string;
}

export interface EditorOptions {
    /** Whether to enable contenteditable immediately (default: true) */
    editable?: boolean;
    /** Whether to unwrap docx-wrapper structure for better editing (default: true) */
    unwrapDocx?: boolean;
}

const defaultEditorOptions: EditorOptions = {
    editable: true,
    unwrapDocx: true,
};

export class DocxEditor {
    private container: HTMLElement;
    private options: EditorOptions;
    private _currentFileName: string = '';

    constructor(container: HTMLElement, options?: EditorOptions) {
        if (!container) throw new Error('DocxEditor: container element is required');
        this.container = container;
        this.options = { ...defaultEditorOptions, ...options };

        if (this.options.editable) {
            this.enableEditing();
        }

        if (this.options.unwrapDocx) {
            this.unwrapDocxContent();
        }
    }

    // ========== Editing Control ==========

    /** Enable contenteditable on the container */
    enableEditing(): void {
        this.container.setAttribute('contenteditable', 'true');
    }

    /** Disable contenteditable on the container */
    disableEditing(): void {
        this.container.setAttribute('contenteditable', 'false');
    }

    /** Check if editing is enabled */
    isEditable(): boolean {
        return this.container.getAttribute('contenteditable') === 'true';
    }

    /** Get the underlying container DOM element */
    getEditorElement(): HTMLElement {
        return this.container;
    }

    // ========== Internal Utilities ==========

    private focusEditor(): void {
        this.container.focus();
    }

    private execCmd(command: string, value?: string): void {
        this.focusEditor();
        document.execCommand(command, false, value || null);
    }

    private static rgbToHex(color: string): string {
        if (!color) return '#000000';
        if (color.startsWith('#')) return color;
        const m = color.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
        if (m) {
            return '#' + [m[1], m[2], m[3]].map(x => parseInt(x).toString(16).padStart(2, '0')).join('');
        }
        return '#000000';
    }

    /**
     * Unwrap docx-rendered HTML to make it more editable.
     * Extracts content from .docx-wrapper > section.docx > article structure.
     */
    unwrapDocxContent(): void {
        const wrapper = this.container.querySelector('.docx-wrapper');
        if (!wrapper) return;

        const fragment = document.createDocumentFragment();
        const sections = wrapper.querySelectorAll('section.docx > article');
        sections.forEach(article => {
            while (article.firstChild) {
                fragment.appendChild(article.firstChild);
            }
        });

        const styles = this.container.querySelectorAll('style');
        const styleFragment = document.createDocumentFragment();
        styles.forEach(s => styleFragment.appendChild(s.cloneNode(true)));

        this.container.innerHTML = '';
        this.container.appendChild(styleFragment);
        this.container.appendChild(fragment);
    }

    // ========== Formatting API ==========

    /** Toggle bold on the current selection */
    bold(): void {
        this.execCmd('bold');
    }

    /** Toggle italic on the current selection */
    italic(): void {
        this.execCmd('italic');
    }

    /** Toggle underline on the current selection */
    underline(): void {
        this.execCmd('underline');
    }

    /** Toggle strikethrough on the current selection */
    strikeThrough(): void {
        this.execCmd('strikeThrough');
    }

    /**
     * Set font size on the current selection.
     * @param size - CSS font-size value, e.g. "14pt", "20px"
     */
    setFontSize(size: string): void {
        if (!size) return;
        this.focusEditor();

        const sel = window.getSelection();
        if (!sel.rangeCount) return;

        // Use execCommand fontSize with a marker value, then replace <font> with <span>
        document.execCommand('fontSize', false, '7');

        const fonts = this.container.querySelectorAll('font[size="7"]');
        fonts.forEach(font => {
            const span = document.createElement('span');
            span.style.fontSize = size;
            span.innerHTML = font.innerHTML;
            font.parentNode.replaceChild(span, font);
        });
    }

    /**
     * Set font color on the current selection.
     * @param color - CSS color value, e.g. "#FF0000", "red"
     */
    setFontColor(color: string): void {
        if (!color) return;
        this.execCmd('foreColor', color);
    }

    /**
     * Set paragraph/heading style on the current block.
     * @param level - "p", "h1", "h2", "h3", "h4", "h5", "h6"
     */
    setHeading(level: string): void {
        if (!level) return;
        const tag = level.toLowerCase();
        if (tag === 'p') {
            this.execCmd('formatBlock', '<p>');
        } else if (/^h[1-6]$/.test(tag)) {
            this.execCmd('formatBlock', `<${tag}>`);
        }
    }

    /** Align current block to the left */
    alignLeft(): void {
        this.execCmd('justifyLeft');
    }

    /** Align current block to the center */
    alignCenter(): void {
        this.execCmd('justifyCenter');
    }

    /** Align current block to the right */
    alignRight(): void {
        this.execCmd('justifyRight');
    }

    /** Align current block to justify */
    alignJustify(): void {
        this.execCmd('justifyFull');
    }

    /** Toggle ordered (numbered) list */
    orderedList(): void {
        this.execCmd('insertOrderedList');
    }

    /** Toggle unordered (bullet) list */
    unorderedList(): void {
        this.execCmd('insertUnorderedList');
    }

    /** Indent the current block */
    indent(): void {
        this.execCmd('indent');
    }

    /** Outdent the current block */
    outdent(): void {
        this.execCmd('outdent');
    }

    /** Undo the last action */
    undo(): void {
        this.execCmd('undo');
    }

    /** Redo the last undone action */
    redo(): void {
        this.execCmd('redo');
    }

    /** Remove all formatting from the current selection */
    removeFormat(): void {
        this.execCmd('removeFormat');
    }

    // ========== Document Operations ==========

    /** Create a new blank document (clears editor) */
    newDocument(): void {
        this.container.innerHTML = '';
        this._currentFileName = '';
    }

    /**
     * Open a .docx file and render it into the editor container.
     * Requires the docx-preview library's renderAsync to be available globally.
     * @param fileBlob - The docx file data
     * @param name - Optional display name for the file
     */
    async openFile(fileBlob: Blob | File | ArrayBuffer | Uint8Array, name?: string): Promise<void> {
        if (!fileBlob) throw new Error('fileBlob is required');

        // Access docx.renderAsync from global scope
        const docxLib = (window as any).docx;
        if (!docxLib || typeof docxLib.renderAsync !== 'function') {
            throw new Error('docx-preview library (docx.renderAsync) not available on window');
        }

        this.container.innerHTML = '';

        await docxLib.renderAsync(fileBlob, this.container, null, {
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

        if (this.options.unwrapDocx) {
            this.unwrapDocxContent();
        }

        const displayName = name || ((fileBlob as File).name ? (fileBlob as File).name : '');
        this._currentFileName = displayName;
    }

    /**
     * Export the current editor content as a .docx Blob (no download triggered).
     * @param filename - Output filename (used internally for metadata)
     * @returns The generated docx Blob
     */
    async exportDocx(filename?: string): Promise<Blob> {
        const exportName = filename || this._currentFileName || 'document.docx';
        return exportToDocx(this.container, exportName.endsWith('.docx') ? exportName : exportName + '.docx');
    }

    /**
     * Export the current editor content as a .docx file and trigger download.
     * @param filename - Output filename
     * @returns The generated docx Blob
     */
    async exportAndDownload(filename?: string): Promise<Blob> {
        const exportName = filename || this._currentFileName || 'document.docx';
        const finalName = exportName.endsWith('.docx') ? exportName : exportName + '.docx';
        return exportAndDownload(this.container, finalName);
    }

    // ========== Content Access ==========

    /**
     * Get the current HTML content of the editor.
     */
    getContent(): string {
        return this.container.innerHTML;
    }

    /**
     * Set the editor's HTML content.
     * @param html - HTML string to set
     */
    setContent(html: string): void {
        this.container.innerHTML = html || '';
    }

    /**
     * Get the current file name.
     */
    getFileName(): string {
        return this._currentFileName;
    }

    /**
     * Set the current file name.
     */
    setFileName(name: string): void {
        this._currentFileName = name || '';
    }

    /**
     * Get the formatting state at the current cursor position.
     */
    getState(): EditorState {
        const state: EditorState = {
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
            state.foreColor = DocxEditor.rgbToHex(document.queryCommandValue('foreColor'));
        } catch (e) {
            // queryCommandState may throw in some browsers
        }
        return state;
    }
}
