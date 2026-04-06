/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */

export type HElement = {
    ns?: string;
    tagName: string;
    className?: string;
    style?: Record<string, string> | string;
    children?: (HElement | Node | string)[];
} & Record<string, any>;

export interface Options {
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
    renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
    useBase64URL: boolean;
    renderChanges: boolean;
    renderComments: boolean;
    renderAltChunks: boolean;
    h: (elemOrText: HElement | Node | string) => Node; //experimental, subject to change
}

//stub
export type WordDocument = any;
export declare const defaultOptions: Options;
export declare function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<WordDocument>;
export declare function renderDocument(document: WordDocument, userOptions?: Partial<Options>): Promise<Node[]>;
export declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;

// ========== Editor API ==========

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

export declare class DocxEditor {
    constructor(container: HTMLElement, options?: EditorOptions);

    // Editing control
    enableEditing(): void;
    disableEditing(): void;
    isEditable(): boolean;
    getEditorElement(): HTMLElement;

    // Formatting
    bold(): void;
    italic(): void;
    underline(): void;
    strikeThrough(): void;
    setFontSize(size: string): void;
    setFontColor(color: string): void;
    setHeading(level: string): void;
    alignLeft(): void;
    alignCenter(): void;
    alignRight(): void;
    alignJustify(): void;
    orderedList(): void;
    unorderedList(): void;
    indent(): void;
    outdent(): void;
    undo(): void;
    redo(): void;
    removeFormat(): void;

    // Document operations
    newDocument(): void;
    openFile(fileBlob: Blob | File | ArrayBuffer | Uint8Array, name?: string): Promise<void>;
    exportDocx(filename?: string): Promise<Blob>;
    exportAndDownload(filename?: string): Promise<Blob>;
    unwrapDocxContent(): void;

    // Content access
    getContent(): string;
    setContent(html: string): void;
    getFileName(): string;
    setFileName(name: string): void;
    getState(): EditorState;
}

export declare function exportToDocx(editorElement: HTMLElement, filename?: string): Promise<Blob>;
export declare function exportAndDownload(editorElement: HTMLElement, filename?: string): Promise<Blob>;
export declare function createEditor(container: HTMLElement, options?: EditorOptions): DocxEditor;
