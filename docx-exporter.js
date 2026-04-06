/**
 * docx-exporter.js
 * Converts HTML content from the editor into a valid .docx file using JSZip.
 * Builds the OOXML package manually (document.xml, styles.xml, rels, content_types).
 */

const DocxExporter = (function () {
    'use strict';

    // XML escaping
    function escXml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    // Convert CSS color (#rrggbb / rgb() / named) to OOXML hex (RRGGBB)
    function colorToHex(color) {
        if (!color || color === 'inherit' || color === 'transparent' || color === 'initial') return null;
        // Already hex
        if (/^#([0-9a-f]{6})$/i.test(color)) return color.substring(1).toUpperCase();
        if (/^#([0-9a-f]{3})$/i.test(color)) {
            const c = color.substring(1);
            return (c[0]+c[0]+c[1]+c[1]+c[2]+c[2]).toUpperCase();
        }
        // rgb(r,g,b)
        const m = color.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
        if (m) {
            const hex = [m[1], m[2], m[3]].map(x => parseInt(x).toString(16).padStart(2, '0')).join('');
            return hex.toUpperCase();
        }
        // Named colors - common ones
        const named = {
            'black':'000000','white':'FFFFFF','red':'FF0000','green':'008000',
            'blue':'0000FF','yellow':'FFFF00','purple':'800080','orange':'FFA500',
            'gray':'808080','grey':'808080','pink':'FFC0CB','brown':'A52A2A',
        };
        if (named[color.toLowerCase()]) return named[color.toLowerCase()];
        return null;
    }

    // Convert pt/px/em to half-points (for w:sz)
    function fontSizeToHalfPt(size) {
        if (!size) return null;
        const s = String(size).trim().toLowerCase();
        let pt = null;
        if (s.endsWith('pt')) pt = parseFloat(s);
        else if (s.endsWith('px')) pt = parseFloat(s) * 0.75;
        else if (s.endsWith('em')) pt = parseFloat(s) * 12;
        else if (s.endsWith('rem')) pt = parseFloat(s) * 12;
        else if (/^\d+(\.\d+)?$/.test(s)) pt = parseFloat(s); // assume pt
        if (pt != null && !isNaN(pt)) return Math.round(pt * 2);
        return null;
    }

    // Map of execCommand font sizes (1-7) to approximate pt
    const fontSizeCommandMap = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 };

    // Detect alignment from element style or attribute
    function getAlignment(el) {
        const ta = el.style?.textAlign || el.getAttribute?.('align') || '';
        switch (ta) {
            case 'left': return 'left';
            case 'center': return 'center';
            case 'right': return 'right';
            case 'justify': return 'both';
        }
        return null;
    }

    // Get computed style helper
    function getStyle(el, prop) {
        if (!el || el.nodeType !== 1) return '';
        return el.style?.[prop] || '';
    }

    // Check if element or ancestors have a style
    function hasStyle(el, prop, value) {
        const v = getStyle(el, prop);
        return v && v.indexOf(value) !== -1;
    }

    // Heading tag to style id mapping
    const headingMap = {
        'H1': 'Heading1', 'H2': 'Heading2', 'H3': 'Heading3',
        'H4': 'Heading4', 'H5': 'Heading5', 'H6': 'Heading6',
    };

    // ========== Run Properties ==========
    function buildRunProperties(el) {
        let rPr = '';
        if (!el || el.nodeType !== 1) return rPr;

        const cs = window.getComputedStyle(el);

        // Bold
        if (cs.fontWeight === 'bold' || parseInt(cs.fontWeight) >= 700) {
            rPr += '<w:b/>';
        }
        // Italic
        if (cs.fontStyle === 'italic') {
            rPr += '<w:i/>';
        }
        // Underline
        if (cs.textDecorationLine?.includes('underline') || cs.textDecoration?.includes('underline')) {
            rPr += '<w:u w:val="single"/>';
        }
        // Strikethrough
        if (cs.textDecorationLine?.includes('line-through') || cs.textDecoration?.includes('line-through')) {
            rPr += '<w:strike/>';
        }
        // Font size
        const fontSize = cs.fontSize;
        const hp = fontSizeToHalfPt(fontSize);
        if (hp && hp !== 22) { // 22 half-pt = 11pt default
            rPr += `<w:sz w:val="${hp}"/>`;
            rPr += `<w:szCs w:val="${hp}"/>`;
        }
        // Color
        const color = colorToHex(cs.color);
        if (color && color !== '000000') {
            rPr += `<w:color w:val="${color}"/>`;
        }
        // Font family
        const ff = cs.fontFamily;
        if (ff) {
            const fontName = ff.split(',')[0].trim().replace(/['"]/g, '');
            if (fontName && fontName !== 'Calibri' && fontName !== 'Arial') {
                rPr += `<w:rFonts w:ascii="${escXml(fontName)}" w:hAnsi="${escXml(fontName)}"/>`;
            }
        }

        return rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
    }

    // ========== Paragraph Properties ==========
    function buildParagraphProperties(el, listInfo) {
        let pPr = '';

        // Heading style
        const tag = el.tagName?.toUpperCase();
        if (headingMap[tag]) {
            pPr += `<w:pStyle w:val="${headingMap[tag]}"/>`;
        }

        // Alignment
        const align = getAlignment(el);
        if (align) {
            pPr += `<w:jc w:val="${align}"/>`;
        }

        // List numbering
        if (listInfo) {
            pPr += `<w:numPr><w:ilvl w:val="${listInfo.level}"/><w:numId w:val="${listInfo.numId}"/></w:numPr>`;
        }

        return pPr ? `<w:pPr>${pPr}</w:pPr>` : '';
    }

    // ========== Process Inline Nodes into Runs ==========
    function processInlineNodes(node) {
        let runs = '';
        if (!node) return runs;

        for (const child of node.childNodes) {
            if (child.nodeType === 3) { // Text node
                const text = child.textContent;
                if (!text) continue;
                // Get run props from parent element
                const rPr = buildRunProperties(child.parentElement);
                runs += `<w:r>${rPr}<w:t xml:space="preserve">${escXml(text)}</w:t></w:r>`;
            } else if (child.nodeType === 1) { // Element
                const tag = child.tagName.toUpperCase();
                // Skip non-content elements
                if (['STYLE', 'SCRIPT', 'LINK', 'META', 'NOSCRIPT'].includes(tag)) continue;
                if (tag === 'BR') {
                    runs += `<w:r><w:br/></w:r>`;
                } else if (tag === 'IMG') {
                    // Skip images for now (complex)
                    continue;
                } else if (['SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'DEL',
                             'SUB', 'SUP', 'A', 'FONT', 'MARK'].includes(tag)) {
                    // Inline elements - recurse but the rPr will pick up computed styles
                    runs += processInlineNodes(child);
                } else {
                    // Other elements treated as inline
                    runs += processInlineNodes(child);
                }
            }
        }
        return runs;
    }

    // ========== Process Block Elements ==========
    let numberingDefs = [];
    let nextNumId = 1;

    function processBlockElement(el, listInfo) {
        let xml = '';
        if (!el || el.nodeType !== 1) return xml;

        const tag = el.tagName.toUpperCase();

        // List containers
        if (tag === 'UL' || tag === 'OL') {
            const numId = nextNumId++;
            const isOrdered = tag === 'OL';
            numberingDefs.push({ numId, isOrdered, level: (listInfo?.level ?? -1) + 1 });

            for (const li of el.children) {
                if (li.tagName?.toUpperCase() === 'LI') {
                    xml += processListItem(li, { numId, level: (listInfo?.level ?? -1) + 1, isOrdered });
                }
            }
            return xml;
        }

        // Table
        if (tag === 'TABLE') {
            xml += processTable(el);
            return xml;
        }

        // Block elements that become paragraphs
        if (['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE', 'ADDRESS'].includes(tag)) {
            const pPr = buildParagraphProperties(el, listInfo);
            const runs = processInlineNodes(el);
            xml += `<w:p>${pPr}${runs}</w:p>`;
            return xml;
        }

        // Section / Article - just recurse children
        if (['SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'NAV', 'MAIN', 'ASIDE'].includes(tag)) {
            for (const child of el.children) {
                xml += processBlockElement(child, listInfo);
            }
            return xml;
        }

        // Fallback: treat as paragraph
        if (el.childNodes.length > 0) {
            const hasBlockChildren = Array.from(el.children).some(c => isBlockElement(c));
            if (hasBlockChildren) {
                for (const child of el.childNodes) {
                    if (child.nodeType === 1 && isBlockElement(child)) {
                        xml += processBlockElement(child, listInfo);
                    } else if (child.nodeType === 1 || (child.nodeType === 3 && child.textContent.trim())) {
                        const pPr = buildParagraphProperties(el, listInfo);
                        const tmpDiv = document.createElement('div');
                        tmpDiv.appendChild(child.cloneNode(true));
                        const runs = processInlineNodes(tmpDiv);
                        if (runs) xml += `<w:p>${pPr}${runs}</w:p>`;
                    }
                }
            } else {
                const pPr = buildParagraphProperties(el, listInfo);
                const runs = processInlineNodes(el);
                xml += `<w:p>${pPr}${runs}</w:p>`;
            }
        }

        return xml;
    }

    function isBlockElement(el) {
        if (!el || el.nodeType !== 1) return false;
        const blockTags = ['P','DIV','H1','H2','H3','H4','H5','H6','UL','OL','LI',
                           'TABLE','BLOCKQUOTE','PRE','SECTION','ARTICLE','HEADER',
                           'FOOTER','NAV','MAIN','ASIDE','ADDRESS','HR','FIGURE'];
        return blockTags.includes(el.tagName.toUpperCase());
    }

    function processListItem(li, listInfo) {
        let xml = '';
        const hasBlockChildren = Array.from(li.children).some(c =>
            isBlockElement(c) && !['UL','OL'].includes(c.tagName.toUpperCase()));

        if (hasBlockChildren) {
            let isFirst = true;
            for (const child of li.childNodes) {
                if (child.nodeType === 1 && isBlockElement(child)) {
                    if (['UL','OL'].includes(child.tagName.toUpperCase())) {
                        xml += processBlockElement(child, listInfo);
                    } else {
                        xml += processBlockElement(child, isFirst ? listInfo : null);
                        isFirst = false;
                    }
                } else if (child.nodeType === 3 && child.textContent.trim()) {
                    const pPr = buildParagraphProperties(li, isFirst ? listInfo : null);
                    xml += `<w:p>${pPr}<w:r><w:t xml:space="preserve">${escXml(child.textContent)}</w:t></w:r></w:p>`;
                    isFirst = false;
                }
            }
        } else {
            const pPr = buildParagraphProperties(li, listInfo);
            const runs = processInlineNodes(li);
            xml += `<w:p>${pPr}${runs}</w:p>`;
        }

        // Nested lists
        for (const child of li.children) {
            if (['UL','OL'].includes(child.tagName?.toUpperCase())) {
                xml += processBlockElement(child, listInfo);
            }
        }

        return xml;
    }

    function processTable(table) {
        let xml = '<w:tbl>';
        xml += '<w:tblPr><w:tblBorders>';
        xml += '<w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>';
        xml += '<w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>';
        xml += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>';
        xml += '<w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>';
        xml += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/>';
        xml += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/>';
        xml += '</w:tblBorders></w:tblPr>';

        const rows = table.querySelectorAll('tr');
        for (const row of rows) {
            xml += '<w:tr>';
            const cells = row.querySelectorAll('td, th');
            for (const cell of cells) {
                xml += '<w:tc>';
                xml += '<w:tcPr>';
                if (cell.colSpan > 1) {
                    xml += `<w:gridSpan w:val="${cell.colSpan}"/>`;
                }
                xml += '</w:tcPr>';
                // Cell content
                const hasBlocks = Array.from(cell.children).some(c => isBlockElement(c));
                if (hasBlocks) {
                    for (const child of cell.childNodes) {
                        if (child.nodeType === 1 && isBlockElement(child)) {
                            xml += processBlockElement(child);
                        } else if (child.nodeType === 3 && child.textContent.trim()) {
                            xml += `<w:p><w:r><w:t xml:space="preserve">${escXml(child.textContent)}</w:t></w:r></w:p>`;
                        }
                    }
                } else {
                    const runs = processInlineNodes(cell);
                    xml += `<w:p>${runs}</w:p>`;
                }
                xml += '</w:tc>';
            }
            xml += '</w:tr>';
        }

        xml += '</w:tbl>';
        return xml;
    }

    // ========== Build Document XML ==========
    function buildDocumentXml(editorEl) {
        numberingDefs = [];
        nextNumId = 1;

        let bodyXml = '';

        // Process children of editor
        const children = editorEl.childNodes;
        for (const child of children) {
            if (child.nodeType === 1) {
                // Skip style/script/comment elements
                const skipTags = ['STYLE', 'SCRIPT', 'LINK', 'META', 'NOSCRIPT'];
                if (skipTags.includes(child.tagName?.toUpperCase())) continue;

                // Check for docx-wrapper (rendered from docx-preview)
                if (child.classList?.contains('docx-wrapper')) {
                    // Process all sections inside wrapper
                    for (const section of child.children) {
                        for (const article of section.children) {
                            for (const block of article.childNodes) {
                                if (block.nodeType === 1) {
                                    bodyXml += processBlockElement(block);
                                }
                            }
                        }
                    }
                } else {
                    bodyXml += processBlockElement(child);
                }
            } else if (child.nodeType === 3 && child.textContent.trim()) {
                bodyXml += `<w:p><w:r><w:t xml:space="preserve">${escXml(child.textContent)}</w:t></w:r></w:p>`;
            }
        }

        // Ensure at least one paragraph
        if (!bodyXml) {
            bodyXml = '<w:p/>';
        }

        // Section properties (A4 / Letter)
        const sectPr = `<w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>`;

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    mc:Ignorable="w14 wp14">
    <w:body>
        ${bodyXml}
        ${sectPr}
    </w:body>
</w:document>`;
    }

    // ========== Build Styles XML ==========
    function buildStylesXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="w14">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun" w:cs="Times New Roman"/>
                <w:sz w:val="22"/>
                <w:szCs w:val="22"/>
                <w:lang w:val="en-US" w:eastAsia="zh-CN"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault>
            <w:pPr>
                <w:spacing w:after="0" w:line="276" w:lineRule="auto"/>
            </w:pPr>
        </w:pPrDefault>
    </w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:qFormat/>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="0"/></w:pPr>
        <w:rPr><w:b/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr>
        <w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading3">
        <w:name w:val="heading 3"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr>
        <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading4">
        <w:name w:val="heading 4"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="3"/></w:pPr>
        <w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading5">
        <w:name w:val="heading 5"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="4"/></w:pPr>
        <w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading6">
        <w:name w:val="heading 6"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:qFormat/>
        <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="5"/></w:pPr>
        <w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="ListParagraph">
        <w:name w:val="List Paragraph"/>
        <w:basedOn w:val="Normal"/>
        <w:qFormat/>
        <w:pPr><w:ind w:left="720"/></w:pPr>
    </w:style>
</w:styles>`;
    }

    // ========== Build Numbering XML ==========
    function buildNumberingXml() {
        if (numberingDefs.length === 0) return null;

        let abstractNums = '';
        let nums = '';
        const seen = new Map();

        for (const def of numberingDefs) {
            const key = `${def.isOrdered}-${def.level}`;
            if (seen.has(key)) continue;
            seen.set(key, def.numId);

            const abstractId = def.numId;
            const fmt = def.isOrdered ? 'decimal' : 'bullet';
            const text = def.isOrdered ? `%${def.level + 1}.` : '\uF0B7';
            const font = def.isOrdered ? '' : '<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>';

            abstractNums += `
            <w:abstractNum w:abstractNumId="${abstractId}">
                <w:multiLevelType w:val="hybridMultilevel"/>
                <w:lvl w:ilvl="${def.level}">
                    <w:start w:val="1"/>
                    <w:numFmt w:val="${fmt}"/>
                    <w:lvlText w:val="${escXml(text)}"/>
                    <w:lvlJc w:val="left"/>
                    <w:pPr><w:ind w:left="${720 * (def.level + 1)}" w:hanging="360"/></w:pPr>
                    ${font ? `<w:rPr>${font}</w:rPr>` : ''}
                </w:lvl>
            </w:abstractNum>`;

            nums += `<w:num w:numId="${def.numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`;
        }

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ${abstractNums}
    ${nums}
</w:numbering>`;
    }

    // ========== Content Types ==========
    function buildContentTypes(hasNumbering) {
        let types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`;

        if (hasNumbering) {
            types += `\n    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`;
        }

        types += `\n</Types>`;
        return types;
    }

    // ========== Relationships ==========
    function buildRootRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
    }

    function buildDocumentRels(hasNumbering) {
        let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;

        if (hasNumbering) {
            rels += `\n    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;
        }

        rels += `\n</Relationships>`;
        return rels;
    }

    // ========== Main Export Function ==========
    async function exportToDocx(editorElement, filename) {
        if (typeof JSZip === 'undefined') {
            throw new Error('JSZip is required for export. Please include jszip library.');
        }

        filename = filename || 'document.docx';

        // Build document XML (this also populates numberingDefs)
        const documentXml = buildDocumentXml(editorElement);
        const stylesXml = buildStylesXml();
        const numberingXml = buildNumberingXml();
        const hasNumbering = numberingXml !== null;

        // Create ZIP
        const zip = new JSZip();

        // [Content_Types].xml
        zip.file('[Content_Types].xml', buildContentTypes(hasNumbering));

        // _rels/.rels
        zip.file('_rels/.rels', buildRootRels());

        // word/document.xml
        zip.file('word/document.xml', documentXml);

        // word/styles.xml
        zip.file('word/styles.xml', stylesXml);

        // word/_rels/document.xml.rels
        zip.file('word/_rels/document.xml.rels', buildDocumentRels(hasNumbering));

        // word/numbering.xml
        if (hasNumbering) {
            zip.file('word/numbering.xml', numberingXml);
        }

        // Generate blob
        const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

        // Download
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        return blob;
    }

    return { exportToDocx };
})();
