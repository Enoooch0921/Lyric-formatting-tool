(function (root, factory) {
    const api = factory(root);
    if (typeof module === 'object' && module.exports) module.exports = api;
    root.DocxLyricImporter = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
    const MAX_FILE_BYTES = 50 * 1024 * 1024;
    const MAX_XML_BYTES = 4 * 1024 * 1024;
    const WORD_NAMESPACES = new Set([
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'http://purl.oclc.org/ooxml/wordprocessingml/main'
    ]);

    function formatter() {
        return root.LyricFormatter || require('./lyric-formatter.js');
    }

    function isWord(node, name) {
        return node?.nodeType === 1 && WORD_NAMESPACES.has(node.namespaceURI)
            && node.localName === name;
    }

    function children(node) {
        return Array.from(node?.childNodes || []).filter(child => child.nodeType === 1);
    }

    function child(node, name) {
        return children(node).find(item => isWord(item, name));
    }

    function attribute(node, name) {
        if (!node) return '';
        for (const ns of WORD_NAMESPACES) {
            if (node.hasAttributeNS(ns, name)) return node.getAttributeNS(ns, name);
        }
        return '';
    }

    function enabled(node) {
        return Boolean(node) && !/^(?:0|false|off)$/i.test(attribute(node, 'val'));
    }

    function parseXml(xml, Parser) {
        if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('DOCX 含有不支援的文字定義。');
        try {
            const document = new Parser({ onError: () => { throw new Error('Invalid XML'); } })
                .parseFromString(xml, 'application/xml');
            if (!document.documentElement || document.getElementsByTagName('parsererror').length) {
                throw new Error('Invalid XML');
            }
            return document;
        } catch (_error) {
            throw new Error('DOCX 文字結構損壞，請用 Word 重新另存為 .docx。');
        }
    }

    function extractDocument(document, stylesDocument) {
        const styles = new Map();
        children(stylesDocument?.documentElement).forEach(style => {
            if (!isWord(style, 'style')) return;
            const name = attribute(child(style, 'name'), 'val');
            const outline = attribute(child(child(style, 'pPr'), 'outlineLvl'), 'val');
            styles.set(attribute(style, 'styleId'), {
                isHeading: /^(?:title|heading\s*[12]|標題\s*[12]?)$/i.test(name)
                    || (outline !== '' && Number(outline) <= 1),
                fontSize: Number(attribute(child(child(style, 'rPr'), 'sz'), 'val')) / 2
            });
        });
        const body = child(document.documentElement, 'body');
        if (!isWord(document.documentElement, 'document') || !body) {
            throw new Error('檔案不是有效的 Word DOCX 文件。');
        }

        const paragraphs = [];
        const warnings = new Set();
        function readParagraph(node) {
            let text = '';
            let fontSize = 0;
            let hasPageBreak = false;
            const properties = child(node, 'pPr');
            const styleId = attribute(child(properties, 'pStyle'), 'val');
            const style = styles.get(styleId);

            function readInline(item) {
                if (isWord(item, 'del') || isWord(item, 'moveFrom') || isWord(item, 'pPr')) return;
                if (isWord(item, 'drawing') || isWord(item, 'pict') || isWord(item, 'txbxContent')) {
                    warnings.add('圖片與文字方塊未匯入，請核對是否有遺漏歌詞。');
                    return;
                }
                if (isWord(item, 'r')) {
                    const runProperties = child(item, 'rPr');
                    if (enabled(child(runProperties, 'vanish')) || enabled(child(runProperties, 'webHidden'))) return;
                    fontSize = Math.max(fontSize, Number(attribute(child(runProperties, 'sz'), 'val')) / 2);
                }
                if (isWord(item, 't')) {
                    text += item.textContent;
                } else if (isWord(item, 'tab')) {
                    text += ' ';
                } else if (isWord(item, 'br') || isWord(item, 'cr')) {
                    text += '\n';
                    if (attribute(item, 'type') === 'page') hasPageBreak = true;
                } else if (isWord(item, 'noBreakHyphen')) {
                    text += '-';
                } else if (!isWord(item, 'rPr')) {
                    children(item).forEach(readInline);
                }
            }
            children(node).forEach(readInline);
            const lines = text.replace(/\r\n?/g, '\n').split('\n')
                .map(line => line.replace(/[\t\u00a0 ]+/g, ' ').trim());
            paragraphs.push({
                index: paragraphs.length + 1,
                lines,
                text: lines.join('\n').trim(),
                fontSize: fontSize || style?.fontSize || 0,
                isHeading: Boolean(style?.isHeading || /^(?:Title|Heading[12])$/i.test(styleId)),
                pageBreakBefore: enabled(child(properties, 'pageBreakBefore')),
                hasPageBreak
            });
        }

        function walk(node) {
            if (isWord(node, 'del') || isWord(node, 'moveFrom')) return;
            if (isWord(node, 'p')) return readParagraph(node);
            if (isWord(node, 'tbl')) {
                warnings.add('表格依照列與儲存格順序讀取，請核對雙欄或雙語歌詞的順序。');
            }
            // Word can store the same content twice for compatibility.
            if (node.localName === 'AlternateContent') {
                const alternative = children(node).find(item => item.localName === 'Choice')
                    || children(node).find(item => item.localName === 'Fallback');
                if (alternative) walk(alternative);
                return;
            }
            children(node).forEach(walk);
        }
        walk(body);
        if (!paragraphs.some(paragraph => paragraph.text)) {
            throw new Error('DOCX 裡找不到可匯入的文字。圖片或掃描歌詞目前不支援。');
        }
        return { paragraphs, warnings: [...warnings] };
    }

    function parseDocxBytes(bytes, explicitZipApi, explicitParser) {
        const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
        if (data.byteLength > MAX_FILE_BYTES) throw new Error('DOCX 超過 50 MB，請拆分檔案後再試。');
        const zip = explicitZipApi || root.fflate;
        const Parser = explicitParser || root.DOMParser;
        if (!zip || !Parser) throw new Error('缺少文件讀取元件，請重新整理頁面後再試。');
        let files;
        try {
            files = zip.unzipSync(data, {
                filter(file) {
                    if (!['word/document.xml', 'word/styles.xml'].includes(file.name)) return false;
                    if (file.originalSize > MAX_XML_BYTES) throw new Error('文字資料過大');
                    return true;
                }
            });
        } catch (_error) {
            throw new Error('無法讀取 DOCX，檔案可能損壞、受密碼保護或文字資料過大。');
        }
        if (!files['word/document.xml']) throw new Error('找不到 DOCX 的文件內容，請確認是 .docx 格式。');
        const decoder = new TextDecoder('utf-8');
        function readPart(name) {
            const bytes = files[name];
            if (!bytes) return null;
            if (bytes.byteLength > MAX_XML_BYTES) throw new Error('DOCX 文字資料過大，請拆分檔案。');
            const utf16 = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0);
            const utf16be = (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c);
            const xml = utf16 ? new TextDecoder('utf-16le').decode(bytes)
                : utf16be ? new TextDecoder('utf-16be').decode(bytes) : decoder.decode(bytes);
            return parseXml(xml, Parser);
        }
        return extractDocument(readPart('word/document.xml'), readPart('word/styles.xml'));
    }

    async function parseDocxFile(file) {
        if (!file || !/\.docx$/i.test(file.name || '')) throw new Error('請選擇 .docx 檔案；舊版 .doc 請先另存為 .docx。');
        if (file.size > MAX_FILE_BYTES) throw new Error('DOCX 超過 50 MB，請拆分檔案後再試。');
        return parseDocxBytes(await file.arrayBuffer());
    }

    function detectSongs(paragraphs) {
        const contentIndexes = paragraphs.flatMap((paragraph, index) => paragraph.text ? [index] : []);
        const titleIndexes = new Set();
        contentIndexes.forEach((index, position) => {
            const paragraph = paragraphs[index];
            if (paragraph.text.length > 80 || paragraph.text.includes('\n')
                || formatter().parseSectionLine(paragraph.text)) return;
            const nextIndex = contentIndexes[position + 1];
            if (nextIndex === undefined) return;
            const next = paragraphs[nextIndex];
            const previousIndex = contentIndexes[position - 1];
            const beforeGap = previousIndex === undefined ? index : index - previousIndex - 1;
            const afterGap = nextIndex - index - 1;
            const isLarger = paragraph.fontSize > 0 && next.fontSize > 0
                && paragraph.fontSize >= next.fontSize * 1.3;
            const titleFormatting = paragraph.isHeading || isLarger;
            const boundary = previousIndex === undefined || beforeGap >= 2 || paragraph.pageBreakBefore
                || paragraphs[previousIndex]?.hasPageBreak || titleFormatting;
            const nextLines = paragraphs.slice(nextIndex, nextIndex + 4)
                .flatMap(item => item.lines).filter(Boolean).length;
            if (boundary && (afterGap >= 1 || titleFormatting) && nextLines >= 2) titleIndexes.add(index);
        });

        const songs = [];
        let group = [];
        let title = '';
        function finish() {
            if (!group.some(paragraph => paragraph.text)) return;
            const nonempty = group.filter(paragraph => paragraph.text);
            songs.push({
                title,
                paragraphs: group,
                startParagraph: nonempty[0].index,
                endParagraph: nonempty[nonempty.length - 1].index,
                preview: nonempty.flatMap(paragraph => paragraph.lines).filter(Boolean).slice(0, 2).join('\n')
            });
        }
        paragraphs.forEach((paragraph, index) => {
            if (titleIndexes.has(index)) {
                finish();
                title = paragraph.text;
                group = [];
            } else {
                group.push(paragraph);
            }
        });
        finish();
        return songs;
    }

    function consolidateSong(paragraphs) {
        const blocks = [];
        let lines = [];
        let explicitLabel = null;
        function finish() {
            if (lines.length) blocks.push({ lines, explicitLabel });
            lines = [];
        }
        paragraphs.forEach(paragraph => paragraph.lines.forEach(line => {
            if (!line.trim()) {
                finish();
                return;
            }
            const section = formatter().parseSectionLine(line);
            if (section) {
                finish();
                explicitLabel = section.label;
                if (section.content) lines.push(section.content);
            } else {
                lines.push(line);
            }
        }));
        finish();

        // Only merge whole matching stanzas; recurring words or partial lines stay intact.
        const unique = [];
        blocks.forEach(block => {
            const key = block.lines.map(line => line.normalize('NFKC').replace(/\s+/g, ' ').trim()).join('\n');
            const existing = unique.find(item => item.key === key && item.explicitLabel === block.explicitLabel);
            if (existing) existing.repeatCount++;
            else unique.push({ ...block, key, repeatCount: 1 });
        });
        const reserved = new Set(unique.map(block => block.explicitLabel).filter(Boolean));
        let verse = 1;
        const sections = [];
        unique.forEach(block => {
            let label = block.explicitLabel;
            if (!label) {
                while (verse < 8 && reserved.has(`[${verse}]`)) verse++;
                label = `[${Math.min(verse++, 8)}]`;
            }
            const overflow = !block.explicitLabel && verse > 9;
            const verseEight = overflow && sections.find(section => section.label === '[8]' && !section.explicit);
            if (verseEight) {
                verseEight.lines.push('', ...block.lines);
                verseEight.repeatCount = Math.max(verseEight.repeatCount, block.repeatCount);
            } else {
                sections.push({ label, lines: block.lines.slice(), repeatCount: block.repeatCount, explicit: Boolean(block.explicitLabel) });
            }
        });
        return {
            text: sections.map(section => [section.label, ...section.lines].join('\n')).join('\n\n'),
            sections,
            sourceParagraphCount: paragraphs.filter(paragraph => paragraph.text).length,
            mergedBlockCount: blocks.length - unique.length
        };
    }

    return { parseDocxBytes, parseDocxFile, detectSongs, consolidateSong };
});
