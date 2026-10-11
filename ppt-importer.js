(function (root, factory) {
    const api = factory(root);
    if (typeof module === 'object' && module.exports) module.exports = api;
    root.PptLyricImporter = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
    const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    const MAX_BYTES = 50 * 1024 * 1024;

    function isLegacyPpt(bytes) {
        return SIGNATURE.every((value, index) => bytes[index] === value);
    }

    function parsePptBytes(input, explicitCfb) {
        const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
        if (bytes.byteLength > MAX_BYTES) throw new Error('PPT 超過 50 MB，請拆分檔案後再試。');
        if (!isLegacyPpt(bytes)) throw new Error('檔案不是有效的舊版 PPT。請用 PowerPoint 另存為 PPTX 後再試。');
        const cfbApi = explicitCfb || root.CFB || (typeof require === 'function' && require('./vendor/cfb-1.2.2.min.js'));
        const slideApi = root.PptxLyricImporter || require('./pptx-importer.js');
        let cfb;
        try { cfb = cfbApi.read(bytes, { type: 'array' }); }
        catch (_error) { throw new Error('無法讀取 PPT，檔案可能損壞。請重新另存為 PPTX。'); }
        const encrypted = cfbApi.find(cfb, 'EncryptedPackage') || cfbApi.find(cfb, 'EncryptionInfo');
        if (encrypted) throw new Error('這份檔案有密碼保護，請解除保護後再匯入。');
        function stream(name) {
            const entry = cfbApi.find(cfb, name);
            if (!entry?.content) throw new Error('缺少 PPT 投影片資料，請另存為 PPTX 後再試。');
            return Uint8Array.from(entry.content);
        }
        const document = stream('PowerPoint Document');
        const currentUser = stream('Current User');
        const view = new DataView(document.buffer, document.byteOffset, document.byteLength);
        const user = new DataView(currentUser.buffer, currentUser.byteOffset, currentUser.byteLength);
        const malformed = () => new Error('PPT 結構損壞或使用尚未支援的格式，請另存為 PPTX 後再試。');
        if (currentUser.length < 20 || user.getUint16(2, true) !== 4086) throw malformed();
        if (user.getUint32(12, true) !== 0xe391c05f) {
            throw new Error('這份 PPT 有密碼保護或使用不支援的格式，請解除保護並另存為 PPTX。');
        }

        let recordCount = 0;
        function record(offset, limit = document.length) {
            if (++recordCount > 200000 || offset < 0 || offset + 8 > limit) throw malformed();
            const options = view.getUint16(offset, true);
            const end = offset + 8 + view.getUint32(offset + 4, true);
            if (end > limit || end > document.length) throw malformed();
            return { offset, start: offset + 8, end, length: end - offset - 8,
                type: view.getUint16(offset + 2, true), version: options & 15, instance: options >>> 4 };
        }
        function records(start, end) {
            const result = [];
            for (let offset = start; offset < end;) {
                const item = record(offset, end);
                result.push(item);
                offset = item.end;
            }
            return result;
        }
        function children(container) { return records(container.start, container.end); }
        function uint(item, relative) {
            if (relative + 4 > item.length) throw malformed();
            return view.getUint32(item.start + relative, true);
        }

        // Resolve only the latest objects. Old edits, masters, notes and metadata are not slide text.
        const persist = new Map();
        const seenEdits = new Set();
        let editOffset = user.getUint32(16, true);
        let documentId;
        do {
            if (seenEdits.has(editOffset) || seenEdits.size >= 10000) throw malformed();
            seenEdits.add(editOffset);
            const edit = record(editOffset);
            if (edit.type !== 4085 || edit.length < 28) throw malformed();
            if (documentId === undefined) documentId = uint(edit, 16);
            const directory = record(uint(edit, 12));
            if (![6001, 6002].includes(directory.type)) throw malformed();
            for (let offset = directory.start; offset < directory.end;) {
                if (offset + 4 > directory.end) throw malformed();
                const entry = view.getUint32(offset, true);
                const count = entry >>> 20;
                const firstId = entry & 0xfffff;
                offset += 4;
                if (!count || offset + count * 4 > directory.end) throw malformed();
                for (let index = 0; index < count; index++, offset += 4) {
                    if (!persist.has(firstId + index)) persist.set(firstId + index, view.getUint32(offset, true));
                }
            }
            editOffset = uint(edit, 8);
        } while (editOffset);
        function persisted(id, expectedType) {
            const offset = persist.get(id);
            if (offset === undefined) throw malformed();
            const result = record(offset);
            if (result.type !== expectedType) throw malformed();
            return result;
        }
        const presentation = persisted(documentId, 1000);
        const presentationRecords = children(presentation);
        const documentAtom = presentationRecords.find(item => item.type === 1001);
        const height = documentAtom && documentAtom.length >= 8 ? uint(documentAtom, 4) : 0;
        const slideList = presentationRecords.find(item => item.type === 4080 && item.instance === 0);
        if (!slideList) throw malformed();
        const orderedSlides = [];
        let currentSlide;
        let currentText;
        function decodeText(item) {
            if (item.type === 4000 && item.length % 2) throw malformed();
            // TextBytesAtom stores the low byte of each Unicode character, not a document codepage.
            const data = document.subarray(item.start, item.end);
            const text = item.type === 4000 ? new TextDecoder('utf-16le').decode(data)
                : Array.from(data, value => String.fromCharCode(value)).join('');
            return text.replace(/[\r\v\u0000]/g, '\n').split('\n')
                .map(line => line.replace(/\u00a0/g, ' ').replace(/[\t ]+/g, ' ').trim())
                .filter(Boolean);
        }
        children(slideList).forEach(item => {
            if (item.type === 1011) {
                currentSlide = { persistId: uint(item, 0), outline: [] };
                orderedSlides.push(currentSlide);
                currentText = null;
            } else if (item.type === 3999 && currentSlide) {
                currentText = { type: uint(item, 0), lines: [] };
                currentSlide.outline.push(currentText);
            } else if ([4000, 4008].includes(item.type) && currentText) {
                currentText.lines.push(...decodeText(item));
            }
        });
        if (!orderedSlides.length || orderedSlides.length > 5000) throw malformed();

        const slides = orderedSlides.map((info, index) => {
            const slide = persisted(info.persistId, 1006);
            const shapes = [];
            const usedOutline = new Set();
            function walk(container, depth = 0) {
                if (depth > 40) throw malformed();
                children(container).forEach(item => {
                    if (item.type === 0xf004) {
                        const parts = children(item);
                        const anchor = parts.find(part => part.type === 0xf010);
                        let x = 0, y = 0;
                        if (anchor?.length === 8) {
                            y = view.getInt16(anchor.start, true);
                            x = view.getInt16(anchor.start + 2, true);
                        } else if (anchor?.length === 16) {
                            x = view.getInt32(anchor.start, true);
                            y = view.getInt32(anchor.start + 4, true);
                        }
                        const client = parts.find(part => part.type === 0xf011);
                        const textbox = parts.find(part => part.type === 0xf00d);
                        const textParts = textbox ? children(textbox) : [];
                        const reference = textParts.find(part => part.type === 3998);
                        const outlineIndex = reference ? uint(reference, 0) : null;
                        if (outlineIndex !== null) usedOutline.add(outlineIndex);
                        const placeholder = client && children(client).find(part => part.type === 3011);
                        const placeholderType = placeholder?.length >= 5 ? view.getUint8(placeholder.start + 4) : 0;
                        if ([7, 8, 9, 10, 11, 12].includes(placeholderType)) return;
                        if (!textbox) return;
                        if (textParts.some(part => [4056, 4087, 4088].includes(part.type))) return;
                        let lines;
                        if (reference) {
                            const outline = info.outline[outlineIndex];
                            if (!outline) throw malformed();
                            lines = outline.lines;
                        } else {
                            lines = textParts.filter(part => [4000, 4008].includes(part.type)).flatMap(decodeText);
                        }
                        if (lines.length) shapes.push({ x, y, lines });
                    } else if (item.version === 15 || item.type === 1036) {
                        // Only descend into slide drawing containers, never arbitrary text-bearing records.
                        if ([1036, 0xf002, 0xf003].includes(item.type)) walk(item, depth + 1);
                    }
                });
            }
            walk(slide);
            // Some producers keep body text only in the slide list without a textbox reference.
            info.outline.forEach((text, outlineIndex) => {
                if (!usedOutline.has(outlineIndex) && text.type !== 2 && text.lines.length) {
                    shapes.push({ x: 0, y: 0, lines: text.lines });
                }
            });
            const cleaned = slideApi.cleanSlideShapes(shapes, height);
            return { index: index + 1, part: `ppt/slide-${info.persistId}`, ...cleaned, backgroundSignature: '' };
        }).filter(slide => slide.lines.length);
        if (!slides.length) throw new Error('PPT 裡找不到可匯入的文字。圖片或掃描歌詞目前不支援。');
        return slides;
    }

    async function parsePptFile(file) {
        if (!file || !/\.ppt$/i.test(file.name || '')) throw new Error('請選擇 .ppt 檔案。');
        if (file.size > MAX_BYTES) throw new Error('PPT 超過 50 MB，請拆分檔案後再試。');
        const bytes = new Uint8Array(await file.arrayBuffer());
        // Some files have a .ppt suffix but contain the newer ZIP-based PPTX format.
        if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
            const api = root.PptxLyricImporter || require('./pptx-importer.js');
            return api.parsePptxBytes(bytes);
        }
        return parsePptBytes(bytes);
    }

    return { isLegacyPpt, parsePptBytes, parsePptFile };
});
