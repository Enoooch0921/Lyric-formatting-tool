(function (root, factory) {
    const api = factory(root);
    if (typeof module === 'object' && module.exports) module.exports = api;
    root.SlideSections = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
    const formatter = () => root.LyricFormatter || require('./lyric-formatter.js');
    const importer = () => root.PptxLyricImporter || require('./pptx-importer.js');
    const exactKey = lines => lines.map(line => line.normalize('NFKC').trim()).join('\n');
    const comparison = text => text.normalize('NFKD').toLowerCase().replace(/[\p{M}\p{P}\p{S}\s]+/gu, '');

    function parseLabel(line) {
        const parsed = formatter().parseSectionLine(line);
        if (parsed) return parsed;
        const text = line.normalize('NFKC').trim();
        let match = text.match(/^主歌\s*([1-8一二三四五六七八])(?=\s|[:.、]|$)\s*[:.、]?\s*(.*)$/);
        if (match) {
            const number = Number(match[1]) || '一二三四五六七八'.indexOf(match[1]) + 1;
            return { label: `[${number}]`, content: match[2] };
        }
        match = text.match(/^(前副歌|副歌前段|副歌|橋段)(?=\s|[:.]|$)\s*[:.]?\s*(.*)$/);
        if (!match) return null;
        return { label: match[1] === '副歌' ? '[chorus]' : match[1] === '橋段' ? '[bridge]' : '[prechorus]', content: match[2] };
    }

    function pagesFromSlides(slides) {
        const pages = [];
        slides.forEach(slide => {
            const lines = slide.lines.map(line => line.trim()).filter(Boolean);
            if (!lines.length) return;
            const previous = pages[pages.length - 1];
            if (previous && exactKey(previous.lines) === exactKey(lines)) {
                previous.slideIndices.push(slide.index);
            } else {
                pages.push({ lines, slideIndices: [slide.index] });
            }
        });
        return pages;
    }

    // Remove repeated whole-page cycles only within an already established stanza.
    function collapseCycles(pages) {
        const result = [];
        for (let start = 0; start < pages.length;) {
            let width = 1;
            let copies = 1;
            // Long unusual cycles are kept instead of doing unbounded quadratic work.
            for (let candidate = 1; candidate <= Math.min(32, (pages.length - start) / 2); candidate++) {
                const same = offset => {
                    for (let index = 0; index < candidate; index++) {
                        if (exactKey(pages[start + index].lines) !== exactKey(pages[offset + index].lines)) return false;
                    }
                    return true;
                };
                if (!same(start + candidate)) continue;
                width = candidate;
                while (start + (copies + 1) * width <= pages.length && same(start + copies * width)) copies++;
                break;
            }
            result.push(...pages.slice(start, start + width));
            start += width * copies;
        }
        return result;
    }

    function section(label, pages) {
        return { label, lines: collapseCycles(pages).flatMap(page => page.lines),
            slideIndices: pages.flatMap(page => page.slideIndices), repeatCount: 1 };
    }

    function resultFromSections(sections, slides, mode, warnings = []) {
        const unique = [];
        sections.filter(item => item.lines.length).forEach(item => {
            const duplicate = unique.find(previous => previous.label === item.label && exactKey(previous.lines) === exactKey(item.lines));
            if (duplicate) {
                duplicate.repeatCount++;
                duplicate.slideIndices.push(...item.slideIndices);
            } else unique.push({ ...item, lines: [...item.lines], slideIndices: [...item.slideIndices] });
        });
        return { sections: unique, text: unique.map(item => [item.label, ...item.lines].join('\n')).join('\n\n'),
            sourceSlideCount: slides.length, mode, warnings, missingSections: [] };
    }

    function explicitSections(pages) {
        if (!pages.some(page => page.lines.some(line => parseLabel(line)))) return null;
        const sections = [];
        let current;
        pages.forEach(page => {
            page.lines.forEach(line => {
                const tag = parseLabel(line);
                if (tag && (!current || current.label !== tag.label)) {
                    current = { label: tag.label, pages: [] };
                    sections.push(current);
                }
                const content = tag ? tag.content : line;
                if (!content) return;
                if (!current) {
                    current = { label: '[1]', pages: [] };
                    sections.push(current);
                }
                const last = current.pages[current.pages.length - 1];
                if (last && last.slideIndices === page.slideIndices) last.lines.push(content);
                else current.pages.push({ lines: [content], slideIndices: page.slideIndices });
            });
        });
        return sections.map(item => section(item.label, item.pages));
    }

    function parseReference(text) {
        if (text.length > 20000) throw new Error('參考歌詞最多 20,000 字，請只貼入這一首歌。');
        const sections = [];
        let current;
        text.replace(/\r\n?/g, '\n').split('\n').forEach(line => {
            const trimmed = line.trim();
            if (!trimmed) return;
            if (/^(?:\d+|(?:Verse|V)\s*\d+)\s*[.、:]?$/i.test(trimmed)
                && !parseLabel(trimmed.replace(/、$/, '.'))) {
                throw new Error('參考主歌編號只支援 1 到 8。');
            }
            const tag = parseLabel(trimmed.replace(/^([1-8])、/, '$1.'));
            if (tag) {
                current = { label: tag.label, lines: [] };
                sections.push(current);
                if (tag.content) current.lines.push(tag.content);
            } else if (current) current.lines.push(trimmed);
            // A title or language heading before the first label is not lyric content.
        });
        const unique = [];
        sections.filter(item => item.lines.length).forEach(item => {
            const previous = unique.find(candidate => candidate.label === item.label);
            if (previous && comparison(previous.lines.join('')) !== comparison(item.lines.join(''))) {
                throw new Error(`參考歌詞 ${item.label} 有不同內容，請使用不同段落標記。`);
            }
            if (!previous) unique.push(item);
        });
        if (!unique.length) throw new Error('請在參考歌詞加入 1.、2. 或 Verse／Chorus 等段落標記，再比對。');
        if (unique.length > 16) throw new Error('參考歌詞段落過多，請只貼入這一首歌。');
        return unique;
    }

    function pairs(text) {
        const counts = new Map();
        for (let index = 0; index < text.length - 1; index++) {
            const token = text.slice(index, index + 2);
            counts.set(token, (counts.get(token) || 0) + 1);
        }
        return counts;
    }
    function coverage(source, targetPairs) {
        let total = 0, covered = 0;
        pairs(source).forEach((count, token) => {
            total += count;
            covered += Math.min(count, targetPairs.get(token) || 0);
        });
        return total ? covered / total : 0;
    }

    function matchReference(pages, references) {
        const normalized = references.map(item => comparison(item.lines.join('')));
        const tokens = normalized.map(pairs);
        return pages.map(page => {
            const content = comparison(page.lines.map(line => {
                const tag = parseLabel(line);
                return tag ? tag.content : line;
            }).join(''));
            const scores = normalized.map((reference, index) => ({ index,
                score: content && reference.includes(content) ? 1 : coverage(content, tokens[index]) }));
            const candidates = scores.filter(item => item.score >= 0.85).sort((left, right) => right.score - left.score);
            const confident = content.length >= 8 && candidates.length
                && (candidates.length === 1 || candidates[0].score - candidates[1].score >= 0.1);
            return { index: confident ? candidates[0].index : -1, candidates: candidates.map(item => item.index) };
        });
    }

    function consolidateSlides(slides, options = {}) {
        const pages = pagesFromSlides(slides);
        if (!pages.length) throw new Error('這首歌沒有可匯入的歌詞。');
        const explicit = explicitSections(pages);
        let result;
        if (explicit) {
            result = resultFromSections(explicit, slides, 'explicit', ['已有明確段落標記，優先依標記合併分頁，不套用固定頁數。']);
        } else if (options.pagesPerSection && options.pagesPerSection !== 'auto') {
            const size = Number(options.pagesPerSection);
            if (!Number.isInteger(size) || size < 1 || size > 4) throw new Error('每段頁數請選擇 1 到 4。');
            const sections = [];
            for (let start = 0; start < pages.length; start += size) {
                const label = `[${Math.min(sections.length + 1, 8)}]`;
                const chunk = section(label, pages.slice(start, start + size));
                const duplicate = sections.find(item => exactKey(item.lines) === exactKey(chunk.lines));
                if (duplicate) {
                    duplicate.slideIndices.push(...chunk.slideIndices);
                    duplicate.repeatCount++;
                } else if (sections.length >= 8) {
                    sections[7].lines.push(...chunk.lines);
                    sections[7].slideIndices.push(...chunk.slideIndices);
                } else sections.push(chunk);
            }
            result = resultFromSections(sections, slides, 'pages', [`每 ${size} 個歌詞頁合成一段；連續相同頁先去重，不足頁數的末段仍保留。請核對分段。`]);
        } else {
            // Repetition counts matter to inference; do not remove performance copies first.
            const inferred = importer().consolidateSong(slides);
            if (inferred.sections.some(item => /chorus|bridge/.test(item.label))) {
                result = { ...inferred, mode: 'auto', missingSections: [],
                    warnings: ['依重複演唱順序推測段落，請核對；投影片分頁不一定是段落。'] };
            } else {
                result = resultFromSections([section('[1]', pages)], slides, 'auto',
                    ['沒有可靠段落界線，暫保留為一段。可選每幾頁合併，或貼入有編號的參考歌詞。']);
            }
        }
        if (!(options.referenceText || '').trim()) return result;

        const references = parseReference(options.referenceText);
        const matches = matchReference(pages, references);
        const unmatched = matches.filter(item => item.index < 0).length;
        const possible = new Set(matches.flatMap(item => item.candidates));
        const missing = references.filter((_, index) => !possible.has(index)).map(item => item.label);
        const warnings = [];
        if (missing.length) warnings.push(`參考歌詞 ${missing.join('、')} 未在投影片找到；未自動補入。`);
        if (unmatched) warnings.push(`有 ${unmatched} 個歌詞頁無法可靠對應參考段落；保留目前分段，請核對。`);
        references.forEach((reference, index) => {
            const matched = pages.filter((_, pageIndex) => matches[pageIndex].index === index);
            if (!matched.length) return;
            const actual = comparison(collapseCycles(matched).flatMap(page => page.lines)
                .map(line => { const tag = parseLabel(line); return tag ? tag.content : line; }).join(''));
            const expected = comparison(reference.lines.join(''));
            if (actual !== expected) {
                const complete = coverage(expected, pairs(actual)) >= 0.85 && coverage(actual, pairs(expected)) >= 0.85;
                warnings.push(`參考 ${reference.label} ${complete ? '有字詞差異' : '可能缺少部分歌詞或有內容差異'}；保留投影片原文，請核對。`);
            }
        });
        if (!explicit && !unmatched) {
            const sections = [];
            const firstAppearance = [...new Set(matches.map(item => item.index))];
            firstAppearance.forEach(index => {
                sections.push(section(references[index].label, pages.filter((_, pageIndex) => matches[pageIndex].index === index)));
            });
            result = resultFromSections(sections, slides, 'reference');
            if (firstAppearance.some((value, index) => index && value < firstAppearance[index - 1])) {
                warnings.push('投影片段落順序與參考不同，已保留投影片首次出現的順序。');
            }
            warnings.unshift('依參考歌詞對應段落；歌詞保留投影片原文，沒有自動補字。');
        } else if (explicit) {
            warnings.unshift('參考歌詞僅用於比對；投影片已有的段落標記保持不變。');
            result.sections.forEach(item => {
                const related = new Set(pages.flatMap((page, pageIndex) => (
                    matches[pageIndex].index >= 0 && page.slideIndices.some(index => item.slideIndices.includes(index))
                        ? [references[matches[pageIndex].index].label] : []
                )));
                if (related.size === 1 && !related.has(item.label)) {
                    warnings.push(`投影片 ${item.label} 內容較接近參考 ${[...related][0]}；保留投影片標記，請核對。`);
                }
            });
        }
        return { ...result, missingSections: missing, warnings: [...result.warnings, ...warnings] };
    }

    return { parseLabel, parseReference, consolidateSlides };
});
