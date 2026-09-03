(function (root, factory) {
    const api = factory();

    if (typeof module === 'object' && module.exports) {
        module.exports = api;
    }

    root.LyricFormatter = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    function fullWidthToHalfWidth(str) {
        return str.replace(/[\uff01-\uff5e]/g, function (ch) {
            return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
        }).replace(/\u3000/g, ' ');
    }

    function getVisualLength(str) {
        let len = 0;
        for (let i = 0; i < str.length; i++) {
            len += str.charCodeAt(i) <= 127 ? 1 : 2;
        }
        return len;
    }

    function normalizeLine(line) {
        return fullWidthToHalfWidth(line).replace(/\s+/g, ' ').trim();
    }

    function makeSection(label, content) {
        return {
            label: `[${label}]`,
            content: (content || '').trim()
        };
    }

    function parseSectionLine(line) {
        const normalized = normalizeLine(line);
        if (!normalized) return null;

        let match = normalized.match(/^\[\s*([1-8])\s*\]\s*(?:[:.]\s*)?(.*)$/);
        if (match) return makeSection(match[1], match[2]);

        match = normalized.match(/^\[?\s*(?:v|verse)\s*([1-8])\s*\]?(?=\s|[:.]|$)\s*(?:[:.]\s*)?(.*)$/i);
        if (match) return makeSection(match[1], match[2]);

        match = normalized.match(/^([1-8])\s+(.+)$/);
        if (match) return makeSection(match[1], match[2]);

        match = normalized.match(/^([1-8])\s*(?:[:.]\s*(.*))?$/);
        if (match) return makeSection(match[1], match[2]);

        match = normalized.match(/^\[?\s*(pre[\s-]?(?:chorus|chrous|c)|pro[\s-]?(?:chorus|chrous|c)|chorus|chrous|chrouse|chros|bridge|brisge|brigde|br|pre|c|b|p)\s*\]?(?=\s|[:.]|$)\s*(?:[:.]\s*)?(.*)$/i);
        if (!match) return null;

        const tag = match[1].toLowerCase().replace(/[\s-]/g, '');
        if (tag === 'c' || tag.startsWith('ch')) {
            return makeSection('chorus', match[2]);
        }
        if (tag === 'b' || tag === 'br' || tag.startsWith('bri')) {
            return makeSection('bridge', match[2]);
        }
        return makeSection('prechorus', match[2]);
    }

    function cleanLyricContent(content) {
        const apostropheTokens = {
            "'": '\ue000',
            '\u2018': '\ue001',
            '\u2019': '\ue002'
        };
        const tokenApostrophes = Object.fromEntries(
            Object.entries(apostropheTokens).map(([apostrophe, token]) => [token, apostrophe])
        );

        return fullWidthToHalfWidth(content)
            // Keep apostrophes that are part of an English word while cleaning other punctuation.
            .replace(/(\p{Script=Latin})(['\u2018\u2019])(?=\p{Script=Latin})/gu, (
                _match,
                letter,
                apostrophe
            ) => letter + apostropheTokens[apostrophe])
            // Treat all remaining punctuation and symbols as word boundaries.
            .replace(/[\p{P}\p{S}]+/gu, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .replace(/[\ue000-\ue002]/g, token => tokenApostrophes[token]);
    }

    function wrapContent(content, visualLimit) {
        const wrapped = [];
        const segments = content.split(' ');
        let currentLine = '';
        let currentLen = 0;

        segments.forEach((segment, index) => {
            if (!segment) return;

            const segmentLen = getVisualLength(segment);
            const spaceWidth = currentLine ? 1 : 0;
            const isLastWord = index === segments.length - 1;
            // A small overflow is easier to read than leaving one English word alone.
            const tolerance = isLastWord ? Math.min(8, Math.max(5, Math.ceil(visualLimit * 0.3))) : 0;

            if (currentLine && currentLen + spaceWidth + segmentLen > visualLimit + tolerance) {
                wrapped.push(currentLine);
                currentLine = segment;
                currentLen = segmentLen;
            } else {
                currentLine += (currentLine ? ' ' : '') + segment;
                currentLen += spaceWidth + segmentLen;
            }
        });

        if (currentLine) wrapped.push(currentLine);
        return wrapped;
    }

    function formatLyricsText(inputLyrics, charLimit) {
        const parsedLimit = Number.parseInt(charLimit, 10);
        const safeLimit = Number.isFinite(parsedLimit)
            ? Math.min(40, Math.max(5, parsedLimit))
            : 12;
        const visualLimit = safeLimit * 2;
        const formattedLines = [];

        inputLyrics.split('\n').forEach(rawLine => {
            const normalizedLine = normalizeLine(rawLine);
            if (!normalizedLine) return;

            const section = parseSectionLine(normalizedLine);
            const content = cleanLyricContent(section ? section.content : normalizedLine);

            if (section) formattedLines.push({ type: 'section', text: section.label });
            wrapContent(content, visualLimit).forEach(text => {
                formattedLines.push({ type: 'lyric', text });
            });
        });

        return formattedLines.map(line => (
            line.type === 'section' ? `${line.text}\n` : `${line.text}\n\n`
        )).join('').trim();
    }

    function classifyLyricLine(line) {
        if (!line.trim()) return 'blank';
        if (parseSectionLine(line)) return 'section';

        const hasChinese = /[\p{Script=Han}\u3100-\u312f\u31a0-\u31bf]/u.test(line);
        const hasEnglish = /\p{Script=Latin}/u.test(line);

        if (hasChinese && hasEnglish) return 'mixed';
        if (hasChinese) return 'chinese';
        if (hasEnglish) return 'english';
        return 'neutral';
    }

    function filterLyricsByLanguage(inputLyrics, language) {
        const removedLines = [];
        let ambiguousCount = 0;

        const keptLines = inputLyrics.split('\n').filter((line, index) => {
            const kind = classifyLyricLine(line);
            const shouldKeep = kind === 'blank'
                || kind === 'section'
                || kind === 'mixed'
                || kind === 'neutral'
                || kind === language;

            if (kind === 'mixed' || kind === 'neutral') ambiguousCount++;
            if (!shouldKeep) removedLines.push({ lineNumber: index + 1, text: line });
            return shouldKeep;
        });

        return {
            text: keptLines.join('\n'),
            removedLines,
            ambiguousCount
        };
    }

    function keepOddContentLines(inputLyrics) {
        const removedLines = [];
        let contentCounter = 0;

        const keptLines = inputLyrics.split('\n').filter((line, index) => {
            if (!line.trim() || parseSectionLine(line)) {
                contentCounter = 0;
                return true;
            }

            contentCounter++;
            const shouldKeep = contentCounter % 2 === 1;
            if (!shouldKeep) removedLines.push({ lineNumber: index + 1, text: line });
            return shouldKeep;
        });

        return {
            text: keptLines.join('\n'),
            removedLines
        };
    }

    return {
        fullWidthToHalfWidth,
        getVisualLength,
        parseSectionLine,
        cleanLyricContent,
        formatLyricsText,
        classifyLyricLine,
        filterLyricsByLanguage,
        keepOddContentLines
    };
});
