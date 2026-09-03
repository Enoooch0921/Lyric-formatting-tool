(function (root, factory) {
    const api = factory(root);

    if (typeof module === 'object' && module.exports) {
        module.exports = api;
    }

    root.PptxLyricImporter = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
    const MAX_PPTX_BYTES = 50 * 1024 * 1024;
    const MAX_SELECTED_XML_BYTES = 20 * 1024 * 1024;
    const MAX_SINGLE_XML_BYTES = 3 * 1024 * 1024;

    function getZipApi(explicitApi) {
        if (explicitApi) return explicitApi;
        if (root.fflate) return root.fflate;
        if (typeof require === 'function') return require('./vendor/fflate-0.8.2.min.js');
        throw new Error('缺少 PPTX 解壓縮元件。');
    }

    function normalizePartPath(path) {
        const parts = [];
        path.replace(/^\//, '').split('/').forEach(part => {
            if (!part || part === '.') return;
            if (part === '..') parts.pop();
            else parts.push(part);
        });
        return parts.join('/');
    }

    function resolvePartPath(basePart, target) {
        if (target.startsWith('/')) return normalizePartPath(target);
        const baseDirectory = basePart.split('/').slice(0, -1).join('/');
        return normalizePartPath(`${baseDirectory}/${target}`);
    }

    function relationshipPartPath(partPath) {
        const pieces = partPath.split('/');
        const fileName = pieces.pop();
        return `${pieces.join('/')}/_rels/${fileName}.rels`;
    }

    function decodeXmlEntities(text) {
        return text
            .replace(/&#x([0-9a-f]+);/gi, (_, value) => String.fromCodePoint(Number.parseInt(value, 16)))
            .replace(/&#(\d+);/g, (_, value) => String.fromCodePoint(Number.parseInt(value, 10)))
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&amp;/g, '&');
    }

    function parseAttributes(source) {
        const attributes = {};
        const attributePattern = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
        let match;
        while ((match = attributePattern.exec(source))) {
            attributes[match[1]] = decodeXmlEntities(match[2] ?? match[3] ?? '');
        }
        return attributes;
    }

    function parseRelationships(xml) {
        const relationships = new Map();
        const pattern = /<Relationship\b([^>]*)\/?\s*>/g;
        let match;
        while ((match = pattern.exec(xml))) {
            const attributes = parseAttributes(match[1]);
            if (attributes.Id) relationships.set(attributes.Id, attributes);
        }
        return relationships;
    }

    function extractSlideOrder(files, decodeFile) {
        const presentationPart = 'ppt/presentation.xml';
        const presentationRelsPart = 'ppt/_rels/presentation.xml.rels';
        const presentationXml = files[presentationPart] ? decodeFile(presentationPart) : '';
        const relationshipsXml = files[presentationRelsPart] ? decodeFile(presentationRelsPart) : '';

        if (presentationXml && relationshipsXml) {
            const relationships = parseRelationships(relationshipsXml);
            const orderedSlides = [];
            const slidePattern = /<p:sldId\b([^>]*)\/?\s*>/g;
            let match;

            while ((match = slidePattern.exec(presentationXml))) {
                const relationshipId = parseAttributes(match[1])['r:id'];
                const relationship = relationships.get(relationshipId);
                if (!relationship || !relationship.Target) continue;
                const slidePart = resolvePartPath(presentationPart, relationship.Target);
                if (files[slidePart]) orderedSlides.push(slidePart);
            }

            if (orderedSlides.length) return orderedSlides;
        }

        return Object.keys(files)
            .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
            .sort((left, right) => {
                const leftNumber = Number(left.match(/slide(\d+)\.xml$/)[1]);
                const rightNumber = Number(right.match(/slide(\d+)\.xml$/)[1]);
                return leftNumber - rightNumber;
            });
    }

    function normalizeExtractedLine(line) {
        return line
            .replace(/\u00a0/g, ' ')
            .replace(/[\t ]+/g, ' ')
            .trim()
            .replace(/^[#*]+\s*/, '')
            .replace(/\s*[#*]+$/, '');
    }

    function extractTextLines(textBody) {
        const lines = [];
        const paragraphPattern = /<a:p(?:\s[^>]*)?>([\s\S]*?)<\/a:p>/g;
        let paragraphMatch;

        while ((paragraphMatch = paragraphPattern.exec(textBody))) {
            const paragraph = paragraphMatch[1];
            const tokenPattern = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>|<a:br\b[^>]*>/g;
            let tokenMatch;
            let currentLine = '';

            while ((tokenMatch = tokenPattern.exec(paragraph))) {
                if (tokenMatch[1] !== undefined) {
                    currentLine += decodeXmlEntities(tokenMatch[1]);
                } else {
                    const normalized = normalizeExtractedLine(currentLine);
                    if (normalized) lines.push(normalized);
                    currentLine = '';
                }
            }

            const normalized = normalizeExtractedLine(currentLine);
            if (normalized) lines.push(normalized);
        }

        return lines;
    }

    function parsePageCounter(line) {
        const match = line.match(/^(\d{1,3})\s*(?:\/|／|of)\s*(\d{1,3})$/i);
        if (!match) return null;
        return { current: Number(match[1]), total: Number(match[2]) };
    }

    function extractSlide(slidePart, slideXml, relationshipsXml, slideIndex) {
        const shapes = [];
        const shapePattern = /<p:sp\b[\s\S]*?<\/p:sp>/g;
        let shapeMatch;

        while ((shapeMatch = shapePattern.exec(slideXml))) {
            const shapeXml = shapeMatch[0];
            const textBodyMatch = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
            if (!textBodyMatch) continue;

            const lines = extractTextLines(textBodyMatch[1]);
            if (!lines.length) continue;

            const offsetMatch = shapeXml.match(/<a:off\b([^>]*)\/?\s*>/);
            const offset = offsetMatch ? parseAttributes(offsetMatch[1]) : {};
            shapes.push({
                x: Number(offset.x || 0),
                y: Number(offset.y || 0),
                lines
            });
        }

        shapes.sort((left, right) => left.y - right.y || left.x - right.x);

        let pageCounter = null;
        const lyricLines = [];
        shapes.forEach(shape => {
            shape.lines.forEach(line => {
                const counter = parsePageCounter(line);
                if (counter) pageCounter = counter;
                else lyricLines.push(line);
            });
        });

        const relationships = relationshipsXml ? parseRelationships(relationshipsXml) : new Map();
        const backgroundParts = [...relationships.values()]
            .filter(relationship => /\/image$/.test(relationship.Type || '') && relationship.Target)
            .map(relationship => resolvePartPath(slidePart, relationship.Target))
            .sort();

        return {
            index: slideIndex,
            part: slidePart,
            lines: lyricLines,
            pageCounter,
            backgroundSignature: backgroundParts.join('|')
        };
    }

    function parsePptxBytes(bytes, explicitZipApi) {
        const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
        if (data.byteLength > MAX_PPTX_BYTES) {
            throw new Error('PPTX 超過 50 MB，請先壓縮圖片或拆分檔案。');
        }

        const zipApi = getZipApi(explicitZipApi);
        let selectedXmlBytes = 0;
        const selectedPartPattern = /^(?:ppt\/presentation\.xml|ppt\/_rels\/presentation\.xml\.rels|ppt\/slides\/slide\d+\.xml|ppt\/slides\/_rels\/slide\d+\.xml\.rels)$/;
        const files = zipApi.unzipSync(data, {
            filter(file) {
                const name = normalizePartPath(file.name);
                if (!selectedPartPattern.test(name)) return false;
                if (file.originalSize > MAX_SINGLE_XML_BYTES) {
                    throw new Error('PPTX 中有異常過大的 XML，已停止解析。');
                }
                selectedXmlBytes += file.originalSize;
                if (selectedXmlBytes > MAX_SELECTED_XML_BYTES) {
                    throw new Error('PPTX 文字資料過大，已停止解析。');
                }
                return true;
            }
        });

        const decoder = new TextDecoder('utf-8');
        const decodeFile = part => decoder.decode(files[part]);
        const slideParts = extractSlideOrder(files, decodeFile);
        if (!slideParts.length) throw new Error('找不到 PPTX 投影片內容。');

        const slides = slideParts.map((slidePart, index) => {
            const relsPart = relationshipPartPath(slidePart);
            return extractSlide(
                slidePart,
                decodeFile(slidePart),
                files[relsPart] ? decodeFile(relsPart) : '',
                index + 1
            );
        }).filter(slide => slide.lines.length);

        if (!slides.length) throw new Error('PPTX 裡找不到可匯入的文字。');
        return slides;
    }

    async function parsePptxFile(file) {
        if (!file || !/\.pptx$/i.test(file.name || '')) {
            throw new Error('目前只支援 .pptx 檔案。');
        }
        if (file.size > MAX_PPTX_BYTES) {
            throw new Error('PPTX 超過 50 MB，請先壓縮圖片或拆分檔案。');
        }
        return parsePptxBytes(await file.arrayBuffer());
    }

    function splitByBackground(slides) {
        const runs = [];
        slides.forEach(slide => {
            const signature = slide.backgroundSignature || '';
            const previousRun = runs[runs.length - 1];
            if (!previousRun || previousRun.signature !== signature) {
                runs.push({ signature, slides: [slide] });
            } else {
                previousRun.slides.push(slide);
            }
        });

        const usefulRuns = runs.filter(run => run.signature);
        const canUseRuns = runs.length >= 2
            && runs.length <= 6
            && usefulRuns.length === runs.length
            && runs.every(run => run.slides.length >= 2);

        return canUseRuns ? runs.map(run => run.slides) : [slides];
    }

    function isLikelyTitleSlide(slides, index) {
        const slide = slides[index];
        if (!slide || slide.lines.length !== 1) return false;

        const title = normalizeForComparison(slide.lines[0]);
        if (!title) return false;

        const nearbyContent = slides.slice(index + 1, index + 3)
            .flatMap(item => item.lines)
            .map(normalizeForComparison);
        const titleAppearsNearby = nearbyContent.some(line => line.includes(title));

        const previousBackground = slides[index - 1]?.backgroundSignature;
        const currentBackground = slide.backgroundSignature;
        const nextBackground = slides[index + 1]?.backgroundSignature;
        const backgroundBreak = Boolean(previousBackground
            && currentBackground
            && nextBackground
            && previousBackground !== currentBackground
            && previousBackground !== nextBackground);

        return titleAppearsNearby || backgroundBreak;
    }

    function detectSongs(slides) {
        const groups = [];
        let currentGroup = [];

        slides.forEach((slide, index) => {
            const hasPriorCounter = currentGroup.some(item => item.pageCounter);
            const counterRestarts = slide.pageCounter?.current === 1 && hasPriorCounter;
            const titleStartsSong = currentGroup.length >= 2 && isLikelyTitleSlide(slides, index);
            const startsNewSong = counterRestarts || titleStartsSong;
            if (startsNewSong) {
                groups.push(currentGroup);
                currentGroup = [];
            }
            currentGroup.push(slide);
        });
        if (currentGroup.length) groups.push(currentGroup);

        const detectedGroups = groups.length > 1 ? groups : splitByBackground(slides);
        return detectedGroups.map((songSlides, index) => ({
            index,
            slides: songSlides,
            contentSlides: isLikelyTitleSlide(songSlides, 0) ? songSlides.slice(1) : songSlides,
            title: isLikelyTitleSlide(songSlides, 0) ? songSlides[0].lines[0] : '',
            startSlide: songSlides[0].index,
            endSlide: songSlides[songSlides.length - 1].index,
            preview: songSlides[0].lines.join(' ').slice(0, 24) || `第 ${index + 1} 首`
        }));
    }

    function normalizeForComparison(text) {
        return text.normalize('NFKC').toLowerCase().replace(/[\p{P}\p{S}\s]+/gu, '');
    }

    function dedupeConsecutiveLines(lines) {
        const deduped = [];
        lines.forEach(line => {
            const normalized = normalizeExtractedLine(line);
            if (!normalized) return;
            const lastLine = deduped[deduped.length - 1];
            if (!lastLine || normalizeForComparison(lastLine) !== normalizeForComparison(normalized)) {
                deduped.push(normalized);
            }
        });
        return deduped;
    }

    function countPatternOccurrences(sequence, pattern) {
        if (!pattern.length) return 0;
        let count = 0;
        for (let index = 0; index <= sequence.length - pattern.length; index++) {
            if (pattern.every((key, offset) => sequence[index + offset] === key)) count++;
        }
        return count;
    }

    function mergeLineSequences(lineGroups) {
        const merged = [];
        lineGroups.forEach(lines => {
            let overlap = 0;
            const maxOverlap = Math.min(merged.length, lines.length);
            for (let size = maxOverlap; size > 0; size--) {
                const suffix = merged.slice(-size).map(normalizeForComparison);
                const prefix = lines.slice(0, size).map(normalizeForComparison);
                if (suffix.every((line, index) => line === prefix[index])) {
                    overlap = size;
                    break;
                }
            }
            merged.push(...lines.slice(overlap));
        });
        return dedupeConsecutiveLines(merged);
    }

    function consolidateSong(slides) {
        const rawBlocks = slides.map(slide => {
            const lines = dedupeConsecutiveLines(slide.lines);
            return {
                slideIndex: slide.index,
                lines,
                flat: lines.map(normalizeForComparison).join('')
            };
        }).filter(block => block.flat);

        if (!rawBlocks.length) throw new Error('這首歌沒有可匯入的歌詞。');

        const definitions = new Map();
        rawBlocks.forEach(block => {
            if (!definitions.has(block.flat)) definitions.set(block.flat, block);
        });

        const parentByKey = new Map();
        definitions.forEach((block, key) => {
            const candidates = [...definitions.values()]
                .filter(candidate => candidate.flat.length > block.flat.length
                    && block.flat.length >= 6
                    && candidate.flat.includes(block.flat))
                .sort((left, right) => left.flat.length - right.flat.length);
            parentByKey.set(key, candidates[0]?.flat || key);
        });

        function resolveParent(key) {
            let resolved = key;
            const visited = new Set();
            while (parentByKey.get(resolved) && parentByKey.get(resolved) !== resolved && !visited.has(resolved)) {
                visited.add(resolved);
                resolved = parentByKey.get(resolved);
            }
            return resolved;
        }

        const sequence = rawBlocks.map(block => resolveParent(block.flat));
        const firstAppearance = [...new Set(sequence)];
        const occurrenceCounts = new Map(firstAppearance.map(key => [
            key,
            sequence.filter(item => item === key).length
        ]));

        const groups = [];
        for (let index = 0; index < firstAppearance.length; index++) {
            const current = firstAppearance[index];
            const next = firstAppearance[index + 1];
            if (next) {
                const pairRepeats = countPatternOccurrences(sequence, [current, next]);
                const currentCount = occurrenceCounts.get(current);
                const nextCount = occurrenceCounts.get(next);
                const balance = Math.min(currentCount, nextCount) / Math.max(currentCount, nextCount);
                if (pairRepeats >= 2 && balance >= 0.6) {
                    groups.push([current, next]);
                    index++;
                    continue;
                }
            }
            groups.push([current]);
        }

        const repeats = groups.map(group => countPatternOccurrences(sequence, group));
        let chorusIndex = -1;
        if (groups.length > 1) {
            repeats.forEach((repeatCount, index) => {
                if (index === 0) return;
                if (repeatCount >= 2 && (chorusIndex === -1 || repeatCount >= repeats[chorusIndex])) {
                    chorusIndex = index;
                }
            });
        }

        const hasRepeatedGroup = repeats.some(repeatCount => repeatCount >= 2);
        if (!hasRepeatedGroup && groups.length > 1) {
            groups.splice(0, groups.length, firstAppearance);
            repeats.splice(0, repeats.length, 1);
        }

        const groupByKey = new Map();
        groups.forEach((group, groupIndex) => {
            group.forEach(key => groupByKey.set(key, groupIndex));
        });
        const performanceSequence = sequence
            .map(key => groupByKey.get(key))
            .filter((groupIndex, index, items) => index === 0 || groupIndex !== items[index - 1]);

        let prechorusIndex = -1;
        let bridgeIndex = -1;
        if (chorusIndex > 0) {
            let bestPrechorusTransitions = 0;
            for (let index = 1; index < groups.length; index++) {
                if (index === chorusIndex || repeats[index] < 2) continue;
                const transitionsToChorus = performanceSequence.reduce((count, groupIndex, position) => (
                    count + (groupIndex === index && performanceSequence[position + 1] === chorusIndex ? 1 : 0)
                ), 0);
                if (transitionsToChorus >= 2 && transitionsToChorus >= bestPrechorusTransitions) {
                    prechorusIndex = index;
                    bestPrechorusTransitions = transitionsToChorus;
                }
            }

            for (let index = 1; index < groups.length; index++) {
                if (index === chorusIndex || index === prechorusIndex || repeats[index] !== 1) continue;
                const position = performanceSequence.indexOf(index);
                const chorusesBefore = performanceSequence.slice(0, position)
                    .filter(groupIndex => groupIndex === chorusIndex).length;
                const chorusAfter = performanceSequence.slice(position + 1).includes(chorusIndex);
                if (position !== -1 && chorusesBefore >= 2 && chorusAfter) {
                    bridgeIndex = index;
                    break;
                }
            }
        }

        let verseNumber = 0;
        const sections = [];
        groups.forEach((group, index) => {
            const lineGroups = group.map(key => definitions.get(key).lines);
            const isChorus = index === chorusIndex;
            const isPrechorus = index === prechorusIndex;
            const isBridge = index === bridgeIndex;
            const isVerse = !isChorus && !isPrechorus && !isBridge;
            if (isVerse) verseNumber++;
            const label = isChorus
                ? '[chorus]'
                : isPrechorus
                    ? '[prechorus]'
                    : isBridge
                        ? '[bridge]'
                        : `[${Math.min(verseNumber, 8)}]`;
            const section = {
                label,
                lines: mergeLineSequences(lineGroups),
                repeatCount: repeats[index]
            };

            if (isVerse && verseNumber > 8) {
                const verseEight = sections.find(item => item.label === '[8]');
                verseEight.lines = mergeLineSequences([verseEight.lines, section.lines]);
                verseEight.repeatCount = Math.max(verseEight.repeatCount, section.repeatCount);
            } else {
                sections.push(section);
            }
        });

        const text = sections.flatMap(section => [section.label, ...section.lines]).join('\n');
        return {
            text,
            sections,
            sourceSlideCount: slides.length,
            uniqueBlockCount: firstAppearance.length
        };
    }

    return {
        parsePageCounter,
        extractTextLines,
        parsePptxBytes,
        parsePptxFile,
        detectSongs,
        consolidateSong,
        normalizeForComparison,
        dedupeConsecutiveLines
    };
});
