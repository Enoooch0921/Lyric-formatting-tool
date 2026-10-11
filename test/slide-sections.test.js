const test = require('node:test');
const assert = require('node:assert/strict');
const { parseLabel, parseReference, consolidateSlides } = require('../slide-sections.js');
const slides = pages => pages.map((lines, index) => ({ index: index + 2, lines }));

// Synthetic text mirrors split hymn stanzas without publishing the user's lyrics.
const hymn = Array.from({ length: 5 }, (_, index) => [
    [`第${index + 1}節開頭甲乙丙丁`, `第${index + 1}節前半戊己庚辛`],
    [`第${index + 1}節後半壬癸子丑`, `第${index + 1}節結尾寅卯辰巳`]
]);
const reference = hymn.map((pages, index) => `${index + 1}.\n${pages.flat().join('，')}`).join('\n\n');

test('automatic mode does not label eight hymn pages as eight verses', () => {
    const pages = [hymn[0], hymn[1], hymn[2], hymn[4]].flat();
    const result = consolidateSlides(slides([...pages, pages[7]]));
    assert.equal(result.mode, 'auto');
    assert.deepEqual(result.sections.map(item => item.label), ['[1]']);
    assert.deepEqual(result.sections[0].lines, pages.flat());
    assert.match(result.warnings.join(' '), /沒有可靠段落界線/);
});

test('automatic mode retains existing verse/chorus inference using the full performance sequence', () => {
    const a = ['開頭主歌第一行甲乙丙', '開頭主歌第二行丁戊己'];
    const b = ['重複副歌第一行庚辛壬', '重複副歌第二行癸子丑'];
    const result = consolidateSlides(slides([a, b, a, b, b, b]));
    assert.equal(result.mode, 'auto');
    assert.deepEqual(result.sections.map(item => item.label), ['[1]', '[chorus]']);
    assert.deepEqual(result.sections.map(item => item.lines), [a, b]);
});

test('two-page rule creates four stanzas after removing the consecutive ending copy', () => {
    const pages = [hymn[0], hymn[1], hymn[2], hymn[4]].flat();
    const result = consolidateSlides(slides([...pages, pages[7]]), { pagesPerSection: '2' });
    assert.deepEqual(result.sections.map(item => item.label), ['[1]', '[2]', '[3]', '[4]']);
    assert.deepEqual(result.sections.map(item => item.lines), [hymn[0], hymn[1], hymn[2], hymn[4]].map(item => item.flat()));
    assert.deepEqual(result.sections[3].slideIndices, [8, 9, 10]);
    assert.equal(result.sourceSlideCount, 9);
});

test('keeps a partial final group and every stanza beyond verse eight', () => {
    const pages = Array.from({ length: 19 }, (_, index) => [`獨立歌詞頁 ${index + 1}`]);
    const result = consolidateSlides(slides(pages), { pagesPerSection: 2 });
    assert.deepEqual(result.sections.map(item => item.label), Array.from({ length: 8 }, (_, index) => `[${index + 1}]`));
    assert.deepEqual(result.sections.flatMap(item => item.lines), pages.flat());
});

test('deduplicates complete repeated groups, not a common page from different stanzas', () => {
    const [a, b] = hymn[0];
    const c = hymn[1][1];
    const result = consolidateSlides(slides([a, b, a, b, a, c]), { pagesPerSection: 2 });
    assert.deepEqual(result.sections.map(item => item.lines), [[...a, ...b], [...a, ...c]]);
    assert.deepEqual(result.sections.map(item => item.label), ['[1]', '[2]']);
});

test('explicit verse and chorus markers override fixed page counts and retain numbering', () => {
    const result = consolidateSlides(slides([
        ['Verse 1', ...hymn[0][0]], hymn[0][1],
        ['Chorus', ...hymn[1][0]], hymn[1][1],
        ['主歌五', ...hymn[4][0]], hymn[4][1]
    ]), { pagesPerSection: 1 });
    assert.equal(result.mode, 'explicit');
    assert.deepEqual(result.sections.map(item => item.label), ['[1]', '[chorus]', '[5]']);
    assert.deepEqual(result.sections[0].lines, hymn[0].flat());
    assert.deepEqual(result.sections[2].lines, hymn[4].flat());
});

test('joins repeated same-label continuations and deduplicates complete page cycles', () => {
    const result = consolidateSlides(slides([
        ['V1', ...hymn[0][0]], ['V1', ...hymn[0][1]],
        ['V1', ...hymn[0][0]], ['V1', ...hymn[0][1]],
        ['C', ...hymn[1][0]], ['C', ...hymn[1][1]],
        ['V1', ...hymn[0][0]], ['V1', ...hymn[0][1]]
    ]));
    assert.deepEqual(result.sections.map(item => item.label), ['[1]', '[chorus]']);
    assert.deepEqual(result.sections[0].lines, hymn[0].flat());
});

test('keeps distinctly worded sections with the same explicit label', () => {
    const result = consolidateSlides(slides([
        ['V1', ...hymn[0][0]], ['C', ...hymn[1][0]], ['V1', ...hymn[2][0]]
    ]));
    assert.deepEqual(result.sections.map(item => item.label), ['[1]', '[chorus]', '[1]']);
    assert.deepEqual(result.sections[2].lines, hymn[2][0]);
});

test('reference restores the original 1, 2, 3, 5 labels and reports missing verse 4 without inserting it', () => {
    const pages = [hymn[0], hymn[1], hymn[2], hymn[4]].flat();
    const result = consolidateSlides(slides([...pages, pages[7]]), { referenceText: `台語\n${reference}` });
    assert.equal(result.mode, 'reference');
    assert.deepEqual(result.sections.map(item => item.label), ['[1]', '[2]', '[3]', '[5]']);
    assert.deepEqual(result.missingSections, ['[4]']);
    assert.deepEqual(result.sections.flatMap(item => item.lines), pages.flat());
    assert.doesNotMatch(result.text, /第4節/);
    assert.match(result.warnings.join(' '), /未自動補入/);
});

test('reference comparison ignores punctuation and line wrapping, preserving source contractions', () => {
    const original = ["I’m standing on Your faithfulness", 'You are always faithful to me'];
    const result = consolidateSlides(slides([[original[0]], [original[1]]]), {
        referenceText: "1.\nI'm standing on Your faithfulness, You are always faithful to me."
    });
    assert.equal(result.mode, 'reference');
    assert.deepEqual(result.sections[0].lines, original);
    assert.doesNotMatch(result.warnings.join(' '), /字詞差異/);
});

test('minor word differences are only reported, not substituted from the reference', () => {
    const original = ['這是一段很長的台語歌詞其中用惦表示所在位置', '後面還有原始詩節內容應該完整保留'];
    const result = consolidateSlides(slides(original.map(line => [line])), {
        referenceText: `1.\n${original.join('，').replace('惦', 'tiàm')}`
    });
    assert.equal(result.mode, 'reference');
    assert.deepEqual(result.sections[0].lines, original);
    assert.match(result.warnings.join(' '), /字詞差異/);
    assert.doesNotMatch(result.text, /tiàm/);
});

test('flags incomplete reference stanzas instead of filling the missing half', () => {
    const result = consolidateSlides(slides([hymn[0][0]]), { referenceText: reference });
    assert.equal(result.mode, 'reference');
    assert.deepEqual(result.sections[0].lines, hymn[0][0]);
    assert.match(result.warnings.join(' '), /可能缺少部分歌詞/);
    assert.doesNotMatch(result.text, /後半/);
});

test('ambiguous common pages keep the selected page grouping and do not guess a reference stanza', () => {
    const common = ['共同出現的歌詞甲乙丙丁'];
    const result = consolidateSlides(slides([common]), {
        pagesPerSection: 2,
        referenceText: `1.\n${common[0]}\n獨特第一段內容\n\n2.\n${common[0]}\n獨特第二段內容`
    });
    assert.equal(result.mode, 'pages');
    assert.deepEqual(result.missingSections, []);
    assert.match(result.warnings.join(' '), /無法可靠對應/);
    assert.deepEqual(result.sections[0].lines, common);
});

test('unknown pages remain in the output and reference mode falls back without partial relabeling', () => {
    const unknown = ['完全不同的未確認文字不能丟失'];
    const result = consolidateSlides(slides([...hymn[0], unknown]), { referenceText: reference, pagesPerSection: 2 });
    assert.equal(result.mode, 'pages');
    assert.deepEqual(result.sections.flatMap(item => item.lines), [...hymn[0].flat(), ...unknown]);
    assert.match(result.warnings.join(' '), /無法可靠對應/);
});

test('reference cannot override explicit slide labels or reorder the presentation silently', () => {
    const explicit = consolidateSlides(slides([['V4', ...hymn[4].flat()]]), { referenceText: reference });
    assert.equal(explicit.mode, 'explicit');
    assert.equal(explicit.sections[0].label, '[4]');
    assert.match(explicit.warnings.join(' '), /標記保持不變/);
    assert.match(explicit.warnings.join(' '), /\[4\] 內容較接近參考 \[5\]/);
    const reordered = consolidateSlides(slides([...hymn[1], ...hymn[0]]), { referenceText: reference });
    assert.deepEqual(reordered.sections.map(item => item.label), ['[2]', '[1]']);
    assert.match(reordered.warnings.join(' '), /順序與參考不同/);
});

test('short shared text stays uncertain rather than being falsely reported as missing', () => {
    const result = consolidateSlides(slides([['哈利路亞']]), { referenceText: '1.\n哈利路亞，主的恩典豐盛。' });
    assert.equal(result.mode, 'auto');
    assert.deepEqual(result.missingSections, []);
    assert.match(result.warnings.join(' '), /無法可靠對應/);
});

test('clearing the reference restores the chosen page rule', () => {
    const pages = [hymn[0], hymn[1], hymn[2], hymn[4]].flat();
    assert.deepEqual(consolidateSlides(slides(pages), { referenceText: reference, pagesPerSection: 2 }).sections.map(item => item.label), ['[1]', '[2]', '[3]', '[5]']);
    assert.deepEqual(consolidateSlides(slides(pages), { referenceText: '', pagesPerSection: 2 }).sections.map(item => item.label), ['[1]', '[2]', '[3]', '[4]']);
});

test('changing rules and references never mutates source slides', () => {
    const source = slides([...hymn[0], ...hymn[1], hymn[1][1]]);
    const original = structuredClone(source);
    consolidateSlides(source, { pagesPerSection: 2 });
    consolidateSlides(source, { referenceText: reference });
    consolidateSlides(source);
    assert.deepEqual(source, original);
});

test('parses supported labels and rejects invalid references and grouping options', () => {
    assert.deepEqual(parseLabel('主歌八：原始內容'), { label: '[8]', content: '原始內容' });
    assert.deepEqual(parseLabel('副歌'), { label: '[chorus]', content: '' });
    assert.deepEqual(parseLabel('前副歌'), { label: '[prechorus]', content: '' });
    assert.deepEqual(parseLabel('Pro Chorus'), { label: '[prechorus]', content: '' });
    assert.equal(parseLabel('副歌聲音真好聽'), null);
    assert.equal(parseReference(`台語\n${reference}`).length, 5);
    assert.throws(() => parseReference('無編號的文字'), /加入/);
    assert.throws(() => parseReference('9.\n歌詞'), /1 到 8/);
    assert.throws(() => parseReference('1.\n甲\n1.\n乙'), /不同內容/);
    assert.throws(() => parseReference('字'.repeat(20001)), /20,000/);
    assert.throws(() => consolidateSlides(slides(hymn[0]), { pagesPerSection: 5 }), /1 到 4/);
    assert.throws(() => consolidateSlides([]), /沒有可匯入/);
});
