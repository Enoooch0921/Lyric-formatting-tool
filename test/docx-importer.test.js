const test = require('node:test');
const assert = require('node:assert/strict');
const { DOMParser } = require('@xmldom/xmldom');
const zip = require('../vendor/fflate-0.8.2.min.js');
const { parseDocxBytes, parseDocxFile, detectSongs, consolidateSong } = require('../docx-importer.js');
const { formatLyricsText } = require('../lyric-formatter.js');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function escape(text) {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function p(text = '', properties = '') {
    return `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}<w:r><w:t>${escape(text)}</w:t></w:r></w:p>`;
}
function makeDocx(body, extraParts = {}) {
    return zip.zipSync({
        'word/document.xml': zip.strToU8(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`),
        ...extraParts
    });
}
function parse(body, extraParts) {
    return parseDocxBytes(makeDocx(body, extraParts), zip, DOMParser);
}
function songs(body, extraParts) {
    return detectSongs(parse(body, extraParts).paragraphs);
}
function sections(body) {
    return consolidateSong(parse(body).paragraphs);
}

test('reads split Word runs, entities, tabs and manual line breaks without losing punctuation', () => {
    const result = parse(`<w:p><w:r><w:t>I’m stand</w:t></w:r><w:r><w:t>ing on Your faithfulness</w:t>
        <w:br/><w:t>Love &amp; hope</w:t><w:tab/><w:t>永遠</w:t><w:cr/><w:t>&lt;script&gt;</w:t></w:r></w:p>${p()}`);
    assert.deepEqual(result.paragraphs[0].lines, ['I’m standing on Your faithfulness', 'Love & hope 永遠', '<script>']);
    assert.deepEqual(result.paragraphs[1].lines, ['']);
    assert.equal(formatLyricsText(consolidateSong(result.paragraphs).text, 13).split('\n')[1], 'I’m standing on Your faithfulness');
});

test('detects the two-song layout in the supplied Word example without publishing its lyrics', () => {
    const first = Array.from({ length: 7 }, (_, i) => `第一首內容第${i + 1}行`);
    first[2] = '第一首歌名 也在歌詞中出現';
    const second = Array.from({ length: 8 }, (_, i) => `第二首內容第${i + 1}行`);
    second[0] = '第二首歌名';
    const result = songs([p('第一首歌名'), p(), ...first.map(line => p(line)), p(), p(),
        p('第二首歌名'), p(), ...second.map(line => p(line))].join(''));
    assert.deepEqual(result.map(song => song.title), ['第一首歌名', '第二首歌名']);
    assert.deepEqual(result.map(song => [song.startParagraph, song.endParagraph]), [[3, 9], [14, 21]]);
    assert.deepEqual(result.map(song => consolidateSong(song.paragraphs).sections.map(section => section.label)), [['[1]'], ['[1]']]);
    assert.deepEqual(result.map(song => consolidateSong(song.paragraphs).sections[0].lines), [first, second]);
});

test('does not split songs on blank spacing alone or remove repeated words from different lines', () => {
    const result = songs([p('首段甲'), p('首段乙'), p(), p(), p('下段甲'), p('下段乙'), p('下段甲 和其他內容')].join(''));
    assert.equal(result.length, 1);
    assert.equal(result[0].title, '');
    const merged = consolidateSong(result[0].paragraphs);
    assert.deepEqual(merged.sections.map(section => section.label), ['[1]', '[2]']);
    assert.deepEqual(merged.sections[1].lines, ['下段甲', '下段乙', '下段甲 和其他內容']);
});

test('uses Word title styles to identify songs and respects explicit Verse and Pre Chorus labels', () => {
    const result = songs([
        p('歌名甲', '<w:pStyle w:val="SongTitle"/>'), p('Verse 1'), p('主歌一行'), p('主歌二行'),
        p('Pre Chorus'), p('副歌前段一行'), p('Chorus'), p('副歌一行'),
        p('歌名乙', '<w:pStyle w:val="SongTitle"/>'), p('Verse 2'), p('另一首的歌詞')
    ].join(''), { 'word/styles.xml': zip.strToU8(`<w:styles xmlns:w="${W}"><w:style w:styleId="SongTitle"><w:name w:val="Title"/></w:style></w:styles>`) });
    assert.equal(result.length, 2);
    assert.deepEqual(result.map(song => song.title), ['歌名甲', '歌名乙']);
    assert.deepEqual(consolidateSong(result[0].paragraphs).sections.map(section => section.label), ['[1]', '[prechorus]', '[chorus]']);
    assert.equal(consolidateSong(result[1].paragraphs).sections[0].label, '[2]');
});

test('merges only complete repeated stanzas, preserves variants and stays conservative about chorus', () => {
    const result = sections([p('主歌甲'), p('主歌乙'), p(), p('重複段甲'), p('重複段乙'), p(),
        p('重複段甲'), p('重複段乙'), p(), p('重複段甲'), p('不同結尾')].join(''));
    assert.deepEqual(result.sections.map(section => section.label), ['[1]', '[2]', '[3]']);
    assert.equal(result.mergedBlockCount, 1);
    assert.equal(result.sections[1].repeatCount, 2);
    assert.deepEqual(result.sections[2].lines, ['重複段甲', '不同結尾']);
});

test('merges explicit repeated choruses without discarding differently labeled identical lyrics', () => {
    const result = sections([p('Verse 1'), p('共同歌詞'), p('Chorus'), p('共同歌詞'), p('Chorus'), p('共同歌詞')].join(''));
    assert.deepEqual(result.sections.map(section => section.label), ['[1]', '[chorus]']);
    assert.equal(result.sections[1].repeatCount, 2);
});

test('keeps every stanza when verse numbering reaches eight', () => {
    const result = sections(Array.from({ length: 10 }, (_, i) => p(`第${i + 1}段歌詞`) + p()).join(''));
    assert.deepEqual(result.sections.map(section => section.label), Array.from({ length: 8 }, (_, i) => `[${i + 1}]`));
    assert.deepEqual(result.sections[7].lines.filter(Boolean), ['第8段歌詞', '第9段歌詞', '第10段歌詞']);
});

test('reads tables in document order and excludes headers, deletions, hidden text and field instructions', () => {
    const result = parse(`${p('正文')}<w:tbl><w:tr><w:tc>${p('左欄')}</w:tc><w:tc>${p('右欄')}</w:tc></w:tr></w:tbl>
        <w:p><w:del><w:r><w:t>刪除文字</w:t></w:r></w:del><w:r><w:rPr><w:vanish/></w:rPr><w:t>隱藏文字</w:t></w:r>
        <w:r><w:instrText>HYPERLINK secret</w:instrText></w:r><w:ins><w:r><w:t>保留文字</w:t></w:r></w:ins></w:p>`,
        { 'word/header1.xml': zip.strToU8(p('頁首頁碼')) });
    assert.deepEqual(result.paragraphs.map(paragraph => paragraph.text), ['正文', '左欄', '右欄', '保留文字']);
    assert.match(result.warnings[0], /表格/);
});

test('reports unsupported image content and rejects image-only documents', () => {
    const drawing = '<w:p><w:r><w:drawing/></w:r></w:p>';
    assert.match(parse(p('歌詞文字') + drawing).warnings[0], /圖片/);
    assert.throws(() => parse(drawing), /圖片或掃描/);
});

test('supports alternate Word namespace prefixes and strict DOCX namespaces', () => {
    const xml = '<x:document xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main"><x:body><x:p><x:r><x:t>歌詞文字</x:t></x:r></x:p></x:body></x:document>';
    const bytes = zip.zipSync({ 'word/document.xml': zip.strToU8(xml) });
    assert.equal(parseDocxBytes(bytes, zip, DOMParser).paragraphs[0].text, '歌詞文字');
});

test('rejects unsupported, corrupt, empty and oversized files with useful errors', async () => {
    await assert.rejects(parseDocxFile({ name: 'old.doc' }), /另存為/);
    await assert.rejects(parseDocxFile({ name: 'big.docx', size: 51 * 1024 * 1024 }), /50 MB/);
    assert.throws(() => parseDocxBytes(new Uint8Array([1, 2, 3]), zip, DOMParser), /無法讀取 DOCX/);
    assert.throws(() => parseDocxBytes(zip.zipSync({ 'unrelated.txt': zip.strToU8('text') }), zip, DOMParser), /找不到 DOCX/);
    assert.throws(() => parse('<w:p><w:r>'), /文字結構損壞/);
    assert.throws(() => parse(''), /找不到可匯入/);
    assert.throws(() => parse(p('x'.repeat(4 * 1024 * 1024))), /文字資料過大/);
});
