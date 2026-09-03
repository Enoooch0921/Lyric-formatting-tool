const test = require('node:test');
const assert = require('node:assert/strict');

const fflate = require('../vendor/fflate-0.8.2.min.js');
const {
    parsePptxBytes,
    detectSongs,
    consolidateSong
} = require('../pptx-importer.js');

function slideXml(lines, counter = '') {
    const runs = lines.map((line, index) => (
        `${index ? '<a:br/>' : ''}<a:r><a:t>${line}</a:t></a:r>`
    )).join('');
    const footer = counter
        ? `<p:sp><p:spPr><a:xfrm><a:off x="0" y="6000000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>${counter}</a:t></a:r></a:p></p:txBody></p:sp>`
        : '';
    return `<p:sld xmlns:p="p" xmlns:a="a"><p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/></a:xfrm></p:spPr><p:txBody><a:p>${runs}</a:p></p:txBody></p:sp>${footer}</p:sld>`;
}

function createTinyPptx() {
    const files = {
        'ppt/presentation.xml': fflate.strToU8('<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId2"/><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>'),
        'ppt/_rels/presentation.xml.rels': fflate.strToU8('<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>'),
        'ppt/slides/slide1.xml': fflate.strToU8(slideXml(['第二張歌詞'], '2/2')),
        'ppt/slides/slide2.xml': fflate.strToU8(slideXml(['#第一張歌詞', '下一行*'], '1/2')),
        'ppt/slides/_rels/slide1.xml.rels': fflate.strToU8('<Relationships><Relationship Id="rId1" Type="x/image" Target="../media/image1.png"/></Relationships>'),
        'ppt/slides/_rels/slide2.xml.rels': fflate.strToU8('<Relationships><Relationship Id="rId1" Type="x/image" Target="../media/image1.png"/></Relationships>'),
        'ppt/media/ignored-large-image.png': new Uint8Array(1000)
    };
    return fflate.zipSync(files);
}

function makeSlide(index, lines, pageCounter = null, backgroundSignature = '') {
    return { index, lines, pageCounter, backgroundSignature };
}

test('parses PPTX slide order, line breaks, counters, and image signatures', () => {
    const slides = parsePptxBytes(createTinyPptx(), fflate);
    assert.equal(slides.length, 2);
    assert.deepEqual(slides[0].lines, ['第一張歌詞', '下一行']);
    assert.deepEqual(slides[0].pageCounter, { current: 1, total: 2 });
    assert.equal(slides[0].backgroundSignature, 'ppt/media/image1.png');
    assert.deepEqual(slides[1].lines, ['第二張歌詞']);
});

test('detects song boundaries when embedded page numbers restart', () => {
    const songs = detectSongs([
        makeSlide(1, ['第一首甲'], { current: 1, total: 2 }, 'a'),
        makeSlide(2, ['第一首乙'], { current: 2, total: 2 }, 'a'),
        makeSlide(3, ['第二首甲'], { current: 1, total: 2 }, 'b'),
        makeSlide(4, ['第二首乙'], { current: 2, total: 2 }, 'b')
    ]);
    assert.deepEqual(songs.map(song => [song.startSlide, song.endSlide]), [[1, 2], [3, 4]]);
});

test('detects and excludes one-line title slides without page numbers', () => {
    const songs = detectSongs([
        makeSlide(1, ['咱攏成作一家人']),
        makeSlide(2, ['咱攏成作一家人', '第一首歌詞']),
        makeSlide(3, ['其他第一首歌詞']),
        makeSlide(4, ['願']),
        makeSlide(5, ['為這塊土地', '願主真理遍照']),
        makeSlide(6, ['第二首其他歌詞'])
    ]);
    assert.equal(songs.length, 2);
    assert.equal(songs[0].title, '咱攏成作一家人');
    assert.equal(songs[0].contentSlides[0].index, 2);
    assert.equal(songs[1].title, '願');
    assert.equal(songs[1].contentSlides[0].index, 5);
});

test('turns a frequently repeated slide into one chorus and removes contained copies', () => {
    const verse = ['主歌第一行', '主歌第二行'];
    const chorus = ['副歌第一行', '副歌第二行', '副歌第三行'];
    const slides = [verse, chorus, verse, chorus, chorus.slice(1), chorus]
        .map((lines, index) => makeSlide(index + 1, lines));

    const result = consolidateSong(slides);
    assert.deepEqual(result.sections.map(section => section.label), ['[1]', '[chorus]']);
    assert.deepEqual(result.sections[0].lines, verse);
    assert.deepEqual(result.sections[1].lines, chorus);
});

test('groups repeated adjacent slide pairs and merges repeated ending lines', () => {
    const a = ['A1', 'A2'];
    const b = ['B1', 'B2'];
    const c = ['C1', 'C2'];
    const d = ['D1', 'D2'];
    const sequence = [a, b, a, b, c, d, c, d, a, b, c, d, c, d, c, d, [...d, 'D2']];
    const result = consolidateSong(sequence.map((lines, index) => makeSlide(index + 1, lines)));

    assert.deepEqual(result.sections.map(section => section.label), ['[1]', '[chorus]']);
    assert.deepEqual(result.sections[0].lines, [...a, ...b]);
    assert.deepEqual(result.sections[1].lines, [...c, ...d]);
});

test('keeps a song with no repetition as one conservative section', () => {
    const result = consolidateSong([
        makeSlide(1, ['第一頁']),
        makeSlide(2, ['第二頁']),
        makeSlide(3, ['第三頁'])
    ]);
    assert.equal(result.sections.length, 1);
    assert.equal(result.sections[0].label, '[1]');
    assert.deepEqual(result.sections[0].lines, ['第一頁', '第二頁', '第三頁']);
});

test('never creates verse labels above eight', () => {
    const verses = Array.from({ length: 9 }, (_, index) => [`唯一段落${index + 1}甲乙丙`]);
    const chorus = ['共同副歌甲乙丙'];
    const sequence = [...verses, chorus, chorus, chorus];
    const result = consolidateSong(sequence.map((lines, index) => makeSlide(index + 1, lines)));

    assert.deepEqual(
        result.sections.map(section => section.label),
        ['[1]', '[2]', '[3]', '[4]', '[5]', '[6]', '[7]', '[8]', '[chorus]']
    );
    assert.deepEqual(result.sections[7].lines, [...verses[7], ...verses[8]]);
});
