const test = require('node:test');
const assert = require('node:assert/strict');

const {
    parseSectionLine,
    cleanLyricContent,
    formatLyricsText,
    classifyLyricLine,
    filterLyricsByLanguage,
    keepOddContentLines
} = require('../lyric-formatter.js');

test('preserves apostrophes inside English contractions', () => {
    assert.equal(cleanLyricContent("You're all I need"), "You're all I need");
    assert.equal(cleanLyricContent('I\u2019m Yours'), 'I\u2019m Yours');
    assert.equal(cleanLyricContent('\u201cYou\u2019re here!\u201d'), 'You\u2019re here');
});

test('keeps a short English lyric together instead of orphaning its last word', () => {
    assert.equal(
        formatLyricsText('I\u2019m standing on Your faithfulness', 13),
        'I\u2019m standing on Your faithfulness'
    );
});

test('removes punctuation that was previously missed', () => {
    assert.equal(
        cleanLyricContent('祢真偉大……永遠掌權—直到萬代！'),
        '祢真偉大 永遠掌權 直到萬代'
    );
});

test('recognizes verse numbers 1 through 8 without matching other numbers or words', () => {
    assert.deepEqual(parseSectionLine('Verse 8: 歌詞'), {
        label: '[8]',
        content: '歌詞'
    });
    assert.deepEqual(parseSectionLine('[7] 歌詞'), {
        label: '[7]',
        content: '歌詞'
    });
    assert.deepEqual(parseSectionLine('6'), {
        label: '[6]',
        content: ''
    });
    assert.deepEqual(parseSectionLine('6 第六段歌詞'), {
        label: '[6]',
        content: '第六段歌詞'
    });
    assert.equal(parseSectionLine('Verse 9'), null);
    assert.equal(parseSectionLine('[9]'), null);
    assert.equal(parseSectionLine('9'), null);
    assert.equal(parseSectionLine('v1ctory belongs to Jesus'), null);
    assert.equal(parseSectionLine('10000 Reasons'), null);
});

test('maps Pro Chorus typo variants to prechorus', () => {
    for (const input of ['Pro Chorus', 'pro-chorus', 'ProChorus', 'PRO C']) {
        assert.deepEqual(parseSectionLine(input), {
            label: '[prechorus]',
            content: ''
        });
    }

    assert.equal(formatLyricsText('Pro Chorus\n祢真偉大', 13), '[prechorus]\n祢真偉大');
});

test('classifies language conservatively', () => {
    assert.equal(classifyLyricLine('單單仰望祢'), 'chinese');
    assert.equal(classifyLyricLine('I worship You'), 'english');
    assert.equal(classifyLyricLine('Jesus 我愛祢'), 'mixed');
    assert.equal(classifyLyricLine('♪ 123'), 'neutral');
    assert.equal(classifyLyricLine('Pre Chorus'), 'section');
});

test('language filters keep section, mixed, neutral, and blank lines', () => {
    const input = [
        'Verse 8',
        '單單仰望祢',
        'I worship You',
        'Jesus 我愛祢',
        '♪',
        '',
        'Pro Chorus'
    ].join('\n');

    const chinese = filterLyricsByLanguage(input, 'chinese');
    assert.equal(chinese.text, [
        'Verse 8',
        '單單仰望祢',
        'Jesus 我愛祢',
        '♪',
        '',
        'Pro Chorus'
    ].join('\n'));
    assert.deepEqual(chinese.removedLines, [{ lineNumber: 3, text: 'I worship You' }]);
    assert.equal(chinese.ambiguousCount, 2);

    const english = filterLyricsByLanguage(input, 'english');
    assert.equal(english.text, [
        'Verse 8',
        'I worship You',
        'Jesus 我愛祢',
        '♪',
        '',
        'Pro Chorus'
    ].join('\n'));
    assert.deepEqual(english.removedLines, [{ lineNumber: 2, text: '單單仰望祢' }]);
});

test('odd-line removal resets per section and stanza while preserving structure', () => {
    const input = [
        'Verse 1',
        '中文一',
        'English one',
        '',
        '中文二',
        'English two',
        'Chorus',
        '中文三',
        'English three'
    ].join('\n');

    const result = keepOddContentLines(input);
    assert.equal(result.text, [
        'Verse 1',
        '中文一',
        '',
        '中文二',
        'Chorus',
        '中文三'
    ].join('\n'));
    assert.deepEqual(result.removedLines.map(line => line.lineNumber), [3, 6, 9]);
});

test('formatting clamps the configured character limit', () => {
    const input = 'one two three four five six seven eight nine ten eleven twelve';
    assert.equal(formatLyricsText(input, 2), formatLyricsText(input, 5));
    assert.equal(formatLyricsText(input, 100), formatLyricsText(input, 40));
});
