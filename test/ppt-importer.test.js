const test = require('node:test');
const assert = require('node:assert/strict');
const CFB = require('../vendor/cfb-1.2.2.min.js');
const fflate = require('../vendor/fflate-0.8.2.min.js');
const { parsePptBytes, parsePptFile, isLegacyPpt } = require('../ppt-importer.js');

// Synthetic binary records: no user's deck or copyrighted lyrics are committed.
function atom(type, data = Buffer.alloc(0), instance = 0, container = false) {
    const header = Buffer.alloc(8);
    header.writeUInt16LE((instance << 4) | (container ? 15 : 0));
    header.writeUInt16LE(type, 2);
    header.writeUInt32LE(data.length, 4);
    return Buffer.concat([header, data]);
}
function uints(...values) {
    const bytes = Buffer.alloc(values.length * 4);
    values.forEach((value, index) => bytes.writeUInt32LE(value, index * 4));
    return bytes;
}
function shape(text, { y = 100, x = 0, reference, placeholder, field, small = true } = {}) {
    const anchor = Buffer.alloc(small ? 8 : 16);
    if (small) { anchor.writeInt16LE(y); anchor.writeInt16LE(x, 2); }
    else { anchor.writeInt32LE(x); anchor.writeInt32LE(y, 4); }
    const parts = [atom(0xf010, anchor)];
    if (placeholder !== undefined) {
        const data = Buffer.alloc(8);
        data[4] = placeholder;
        parts.push(atom(0xf011, atom(3011, data)));
    }
    const textRecords = reference !== undefined
        ? atom(3998, uints(reference))
        : atom(4000, Buffer.from(text || '', 'utf16le'));
    parts.push(atom(0xf00d, Buffer.concat([textRecords, field ? atom(field, uints(0)) : Buffer.alloc(0)])));
    return atom(0xf004, Buffer.concat(parts), 0, true);
}
function slide(...shapes) {
    return atom(1006, atom(1036, atom(0xf002, atom(0xf003, Buffer.concat(shapes), 0, true), 0, true)), 0, true);
}
function ole(streams) {
    const cfb = CFB.utils.cfb_new();
    Object.entries(streams).forEach(([name, bytes]) => CFB.utils.cfb_add(cfb, name, bytes));
    return Uint8Array.from(CFB.write(cfb, { type: 'buffer' }));
}
function fixture({ cycle = false, corruptOffset = false, encrypted = false, noText = false } = {}) {
    const chunks = [];
    let length = 0;
    function add(bytes) { const offset = length; chunks.push(bytes); length += bytes.length; return offset; }
    // Start with unused records and an earlier version of slide persist ID 2.
    const stale = add(slide(shape('歷史編輯舊歌詞，不應匯入')));
    const actual = add(noText ? slide() : slide(
        shape('', { reference: 0, y: 100, small: false }),
        shape('', { reference: 1, placeholder: 9 }),
        shape('9-1', { y: 8500 }),
        shape('<number>', { y: 8900, field: 4056 }),
        shape('2026-10-11', { y: 9000, field: 4087 }),
        shape('日期欄位', { y: 9500, field: 4088 })
    ));
    const other = add(noText ? slide() : slide(shape('第二頁，I’m faithful', { y: 200 })));
    add(atom(1016, atom(4000, Buffer.from('備註，不應匯入', 'utf16le')), 0, true));
    const master = add(atom(1016, atom(4000, Buffer.from('大標題文字，母片，不應匯入', 'utf16le')), 0, true));
    add(atom(4000, Buffer.from('deleted slide text', 'utf16le')));
    const persistSlide = id => atom(1011, uints(id, 0, 0, id, 0));
    const outlines = noText ? [] : [
        atom(3999, uints(1)), atom(4000, Buffer.from('第一頁中文\r保留英文 Don’t\v下一行', 'utf16le')),
        atom(3999, uints(4)), atom(4000, Buffer.from('頁尾參照，不應重新匯入', 'utf16le')),
        atom(3999, uints(5)), atom(4008, Buffer.from('Center body text', 'latin1')),
        atom(3999, uints(2)), atom(4000, Buffer.from('筆記型文字，不應匯入', 'utf16le'))
    ];
    const doc = add(atom(1000, Buffer.concat([
        atom(1001, uints(10000, 10000)),
        atom(4080, persistSlide(4), 1, true), // Master list, not active slides.
        atom(4080, Buffer.concat([persistSlide(3), persistSlide(2), ...outlines]), 0, true)
    ]), 0, true));
    const oldDir = add(atom(6002, uints((4 << 20) | 1, doc, stale, other, master)));
    const oldEdit = add(atom(4085, uints(0, 0, 0, oldDir, 1, 5, 0)));
    const newDir = add(atom(6002, uints((1 << 20) | 2, corruptOffset ? 0xffffffff : actual)));
    const newEditOffset = length;
    const newEdit = add(atom(4085, uints(0, 0, cycle ? newEditOffset : oldEdit, newDir, 1, 5, 0)));
    const currentUser = atom(4086, uints(20, encrypted ? 0xf3d1c4df : 0xe391c05f, newEdit, 0, 0, 0));
    return ole({ 'PowerPoint Document': Buffer.concat(chunks), 'Current User': currentUser,
        'SummaryInformation': Buffer.from('作者及文件屬性，不應匯入', 'utf16le') });
}

test('reads genuine CFB PPT in active slide order and resolves latest edit objects only', () => {
    const bytes = fixture();
    assert.equal(isLegacyPpt(bytes), true);
    const slides = parsePptBytes(bytes);
    assert.deepEqual(slides.map(slide => slide.index), [1, 2]);
    assert.deepEqual(slides[0].lines, ['第二頁，I’m faithful']);
    assert.deepEqual(slides[1].lines, ['Center body text', '第一頁中文', '保留英文 Don’t', '下一行']);
    assert.deepEqual(slides[1].pageCounter, { current: 1, total: 9 });
    assert.doesNotMatch(slides.flatMap(slide => slide.lines).join('\n'), /不應|number|2026|日期欄位/);
});

test('supports legacy PPT File input without any network conversion', async () => {
    const bytes = fixture();
    const result = await parsePptFile({ name: 'sample.PPT', size: bytes.length, arrayBuffer: async () => bytes.buffer });
    assert.equal(result.length, 2);
});

test('accepts ZIP-based presentations mislabeled with a .ppt extension', async () => {
    const bytes = fflate.zipSync({ 'ppt/slides/slide1.xml': fflate.strToU8('<p:sld><p:sp><p:txBody><a:p><a:r><a:t>正常文字</a:t></a:r></a:p></p:txBody></p:sp></p:sld>') });
    const slides = await parsePptFile({ name: 'renamed.ppt', size: bytes.length, arrayBuffer: async () => bytes.buffer });
    assert.deepEqual(slides[0].lines, ['正常文字']);
});

test('rejects encryption, missing streams, invalid pointers and cyclic edit histories', () => {
    assert.throws(() => parsePptBytes(fixture({ encrypted: true })), /密碼保護/);
    assert.throws(() => parsePptBytes(ole({ WordDocument: Buffer.alloc(32) })), /缺少 PPT/);
    assert.throws(() => parsePptBytes(fixture({ corruptOffset: true })), /結構損壞/);
    assert.throws(() => parsePptBytes(fixture({ cycle: true })), /結構損壞/);
    assert.throws(() => parsePptBytes(fixture({ noText: true })), /找不到可匯入的文字/);
    assert.throws(() => parsePptBytes(Uint8Array.from([1, 2, 3])), /不是有效/);
});

test('rejects invalid suffixes and large files before reading them', async () => {
    const arrayBuffer = async () => { throw new Error('must not read'); };
    await assert.rejects(parsePptFile({ name: 'test.doc', size: 10, arrayBuffer }), /請選擇/);
    await assert.rejects(parsePptFile({ name: 'test.ppt', size: 50 * 1024 * 1024 + 1, arrayBuffer }), /50 MB/);
});
