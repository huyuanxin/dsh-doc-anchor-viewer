/**
 * 文档锚点视图 · ZIP 容器安全：测试
 *
 * 覆盖一条真实的攻击面：**docx 就是 ZIP，而 ZIP 能被造出来骗过「文件大小」检查**。
 *
 *   · Z1：正常 docx 照常解析（回归：改坏了要红）
 *   · Z2：只解需要的条目 —— 不在白名单里的条目连解压都不做
 *   · Z3：压缩炸弹被单条目上限挡住（不 OOM）
 *   · Z4：多个条目的累加被总预算挡住
 *   · Z5：伪造的中央目录不会变成天书错误（RangeError 之类）
 *   · Z6：端到端——openDoc 遇到炸弹返回可读的业务失败，宿主不崩
 *
 * 不依赖真实招标文件：样本由 test/fixtures.mjs 现场造。
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { makeDocx, stylesXml, para, makeZip } from './fixtures.mjs'
import { readZipBuffer, MAX_ENTRY_BYTES, MAX_TOTAL_BYTES } from '../lib/ooxml.js'
import { DocAnchorViewer } from '../lib/index.js'

/** 临时数据目录。 */
async function withTemp(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'dav-zip-'))
  const prev = process.env.DOC_ANCHOR_VIEWER_DIR
  process.env.DOC_ANCHOR_VIEWER_DIR = dir
  try { await fn(dir) } finally {
    if (prev === undefined) delete process.env.DOC_ANCHOR_VIEWER_DIR
    else process.env.DOC_ANCHOR_VIEWER_DIR = prev
    rmSync(dir, { recursive: true, force: true })
  }
}

/** 解析一份 docx 真正需要的条目（与 lib/index.js 的 DOC_PARTS 一致）。 */
const DOC_PARTS = ['word/document.xml', 'word/styles.xml']

/**
 * 造一个「高压缩比」的条目：1MB 的全零 → deflate 后约 1KB。
 * 这是压缩炸弹的最小复现单元。
 * @param {number} size 解压后的大小
 * @returns {Buffer}
 */
function bombPayload(size) {
  return Buffer.alloc(size, 0)
}

/* ------------------------------------------------------------------ */
/* Z1：正常路径                                                         */
/* ------------------------------------------------------------------ */

test('Z1 正常 docx 照常解出需要的条目（改了只解白名单也要能跑）', () => {
  const buf = makeDocx([para('第一章', 'H1'), para('正文')], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }]))
  const entries = readZipBuffer(buf, { only: DOC_PARTS })
  assert.ok(entries.has('word/document.xml'), '要拿到 document.xml')
  assert.ok(entries.has('word/styles.xml'), '要拿到 styles.xml')
  assert.match(entries.get('word/document.xml').toString('utf8'), /第一章/)
})

/* ------------------------------------------------------------------ */
/* Z2：只解需要的条目                                                   */
/* ------------------------------------------------------------------ */

test('Z2 不在白名单里的条目连解压都不做（图片/嵌入对象不该进内存）', () => {
  const buf = makeZip([
    { name: 'word/document.xml', data: '<?xml version="1.0"?><w:document xmlns:w="x"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p></w:body></w:document>' },
    // 一个能膨胀到 1MB 的条目：白名单外就该被完全无视
    { name: 'word/media/bomb.bin', data: bombPayload(1024 * 1024) },
  ])

  const entries = readZipBuffer(buf, { only: DOC_PARTS })
  assert.ok(entries.has('word/document.xml'))
  assert.equal(entries.has('word/media/bomb.bin'), false, '白名单外的条目不该出现在结果里')
  assert.equal(entries.size, 1, '只该留下真正需要的条目')
})

test('Z2b 不给白名单时，炸弹条目会被解压上限挡住（说明上限确实在生效）', () => {
  const buf = makeZip([
    { name: 'word/document.xml', data: '<?xml version="1.0"?><w:document xmlns:w="x"><w:body/></w:document>' },
    { name: 'word/media/bomb.bin', data: bombPayload(1024 * 1024) },
  ])
  assert.throws(
    () => readZipBuffer(buf, { maxOutputLength: 64 * 1024 }),
    (e) => String(e?.message ?? '').startsWith('ZIP_ENTRY_TOO_LARGE'),
    '不设白名单 + 小上限 → 必须被挡住',
  )
})

/* ------------------------------------------------------------------ */
/* Z3：压缩炸弹                                                         */
/* ------------------------------------------------------------------ */

test('Z3 压缩炸弹被单条目上限挡住（不会把宿主撑爆）', () => {
  const buf = makeZip([
    { name: 'word/document.xml', data: '<?xml version="1.0"?><w:document xmlns:w="x"><w:body/></w:document>' },
    { name: 'word/media/bomb.bin', data: bombPayload(4 * 1024 * 1024) },
  ])
  assert.throws(
    () => readZipBuffer(buf, { only: ['word/media/bomb.bin'], maxOutputLength: 1024 * 1024 }),
    (e) => {
      const m = String(e?.message ?? '')
      assert.ok(m.startsWith('ZIP_ENTRY_TOO_LARGE'), '要抛自己的错误码，实际：' + m)
      assert.match(m, /bomb\.bin/, '错误里要带条目名（给开发者定位用）')
      return true
    },
  )
})

test('Z3b 默认上限是 64MB/条目、256MB/总量（写死的常量，改了要有人知道）', () => {
  assert.equal(MAX_ENTRY_BYTES, 64 * 1024 * 1024)
  assert.equal(MAX_TOTAL_BYTES, 256 * 1024 * 1024)
})

test('Z3c 上限选项传了非法值也不能把上限**废掉**（NaN 会让比较恒为 false）', () => {
  // 关键在于「非法值要退回**默认上限**」，而不是变成「无上限」。
  // 所以必须用一个**超过默认上限**的条目来证伪：若上限被废掉，它就不会抛错。
  const buf = makeZip([
    { name: 'word/media/bomb.bin', data: bombPayload(MAX_ENTRY_BYTES + 1024 * 1024) },
  ])
  for (const bad of [NaN, 0, -1, undefined, null, 'abc', {}]) {
    assert.throws(
      () => readZipBuffer(buf, { maxOutputLength: bad }),
      (e) => String(e?.message ?? '').startsWith('ZIP_ENTRY_TOO_LARGE'),
      'maxOutputLength=' + String(bad) + ' 也要退回默认上限（64MB），不能让上限失效',
    )
  }
})

/* ------------------------------------------------------------------ */
/* Z4：总预算                                                           */
/* ------------------------------------------------------------------ */

test('Z4 单条目都没超限，但累加超了总预算 → 也要挡住', () => {
  const entries = []
  for (let i = 0; i < 8; i++) {
    entries.push({ name: 'word/x' + i + '.bin', data: bombPayload(512 * 1024) })
  }
  const buf = makeZip(entries)
  assert.throws(
    () => readZipBuffer(buf, { maxOutputLength: 1024 * 1024, maxTotalBytes: 1024 * 1024 }),
    (e) => String(e?.message ?? '').startsWith('ZIP_TOTAL_TOO_LARGE'),
    '8×512KB > 1MB 总量预算 → 必须被挡住',
  )
})

/* ------------------------------------------------------------------ */
/* Z5：恶意中央目录                                                     */
/* ------------------------------------------------------------------ */

test('Z5 伪造的中央目录不会变成天书错误', () => {
  const buf = makeZip([{ name: 'word/document.xml', data: '<?xml version="1.0"?><w:document/>' }])

  // ① 把中央目录的偏移量改成越界值
  const bad = Buffer.from(buf)
  const eocd = bad.length - 22
  bad.writeUInt32LE(0xfffffff0, eocd + 16)
  assert.throws(
    () => readZipBuffer(bad),
    (e) => String(e?.message ?? '').startsWith('ZIP_BAD_CDH'),
    '偏移越界要翻成 ZIP_BAD_CDH，而不是 RangeError',
  )

  // ② 截断文件（中央目录指向已不存在的数据）
  const cut = buf.subarray(0, buf.length - 40)
  assert.throws(
    () => readZipBuffer(cut),
    (e) => /^(ZIP_BAD_CDH|ZIP_TRUNCATED|ZIP_BAD_LFH|ZIP_EOCD_NOT_FOUND)$/.test(String(e?.message ?? '')),
    '截断文件要翻成可识别的 ZIP 错误码',
  )

  // ③ 根本不是 ZIP
  assert.throws(
    () => readZipBuffer(Buffer.from('this is not a zip file at all')),
    (e) => String(e?.message ?? '').startsWith('ZIP_EOCD_NOT_FOUND'),
  )
})

/* ------------------------------------------------------------------ */
/* Z6：端到端                                                           */
/* ------------------------------------------------------------------ */

test('Z6 端到端：openDoc 遇到压缩炸弹 → 可读的业务失败，宿主不崩', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const p = join(dir, 'bomb.docx')
    // 一个真实体积很小（几十 KB）、但能膨胀到 65MB 的条目
    writeFileSync(p, makeZip([
      { name: 'word/document.xml', data: bombPayload(65 * 1024 * 1024) },
    ]))

    const r = await svc.openDoc({ filePath: p })
    assert.equal(r.ok, false, '必须失败，不能崩')
    assert.equal(r.error.code, 'doc/parse-failed')
    assert.match(r.error.message, /压缩炸弹|过大/, '要给出人话：' + r.error.message)
    assert.doesNotMatch(r.error.message, /ZIP_/, '不要把内部错误码端给用户')
  })
})

test('Z6b 正常文档走完整链路不受影响（回归）', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const p = join(dir, 'ok.docx')
    writeFileSync(p, makeDocx(
      [para('第一章', 'H1'), para('正文甲'), para('第二章', 'H1'), para('正文乙')],
      stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }]),
    ))
    const r = await svc.openDoc({ filePath: p })
    assert.equal(r.ok, true, '正常文档必须照常登记')
    assert.equal(r.doc.stats.paragraphs, 4)
    assert.equal(r.doc.stats.headings, 2)
  })
})
