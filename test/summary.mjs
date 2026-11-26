/**
 * 总结 + 批量锚点：测试
 *
 * 覆盖「一个会话能不能真的总结文档并打上锚点」这条链路：
 *   S：getDigest 的分节与截断
 *   B：setMarks 的原子性、锚点解析、replace / 追加
 *   T：工具面上这两步的完整走法
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { makeDocx, stylesXml, para } from './fixtures.mjs'
import { DocAnchorViewer, buildDocTools } from '../lib/index.js'

/** 临时数据目录。 */
async function withTemp(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'dav-sum-'))
  const prev = process.env.DOC_ANCHOR_VIEWER_DIR
  process.env.DOC_ANCHOR_VIEWER_DIR = dir
  try { await fn(dir) } finally {
    if (prev === undefined) delete process.env.DOC_ANCHOR_VIEWER_DIR
    else process.env.DOC_ANCHOR_VIEWER_DIR = prev
    rmSync(dir, { recursive: true, force: true })
  }
}

/** 一份带四级标题、每节都有正文的样本文档。 */
function sampleDocx(dir, name = 's.docx') {
  const p = join(dir, name)
  writeFileSync(p, makeDocx([
    para('第一章 项目背景', 'H1'),
    para('本项目旨在建设统一的采购管理平台。'),
    para('现有系统分散在多个部门。'),
    para('第二章 技术要求', 'H1'),
    para('平台须支持不少于 500 个并发用户。'),
    para('须提供完整的审计日志。'),
    para('第三章 商务要求', 'H1'),
    para('投标人须具备软件企业认定资质。'),
  ], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
  return p
}

/* ------------------------------------------------------------------ */
/* S：摘要底稿                                                         */
/* ------------------------------------------------------------------ */

test('S1 getDigest 按大纲切节，每节带可用的锚点', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: sampleDocx(dir) })
    const d = await svc.getDigest({ docId: r.doc.id })

    assert.equal(d.ok, true)
    assert.equal(d.sectionCount, 3, '三个一级标题 → 三节')
    assert.deepEqual(d.sections.map((s) => s.title), ['第一章 项目背景', '第二章 技术要求', '第三章 商务要求'])
    // 每节都要带锚点三要素，模型才能直接拿去 set_marks
    for (const s of d.sections) {
      assert.equal(typeof s.anchorIndex, 'number')
      assert.ok(s.norm.length > 0, 'norm 不能为空')
      assert.equal(typeof s.occurrence, 'number')
      assert.ok(s.anchorIndex >= 0)
    }
    // 正文要进节里（不只是标题）
    assert.ok(d.sections[0].text.includes('采购管理平台'), '第一节的正文要带上')
  })
})

test('S2 getDigest 的篇幅上限真的生效，并如实标注截断', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: sampleDocx(dir) })
    const full = await svc.getDigest({ docId: r.doc.id })
    assert.equal(full.sectionCount, 3)

    const capped = await svc.getDigest({ docId: r.doc.id, maxSections: 2 })
    assert.ok(capped.sectionCount <= 2, 'maxSections 要生效：' + capped.sectionCount)

    // 每节字符上限也要生效。下限 120 是刻意的：再小就取不到一句完整的话，
    // 摘要会退化成标题列表，所以参数被夹到 120 而不是照单全收。
    const tight = await svc.getDigest({ docId: r.doc.id, maxCharsPerSection: 5 })
    for (const sec of tight.sections) {
      assert.ok(sec.text.length <= 120, `每节应被夹到下限 120 字以内，实际 ${sec.text.length}`)
    }

    // 真正能触发截断的场景：正文比上限长
    const long = join(dir, 'long.docx')
    writeFileSync(long, makeDocx([
      para('唯一一节', 'H1'),
      para('这是一段很长的正文。'.repeat(60)), // ≈ 600 字
    ], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
    const lr = await svc.openDoc({ filePath: long })
    const lt = await svc.getDigest({ docId: lr.doc.id, maxCharsPerSection: 120 })
    assert.ok(lt.sections[0].text.length <= 120, '长正文必须被截到上限内，实际 ' + lt.sections[0].text.length)
  })
})

test('S3 getDigest 对没有标题样式的文档：整篇当一节，不报错', async () => {
  await withTemp(async (dir) => {
    const p = join(dir, 'plain.docx')
    writeFileSync(p, makeDocx([para('第一段正文'), para('第二段正文')]))
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: p })
    const d = await svc.getDigest({ docId: r.doc.id })
    assert.equal(d.ok, true)
    assert.equal(d.sectionCount, 1, '没有大纲时整篇一节')
    assert.equal(d.sections[0].level, 0)
    assert.ok(d.sections[0].text.includes('第一段正文'))
  })
})

/* ------------------------------------------------------------------ */
/* B：批量锚点                                                         */
/* ------------------------------------------------------------------ */

test('B1 setMarks 批量写入，锚点按精确匹配补齐', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: sampleDocx(dir) })
    const docId = r.doc.id

    const res = await svc.setMarks({
      docId,
      marks: [
        { text: '第二章 技术要求', note: '关键技术条款' },
        { text: '须提供完整的审计日志。', note: '合规要求' },
      ],
    })
    assert.equal(res.ok, true, JSON.stringify(res))
    assert.equal(res.added, 2)
    assert.equal(res.marks[0].anchor.index, 3, '第一章共 3 段，第二章标题在第 3 段')
    assert.equal(res.marks[1].anchor.index, 5)

    const l = await svc.listMarks({ docId })
    assert.equal(l.marks.length, 2)
  })
})

test('B2 setMarks 是原子的：有一条对不上就整组不写', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: sampleDocx(dir) })
    const docId = r.doc.id

    // 先写一条，作为「已有的」
    await svc.setMarks({ docId, marks: [{ text: '第一章 项目背景' }] })
    assert.equal((await svc.listMarks({ docId })).marks.length, 1)

    const bad = await svc.setMarks({
      docId,
      marks: [
        { text: '第二章 技术要求' },          // 这条本来能成功
        { text: '文档里没有这句话' },          // 这条不行
      ],
    })
    assert.equal(bad.ok, false)
    assert.equal(bad.error.code, 'doc/marks-partially-invalid')
    assert.equal(bad.rejected.length, 1)

    const after = await svc.listMarks({ docId })
    assert.equal(after.marks.length, 1, '整组未写入，原有标记不受影响')
  })
})

test('B3 setMarks 的 replace 语义', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: sampleDocx(dir) })
    const docId = r.doc.id

    await svc.setMarks({ docId, marks: [{ text: '第一章 项目背景' }, { text: '第二章 技术要求' }] })
    assert.equal((await svc.listMarks({ docId })).marks.length, 2)

    // 追加
    await svc.setMarks({ docId, marks: [{ text: '第三章 商务要求' }] })
    assert.equal((await svc.listMarks({ docId })).marks.length, 3)

    // 替换
    await svc.setMarks({ docId, replace: true, marks: [{ text: '第三章 商务要求', note: '只剩这条' }] })
    const l = await svc.listMarks({ docId })
    assert.equal(l.marks.length, 1)
    assert.equal(l.marks[0].note, '只剩这条')
  })
})

test('B4 setMarks 用 getDigest 的 anchor 精确指定重复标题的第 N 处', async () => {
  await withTemp(async (dir) => {
    const p = join(dir, 'dup.docx')
    writeFileSync(p, makeDocx([
      para('附件一', 'H1'),
      para('第一处正文'),
      para('附件一', 'H1'),
      para('第二处正文'),
    ], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: p })
    const d = await svc.getDigest({ docId: r.doc.id })
    assert.equal(d.sections.length, 2)
    // 两节同名，但 occurrence 不同
    assert.equal(d.sections[0].occurrence, 0)
    assert.equal(d.sections[1].occurrence, 1)

    const res = await svc.setMarks({
      docId: r.doc.id,
      marks: [{ text: '附件一', note: '指向第二处', index: d.sections[1].anchorIndex, norm: d.sections[1].norm, occurrence: d.sections[1].occurrence }],
    })
    assert.equal(res.ok, true, JSON.stringify(res))
    assert.equal(res.marks[0].anchor.index, 2, '要锚到第二处（第 2 段），不是第一处')
    assert.equal(res.marks[0].anchor.occurrence, 1)
  })
})

test('B5 setMarks 的输入校验', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: sampleDocx(dir) })
    const docId = r.doc.id

    const a = await svc.setMarks({ docId, marks: 'not-an-array' })
    assert.equal(a.ok, false)
    assert.equal(a.error.code, 'doc/invalid-input')

    const b = await svc.setMarks({ docId, marks: [] })
    assert.equal(b.ok, true, '空数组是合法的 no-op')
    assert.equal(b.added, 0)

    const c = await svc.setMarks({ docId, marks: Array.from({ length: 201 }, () => ({ text: 'x' })) })
    assert.equal(c.ok, false, '超过 200 条要拒绝')

    const d = await svc.setMarks({ docId: 'nope', marks: [{ text: 'x' }] })
    assert.equal(d.ok, false)
  })
})

/* ------------------------------------------------------------------ */
/* T：工具面上的完整走法                                               */
/* ------------------------------------------------------------------ */

test('T1 「总结 → 打锚点」在工具面上是一步到位的', async () => {
  await withTemp(async (dir) => {
    const reg = []
    buildDocTools(
      { get: (k) => (k === 'tools' ? { register: (d) => { reg.push(d); return () => {} } } : undefined) },
      new DocAnchorViewer({}, {}),
    )
    const def = reg[0]
    const show = (v) => def.output.render({}, v).map((b) => b.text).join('\n')

    const open = await def.execute({ action: 'open', filePath: sampleDocx(dir) })
    const docId = open.doc.id

    // 1) 取底稿
    const digest = await def.execute({ action: 'get_digest', docId })
    const digestText = show(digest)
    assert.ok(digestText.includes('第一章 项目背景'))
    assert.ok(digestText.includes('本项目旨在建设'), '摘要要带正文，不是只有标题')

    // 2) 直接拿 digest 的锚点批量落盘（模拟模型总结后回填）
    const marks = digest.sections.map((s) => ({
      text: s.title,
      note: '本节要点',
      index: s.anchorIndex,
      norm: s.norm,
      occurrence: s.occurrence,
    }))
    const res = await def.execute({ action: 'set_marks', docId, marks })
    assert.equal(res.ok, true, JSON.stringify(res))
    assert.equal(res.added, 3)

    const out = show(res)
    assert.ok(out.includes('已写入 3 条标记'), '要明确说写了几条：' + out.slice(0, 120))

    // 3) 面板侧能读到
    const list = await def.execute({ action: 'list_marks', docId })
    assert.equal(list.marks.length, 3)
    assert.deepEqual(list.marks.map((m) => m.anchor.index), [0, 3, 6])
  })
})
