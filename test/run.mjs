/**
 * 文档锚点视图 · 测试
 *
 * 覆盖四层：
 *   P：解析（标题层级、样式继承、同文计数、域代码/删除文本不计入）
 *   A：锚点（Host 算出的锚点 ↔ 客户端匹配规则，杜绝 includes 误匹配）
 *   H：Host 服务（登记、大纲、按锚点读原文、标记增删、路径安全）
 *   C：契约（REMOTE_METHODS 与 client descriptors 逐位一致、客户端可解析）
 *
 * 不依赖真实招标文件：样本由 test/fixtures.mjs 现场造，测的是结构行为本身。
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtempSync, writeFileSync, rmSync, existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'

import { makeDocx, stylesXml, para } from './fixtures.mjs'
import { parseParagraphs, parseStyles, buildOutline, occurrenceMap, headingLevel, normText } from '../lib/docparse.js'
import { readZipBuffer } from '../lib/ooxml.js'
import { DocAnchorViewer, REMOTE_METHODS, apply } from '../lib/index.js'

/* ------------------------------------------------------------------ */
/* 公共：临时数据目录                                                  */
/* ------------------------------------------------------------------ */

/**
 * 每个用例一个独立数据目录（环境变量覆盖，避免污染真实数据）。
 * @param {Function} fn
 */
async function withTemp(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'dav-test-'))
  const prev = process.env.DOC_ANCHOR_VIEWER_DIR
  process.env.DOC_ANCHOR_VIEWER_DIR = dir
  try {
    await fn(dir)
  } finally {
    if (prev === undefined) delete process.env.DOC_ANCHOR_VIEWER_DIR
    else process.env.DOC_ANCHOR_VIEWER_DIR = prev
    rmSync(dir, { recursive: true, force: true })
  }
}

/** 造一个服务实例（不需要 Cordis：构造器只用到 ctx 存引用）。 */
function makeService() {
  return new DocAnchorViewer({}, { dataDir: '', maxDocxMb: 30 })
}

/** 把 docx 写到临时目录并返回路径。 */
function writeDocx(dir, name, buf) {
  const p = join(dir, name)
  writeFileSync(p, buf)
  return p
}

/* ------------------------------------------------------------------ */
/* P：解析                                                             */
/* ------------------------------------------------------------------ */

test('P1 解析标题层级：outlineLvl 与样式名两条判据都要认', () => {
  const styles = parseStyles(stylesXml([
    { id: 'Heading1', name: 'heading 1', outlineLvl: 0 },
    { id: 'Heading2', name: 'heading 2', outlineLvl: 1 },
    { id: 'MyTitle', name: '标题 3' },
    { id: 'Body', name: 'Normal' },
  ]))
  assert.equal(styles.get('Heading1').outlineLvl, 0)
  assert.equal(styles.get('Heading2').outlineLvl, 1)

  assert.deepEqual(headingLevel({ outlineLvl: 0, name: 'heading 1', styleId: 'Heading1' }), { heading: true, level: 1 })
  assert.deepEqual(headingLevel({ outlineLvl: 2, name: '', styleId: '' }), { heading: true, level: 3 })
  // 只有名字（中文 Word 常见）
  assert.deepEqual(headingLevel({ outlineLvl: null, name: '标题 3', styleId: 'MyTitle' }), { heading: true, level: 3 })
  // 正文：两条都不成立 → 不猜
  assert.deepEqual(headingLevel({ outlineLvl: null, name: 'Normal', styleId: 'Body' }), { heading: false, level: 0 })
})

test('P2 样式继承：outlineLvl 写在 basedOn 的祖先上也要认出来', () => {
  const styles = parseStyles(stylesXml([
    { id: 'Base', name: 'base', outlineLvl: 1 },
    { id: 'Child', name: 'child', basedOn: 'Base' },
  ]))
  const xml = '<w:document><w:body>'
    + para('继承出来的二级标题', 'Child')
    + para('正文一段', 'Child2')
    + '</w:body></w:document>'
  const ps = parseParagraphs(xml, styles)
  assert.equal(ps.length, 2)
  assert.equal(ps[0].heading, true)
  assert.equal(ps[0].level, 2, 'basedOn 的 outlineLvl=1 → 2 级')
  assert.equal(ps[1].heading, false, '没定义样式的段落不猜成标题')
})

test('P3 文本收集：域代码与修订删除不计入，正文正常解码', () => {
  const xml = '<w:document><w:body>'
    + '<w:p><w:r><w:t>第</w:t></w:r>'
    + '<w:r><w:instrText>PAGEREF _Toc1</w:instrText></w:r>'
    + '<w:r><w:t>一</w:t></w:r></w:p>'
    + '<w:p><w:r><w:delText>被删掉的旧文字</w:delText></w:r><w:r><w:t>保留的</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>A&amp;B &lt;C&gt;</w:t></w:r></w:p>'
    + '</w:body></w:document>'
  const ps = parseParagraphs(xml, new Map())
  assert.equal(ps.length, 3)
  assert.equal(ps[0].text, '第一', '域代码不进正文')
  assert.equal(ps[1].text, '保留的', '删除修订不进正文')
  assert.equal(ps[2].text, 'A&B <C>', 'XML 实体正确解码')
})

test('P4 同文计数：occurrence 按全文统计（不是只在大纲里数）', () => {
  const styles = parseStyles(stylesXml([
    { id: 'H1', name: 'heading 1', outlineLvl: 0 },
  ]))
  const xml = '<w:document><w:body>'
    + para('1.1', 'H1')        // 标题，第 0 次
    + para('正文', '')
    + para('1.1', '')          // 正文，第 1 次
    + para('1.1', 'H1')        // 又是标题，第 2 次
    + '</w:body></w:document>'
  const ps = parseParagraphs(xml, styles)
  const occ = occurrenceMap(ps)
  assert.equal(occ.get(0), 0)
  assert.equal(occ.get(2), 1)
  assert.equal(occ.get(3), 2)

  const outline = buildOutline(ps)
  assert.equal(outline.length, 2)
  // 关键：两个标题各自的 occurrence 必须是 0 和 2（跨全文），
  // 若只在大纲里数就会是 0 和 1，第二处会锚错。
  assert.equal(outline[0].occurrence, 0)
  assert.equal(outline[1].occurrence, 2)
})

test('P5 归一化：空白差异不影响匹配键', () => {
  assert.equal(normText('格 式  要求\n'), '格式要求')
  assert.equal(normText('\u3000金额'), '金额')
  const xml = '<w:document><w:body>'
    + '<w:p><w:r><w:t>格 式 要 求</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>格式要求</w:t></w:r></w:p>'
    + '</w:body></w:document>'
  const ps = parseParagraphs(xml, new Map())
  assert.equal(ps[0].norm, ps[1].norm, '两段的归一化键相同 → 靠 occurrence 区分')
})

/* ------------------------------------------------------------------ */
/* A：锚点匹配规则                                                     */
/* ------------------------------------------------------------------ */

test('A1 锚点匹配只用精确相等：短表头不会被长段「包含」而误锚', async () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  // 取出 stampAnchors 函数体（到下一个顶层 function 之前）
  const start = src.indexOf('function stampAnchors')
  assert.ok(start > 0, 'client.js 里应该有 stampAnchors')
  const end = src.indexOf('\n    function ', start + 10)
  const body = src.slice(start, end > 0 ? end : start + 4000)
  assert.ok(/normAnchor\(p\.textContent\)/.test(body), '必须按段落的归一化全文取键')
  assert.ok(!/\.includes\(/.test(body), 'stampAnchors 里不允许出现 includes（会误匹配短文本）')
})

test('A2 锚点键的归一化实现与 Host 完全一致', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const m = /function normAnchor\(s\) \{\s*return (String\(s \?\? ''\)\.replace\(\/\\s\+\/g, ''\)\.trim\(\))/.exec(src)
  assert.ok(m, '客户端 normAnchor 必须与 Host normText 用同一条正则')
  assert.equal(normText(' a b\tc\n'), 'abc')
})

/* ------------------------------------------------------------------ */
/* H：Host 服务                                                        */
/* ------------------------------------------------------------------ */

test('H1 登记 + 大纲 + 按锚点读原文', async () => {
  await withTemp(async (dir) => {
    const buf = makeDocx([
      para('第一章 总则', 'H1'),
      para('本章说明适用范围。'),
      para('1.1 适用范围', 'H2'),
      para('适用于所有投标人。'),
      para('1.1 适用范围', 'H2'),
      para('重复标题下的第二处正文。'),
    ], stylesXml([
      { id: 'H1', name: 'heading 1', outlineLvl: 0 },
      { id: 'H2', name: 'heading 2', outlineLvl: 1 },
    ]))
    const p = writeDocx(dir, 'a.docx', buf)

    const svc = makeService()
    const r = await svc.openDoc({ filePath: p })
    assert.equal(r.ok, true, JSON.stringify(r))
    assert.equal(r.doc.stats.paragraphs, 6)
    assert.equal(r.doc.stats.headings, 3)

    const o = await svc.getOutline({ docId: r.doc.id })
    assert.equal(o.ok, true)
    assert.deepEqual(o.outline.map((x) => x.title), ['第一章 总则', '1.1 适用范围', '1.1 适用范围'])
    // 同名标题的第二处 occurrence 必须是 1（全文第 1 次出现之后）
    assert.equal(o.outline[1].occurrence, 0)
    assert.equal(o.outline[2].occurrence, 1)

    // 按 index 读
    const s1 = await svc.readSection({ docId: r.doc.id, index: 3, before: 1, after: 1 })
    assert.equal(s1.ok, true)
    assert.equal(s1.anchorIndex, 3)
    assert.deepEqual(s1.paragraphs.map((x) => x.text), ['1.1 适用范围', '适用于所有投标人。', '1.1 适用范围'])

    // 按 query 读（默认落到第一个同文处）
    const s2 = await svc.readSection({ docId: r.doc.id, query: '不存在的锚点文字' })
    assert.equal(s2.ok, false)
    assert.equal(s2.error.code, 'doc/anchor-not-found')
  })
})

test('H2 标记：加、列、删；锚点缺失时由 Host 用精确匹配补齐', async () => {
  await withTemp(async (dir) => {
    const buf = makeDocx([
      para('联系人', ''),
      para('联系人信息表如下：姓名、电话、地址、邮编、传真。'),
    ])
    const p = writeDocx(dir, 'b.docx', buf)
    const svc = makeService()
    const r = await svc.openDoc({ filePath: p })
    assert.equal(r.ok, true)

    // 只给文字、不给锚点 → Host 按全文精确匹配（不是 includes）
    const m = await svc.addMark({ docId: r.doc.id, text: '联系人', note: '看这里' })
    assert.equal(m.ok, true)
    assert.equal(m.mark.anchor.index, 0, '应该锚到第 0 段那个**恰好等于**「联系人」的段落')
    assert.equal(m.mark.anchor.norm, '联系人')

    const l = await svc.listMarks({ docId: r.doc.id })
    assert.equal(l.marks.length, 1)

    const rm = await svc.removeMark({ docId: r.doc.id, markId: m.mark.id })
    assert.equal(rm.ok, true)
    const l2 = await svc.listMarks({ docId: r.doc.id })
    assert.equal(l2.marks.length, 0)

    // 再删一次：明确报「不存在」，不静默成功
    const rm2 = await svc.removeMark({ docId: r.doc.id, markId: m.mark.id })
    assert.equal(rm2.ok, false)
    assert.equal(rm2.error.code, 'doc/mark-not-found')
  })
})

test('H3 同一份文件重复登记保留标记；改过之后是新记录', async () => {
  await withTemp(async (dir) => {
    const p = join(dir, 'c.docx')
    writeFileSync(p, makeDocx([para('标题', 'H1')], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
    const svc = makeService()
    const a = await svc.openDoc({ filePath: p })
    await svc.addMark({ docId: a.doc.id, text: '标题' })
    const b = await svc.openDoc({ filePath: p })
    assert.equal(b.doc.id, a.doc.id, '同一路径同一 mtime → 同一条记录')
    assert.equal(b.doc.markCount, 1, '重复登记不能把用户标记冲掉')

    // 改内容（mtime 变）→ 新记录，旧标记不会错挂到新内容上
    await new Promise((r) => setTimeout(r, 20))
    writeFileSync(p, makeDocx([para('标题改了', 'H1')], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
    const c = await svc.openDoc({ filePath: p })
    assert.notEqual(c.doc.id, b.doc.id, '内容变了应该是新记录')
    assert.equal(c.doc.markCount, 0)
  })
})

test('H4 base64 上传：来源无路径也能登记，且重复上传命中同一条', async () => {
  await withTemp(async () => {
    const buf = makeDocx([para('上传的文档', 'H1')], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }]))
    const svc = makeService()
    const a = await svc.openDoc({ base64: buf.toString('base64'), fileName: 'upload.docx' })
    assert.equal(a.ok, true, JSON.stringify(a))
    assert.equal(a.doc.stats.paragraphs, 1)
    const b = await svc.openDoc({ base64: buf.toString('base64'), fileName: 'upload.docx' })
    assert.equal(b.doc.id, a.doc.id, '同样的内容 + 同样的文件名 → 同一条记录')
  })
})

test('H5 错误路径：文件不存在 / 不是 docx / 空 base64 都给出可读原因', async () => {
  await withTemp(async (dir) => {
    const svc = makeService()
    const a = await svc.openDoc({ filePath: join(dir, 'nope.docx') })
    assert.equal(a.ok, false)
    assert.equal(a.error.code, 'doc/file-not-found')

    const txt = writeDocx(dir, 'x.txt', Buffer.from('hello'))
    const b = await svc.openDoc({ filePath: txt })
    assert.equal(b.ok, false)
    assert.equal(b.error.code, 'doc/unsupported-format')

    const bad = writeDocx(dir, 'bad.docx', Buffer.from('not a zip'))
    const c = await svc.openDoc({ filePath: bad })
    assert.equal(c.ok, false)
    assert.ok(c.error.code === 'doc/not-a-docx' || c.error.code === 'doc/parse-failed')

    const d = await svc.openDoc({})
    assert.equal(d.ok, false)
    assert.equal(d.error.code, 'doc/invalid-input')
  })
})

test('H6 docId 路径穿越被拒绝', async () => {
  await withTemp(async () => {
    const svc = makeService()
    const r = await svc.getDoc({ docId: '../../etc/passwd' })
    assert.equal(r.ok, false, '非法 docId 不能读到库外文件')
    assert.equal(r.error.code, 'doc/invalid-input')
  })
})

test('H7 删文档只删登记，不删源文件', async () => {
  await withTemp(async (dir) => {
    const p = writeDocx(dir, 'd.docx', makeDocx([para('内容')]))
    const svc = makeService()
    const a = await svc.openDoc({ filePath: p })
    const rm = await svc.removeDoc({ docId: a.doc.id })
    assert.equal(rm.ok, true)
    assert.equal(existsSync(p), true, '源文件必须还在')
    const list = await svc.listDocs()
    assert.equal(list.docs.length, 0)
  })
})

/* ------------------------------------------------------------------ */
/* C：契约                                                             */
/* ------------------------------------------------------------------ */

test('C1 REMOTE_METHODS 与 client descriptors 逐位一致', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const block = src.slice(src.indexOf('descriptors: ['), src.indexOf('],', src.indexOf('descriptors: [')))
  const names = [...block.matchAll(/desc\('([A-Za-z0-9_]+)'/g)].map((m) => m[1])
  assert.deepEqual(names, REMOTE_METHODS, '两边方法名与顺序必须完全一致（错一位全挂）')
})

test('C1b typertRemote 绑定齐全：缺它 = 整个命名空间 HTTP 404', () => {
  // 这条守的是一个**极难从现象反推**的契约：
  //   Gateway 先读 service 实例上的 typertRemote（service/serviceKey/namespace），
  //   再读 prototype 上的方法列表。只写后者的话，命名空间根本不进路由表，
  //   客户端收到的不是「方法不存在」而是整条路由 404——第一反应会去
  //   怀疑插件没装、宿主没重载，而真正的修法只是补一个对象。
  // 方法列表是 apply() 写上去的（构造函数只负责实例上的绑定），
  // 所以这里要走真实的 apply 路径，而不是裸 new。
  let svc = null
  apply(
    { provide: () => {}, set: (k, v) => { if (k === 'docAnchorViewer') svc = v }, inject: () => {}, effect: () => {} },
    {},
  )
  assert.ok(svc, 'apply 必须把服务挂到 ctx 上')
  const b = svc.typertRemote
  assert.ok(b, '服务实例上必须有 typertRemote（Gateway 的服务发现入口）')
  assert.equal(b.service, svc, 'binding.service 必须指向实例本身')
  assert.equal(b.serviceKey, 'docAnchorViewer', 'binding.serviceKey 必须是 Cordis 服务键')
  assert.equal(b.namespace, 'docAnchor', 'binding.namespace 必须与客户端 descriptors 的 namespace 一致')

  const KEY = '@deepseek-ai/dsh-typert-protocol/remote-methods'
  const markers = Object.getPrototypeOf(svc)[KEY]
  assert.ok(markers, 'prototype 上必须有方法列表')
  assert.deepEqual(markers.methods.map((m) => m.method), REMOTE_METHODS, '方法列表要与 REMOTE_METHODS 一致')
})

test('C2 带参方法的 wire 名必须是 request（等于 Host 形参名）', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const block = src.slice(src.indexOf('descriptors: ['), src.indexOf('],', src.indexOf('descriptors: [')))
  for (const m of block.matchAll(/desc\('([A-Za-z0-9_]+)'(?:,\s*'([^']*)')?\)/g)) {
    const [, name, wire] = m
    const host = DocAnchorViewer.prototype[name]
    if (typeof host === 'function' && host.length > 0) {
      assert.equal(wire, 'request', `${name} 在 Host 上是带参方法，wire 必须是 request`)
    } else {
      assert.equal(wire, undefined, `${name} 在 Host 上无参，descriptor 也不能带 wire`)
    }
  }
})

test('C3 客户端模块能被解析且注册了 main / sidebar.panellist', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  assert.ok(src.includes("id: 'dsh-doc-anchor-viewer'"), '模块 id 必须是包名')
  assert.ok(src.includes("name: 'main'"), '必须注册 main 面板')
  assert.ok(src.includes("id: PANEL_ID"), '必须注册侧栏入口')
  assert.ok(src.includes('DOCX_PREVIEW_IIFE_BEGIN') && src.includes('DOCX_PREVIEW_IIFE_END'), 'docx-preview 标记对必须还在')
  assert.ok(src.length > 200_000, '内联的 docx-preview 应该在（约 170KB）')
})

test('C4 package.json 的 files 覆盖全部运行时文件', () => {
  const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'))
  const files = pkg.files.join(' ')
  assert.ok(files.includes('lib/*.js'), 'lib 必须被打包')
  assert.ok(files.includes('cordis.patch.yml'))
  assert.ok(files.includes('locale/*.json'))
  assert.ok(pkg.exports['./client'], '必须导出 client 半边')
  assert.ok(pkg.dsh.client, '必须声明 dsh.client')
  assert.ok(pkg.dsh.bundle.patch, '必须声明 bundle patch')
})
