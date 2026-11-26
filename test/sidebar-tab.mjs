/**
 * 侧栏标签页：测试
 *
 * 覆盖「同一份 docx 既能在大面板里看，也能在侧栏标签页里看」这条链路：
 *   S：按路径查已登记文档（getDocByPath）——侧栏 body 只有路径，没有 docId
 *   A：从 dsh-resource 地址解出绝对路径——解错就永远「未登记」
 *   R：标签页的两种形态（已登记带大纲 / 未登记只预览）
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtempSync, writeFileSync, rmSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import vm from 'node:vm'

import { makeDocx, stylesXml, para } from './fixtures.mjs'
import { DocAnchorViewer, REMOTE_METHODS } from '../lib/index.js'

/** 临时数据目录。 */
async function withTemp(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'dav-tab-'))
  const prev = process.env.DOC_ANCHOR_VIEWER_DIR
  process.env.DOC_ANCHOR_VIEWER_DIR = dir
  try { await fn(dir) } finally {
    if (prev === undefined) delete process.env.DOC_ANCHOR_VIEWER_DIR
    else process.env.DOC_ANCHOR_VIEWER_DIR = prev
    rmSync(dir, { recursive: true, force: true })
  }
}

/* ------------------------------------------------------------------ */
/* S：按路径查文档                                                     */
/* ------------------------------------------------------------------ */

test('S1 getDocByPath 命中已登记的文档，带回完整大纲', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const p = join(dir, 'a.docx')
    writeFileSync(p, makeDocx(
      [para('第一章 总则', 'H1'), para('正文甲'), para('第二章 要求', 'H1'), para('正文乙')],
      stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }]),
    ))
    const open = await svc.openDoc({ filePath: p })
    assert.equal(open.ok, true)

    const hit = await svc.getDocByPath({ filePath: p })
    assert.equal(hit.ok, true)
    assert.equal(hit.found, true)
    assert.equal(hit.doc.id, open.doc.id, '要拿到同一条记录')
    assert.equal(hit.doc.outline.length, 2, '要带回大纲（侧栏靠它渲染）')
    assert.deepEqual(hit.doc.outline.map((o) => o.title), ['第一章 总则', '第二章 要求'])
  })
})

test('S2 路径比较忽略大小写与分隔符（Windows 用户粘贴的路径花样很多）', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const p = join(dir, 'b.docx')
    writeFileSync(p, makeDocx([para('标题', 'H1')], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
    await svc.openDoc({ filePath: p })

    // 反斜杠 → 正斜杠
    const slash = await svc.getDocByPath({ filePath: p.replace(/\\/g, '/') })
    assert.equal(slash.found, true, '正斜杠写法要能命中')

    // 末尾多一个分隔符
    const trailing = await svc.getDocByPath({ filePath: p + '\\' })
    assert.equal(trailing.found, true, '末尾分隔符要能容忍')

    // 带引号（用户从资源管理器复制）
    const quoted = await svc.getDocByPath({ filePath: '"' + p + '"' })
    assert.equal(quoted.found, true, '两端引号要剥掉')

    // 大小写：只在 Windows 上折平，而且这正是**故意的**——POSIX 的 a.docx 与 A.DOCX
    // 是两个真实存在的不同文件，折平会把两条记录串成一个，比"命中不了"更糟。
    // 所以这条断言必须按平台分支，否则在 Linux CI 上测的是别人的文件系统语义。
    const upper = await svc.getDocByPath({ filePath: p.toUpperCase() })
    if (process.platform === 'win32') {
      assert.equal(upper.found, true, 'Windows 上大小写不同要能命中')
    } else {
      assert.equal(upper.found, false, 'POSIX 上大小写不同就是另一个文件，不能命中')
    }
  })
})

/**
 * S2 的 Windows 分支在 Linux CI 上永远走不到，而它恰恰是 Windows 用户的**主路径**。
 * `getDocByPath` 是在调用时读 `process.platform` 的，所以这里把它临时换成 win32，
 * 让"大小写折平"这件事在 ubuntu-latest 上也能被真实执行到——不是伪造结论，
 * 而是给同一个实现换一个平台输入。
 */
test('S2b Windows 分支：换平台也要把大小写折平（Linux CI 上同样验证得到）', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const p = join(dir, 'c.docx')
    writeFileSync(p, makeDocx([para('标题', 'H1')], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
    const open = await svc.openDoc({ filePath: p })
    assert.equal(open.ok, true)

    const real = process.platform
    Object.defineProperty(process, 'platform', { value: 'win32', configurable: true })
    try {
      assert.equal(process.platform, 'win32', '前提：平台已经换掉了')
      const upper = await svc.getDocByPath({ filePath: p.toUpperCase() })
      assert.equal(upper.found, true, 'Windows 上大小写不同要能命中')
      assert.equal(upper.doc.id, open.doc.id, '命中的必须是同一条记录')
      // 分隔符与大小写叠加的写法：Windows 用户从各种地方复制出来的样子
      const mixed = await svc.getDocByPath({ filePath: p.toUpperCase().replace(/\//g, '\\') + '\\' })
      assert.equal(mixed.found, true, '反斜杠 + 大写 + 末尾分隔符要一起容忍')
    } finally {
      Object.defineProperty(process, 'platform', { value: real, configurable: true })
    }
    assert.equal(process.platform, real, '平台必须还原，不能漏给后面的用例')
  })
})

test('S3 未登记的文档返回 found:false，不报错（侧栏要退化成普通预览）', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.getDocByPath({ filePath: join(dir, 'never-opened.docx') })
    assert.equal(r.ok, true, '查不到不是错误——这是正常分支')
    assert.equal(r.found, false)
    assert.equal(r.doc, undefined)
  })
})

test('S4 getDocByPath 已登记在 REMOTE_METHODS 与 client descriptors 里', () => {
  assert.ok(REMOTE_METHODS.includes('getDocByPath'), 'Host 要暴露这个端点')
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  assert.ok(/desc\('getDocByPath',\s*'request'\)/.test(src), '客户端要有对应 descriptor，wire 必须是 request')
  // 顺序必须与 Host 一致
  const block = src.slice(src.indexOf('descriptors: ['), src.indexOf('],', src.indexOf('descriptors: [')))
  const names = [...block.matchAll(/desc\('([A-Za-z0-9_]+)'/g)].map((m) => m[1])
  assert.deepEqual(names, REMOTE_METHODS, '两边方法名与顺序必须完全一致')
})

/* ------------------------------------------------------------------ */
/* A：地址 → 绝对路径                                                  */
/* ------------------------------------------------------------------ */

/**
 * 从 client.js 里取出 absolutePathOf 单独求值（纯函数）。
 * @returns {(address:string)=>string}
 */
function loadAbsolutePathOf() {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const start = src.indexOf('function absolutePathOf(')
  assert.ok(start > 0, 'client.js 里应该有 absolutePathOf')
  let depth = 0
  let i = src.indexOf('{', start)
  for (; i < src.length; i++) {
    if (src[i] === '{') depth++
    else if (src[i] === '}') { depth--; if (depth === 0) break }
  }
  const code = src.slice(start, i + 1) + '\n;absolutePathOf'
  const sandbox = { String, decodeURIComponent, RegExp }
  sandbox.globalThis = sandbox
  vm.createContext(sandbox)
  return vm.runInContext(code, sandbox, { filename: 'absolutePathOf.js' })
}

test('A1 absolute scope：地址里的路径就是绝对路径', () => {
  const f = loadAbsolutePathOf()
  assert.equal(
    f('dsh-resource://file/absolute/C:/work/%E6%8B%9B%E6%A0%87.docx'),
    '/C:/work/招标.docx',
    'Windows 盘符路径要解出来（中文要解码）',
  )
  assert.equal(f('dsh-resource://file/absolute/home/me/a.docx'), '/home/me/a.docx')
})

test('A2 session scope：绝对路径形态可用，工作区相对路径返回空串', () => {
  const f = loadAbsolutePathOf()
  // 绝对 Windows 路径
  assert.equal(
    f('dsh-resource://file/session/s1/C:/work/a.docx'),
    'C:/work/a.docx',
  )
  // 绝对 POSIX 路径在地址里表现为前导空段（/home/... → 空段 + home）
  assert.equal(f('dsh-resource://file/session/s1//home/me/a.docx'), '/home/me/a.docx')
  // 没有前导斜杠的 POSIX 路径就是工作区相对路径，按规则返回空串
  assert.equal(f('dsh-resource://file/session/s1/home/me/a.docx'), '')
  // 工作区相对路径：拿不到 workspace root，必须如实返回空串（不猜）
  assert.equal(f('dsh-resource://file/session/s1/docs/a.docx'), '', '相对路径不能瞎拼')
})

test('A3 非法地址一律返回空串，不抛', () => {
  const f = loadAbsolutePathOf()
  assert.equal(f(''), '')
  assert.equal(f('https://example.com/a.docx'), '')
  assert.equal(f('dsh-resource://file/'), '')
  assert.equal(f('dsh-resource://file/unknown/a.docx'), '')
  // 坏掉的百分号编码不能抛
  assert.doesNotThrow(() => f('dsh-resource://file/absolute/C:/%E0%A4%A.docx'))
})

/* ------------------------------------------------------------------ */
/* R：标签页注册结构                                                    */
/* ------------------------------------------------------------------ */

test('R1 标签页认得 *.docx，且用 extension 压过产品自带的 office 预览', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const block = src.slice(src.indexOf('sidebarRightTabs.register({'), src.indexOf('}, \'docanchor:sidebar tab type\')'))
  assert.ok(/patterns:\s*\['\*\.docx'\]/.test(block), '必须认领 *.docx')
  assert.ok(/priority:\s*'extension'/.test(block), '必须用 extension 档，否则会被自带 office 预览抢走')
  assert.ok(/kind:\s*TAB_KIND/.test(block) || /kind:/.test(block), '必须声明 kind')
  assert.ok(/id:\s*TAB_ID/.test(block), '必须声明 id（body/title 的 key）')
})

test('R2 侧栏注册失败不能连累整个插件', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  assert.ok(/侧栏标签页未注册/.test(src), '侧栏注册要 catch 并 warn，而不是 throw')
  assert.ok(/console\.warn/.test(src), '失败要打日志，不能静默吞掉')
})

test('R3 标签页有两种形态：已登记（带大纲）与未登记（只预览）', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const fn = src.slice(src.indexOf('function DocxTabPanel'), src.indexOf('function DocxTabTitle'))
  assert.ok(/showSide/.test(fn), '要按「有没有登记」决定是否显示大纲/标记区')
  assert.ok(/只做普通预览/.test(fn), '未登记时要明确说明只做预览')
  assert.ok(/DocxView/.test(fn), '两种形态都复用同一个 Word 渲染器')
  assert.ok(/jumpTo/.test(fn) && /data-anchor/.test(fn), '已登记时仍要能点击跳转')
})

test('R4 侧栏标签页是左右分栏：左固定栏放大纲/标记，右文档', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const fn = src.slice(src.indexOf('function DocxTabPanel'), src.indexOf('function DocxTabTitle'))

  // 结构：左栏 + 右栏，而不是上下堆叠
  assert.ok(/dav-tab-side/.test(fn), '必须有一个固定的左栏容器')
  assert.ok(/dav-tab-main/.test(fn), '必须有右栏容器（文档）')
  // 文档在右栏里，列表在左栏里 —— 顺序不能反（左先右后）
  const iSide = fn.indexOf('dav-tab-side')
  const iMain = fn.indexOf('dav-tab-main')
  assert.ok(iSide > 0 && iMain > 0 && iSide < iMain, '左栏必须在右栏之前渲染')

  // CSS：根是 row，左栏固定宽度、不可压缩
  assert.ok(/\.dav\.dav-tab\{[^}]*flex-direction:row/.test(src), '根容器必须是横向 flex')
  assert.ok(/\.dav-tab-side\{[^}]*flex:0 0 240px/.test(src), '左栏必须固定 240px 且不可收缩')
  // 不再有上下三态残留
  assert.ok(!/data-view="list"/.test(src), '不能残留「只看列表」的上下三态')
  assert.ok(!/\.dav-tab\[data-view=/.test(src), '不能残留 data-view 的分区规则')
})

test('R5 打开有标记的文档时默认停在「标记」页', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const fn = src.slice(src.indexOf('function DocxTabPanel'), src.indexOf('function DocxTabTitle'))
  assert.ok(/if \(\(d\?\.marks \?\? \[\]\)\.length\)/.test(fn), '要判「有没有标记」')
  assert.ok(/setTab\('marks'\)/.test(fn), '有标记就默认停在标记页')
})