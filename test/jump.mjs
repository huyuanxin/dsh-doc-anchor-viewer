/**
 * 跳转链路测试（实际踩过的时序 bug）
 *
 * 守的是一条**只在真实运行时暴露**的问题：
 *   DocxView 的渲染 effect 只依赖 [docId]，而 anchors 来自父组件异步加载的
 *   doc。两者到达顺序不稳定，渲染跑在 anchors 之前时 → 一个锚点都打不上 →
 *   文档显示正常，但点左侧什么都不跳。
 *
 * 这类 bug 用「读代码」很难看出来（代码本身没错，错在时序），所以用测试钉死：
 *   T1 渲染完成时锚点还没到 → 之后补上（不能永远空着）
 *   T2 锚点已就绪再渲染 → 一次到位
 *   T3 已经打过锚点 → 不重复打（幂等）
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import vm from 'node:vm'

/* ------------------------------------------------------------------ */
/* 最小 DOM：够 stampAnchors 跑                                          */
/* ------------------------------------------------------------------ */

/** 一个只支持 querySelectorAll('p') / querySelector('[data-anchor="N"]') 的假 host。 */
function makeHost(paragraphTexts) {
  const ps = paragraphTexts.map((t) => ({
    textContent: t,
    attrs: {},
    setAttribute(k, v) { this.attrs[k] = v },
    getAttribute(k) { return this.attrs[k] ?? null },
  }))
  return {
    ps,
    querySelectorAll(sel) { return sel === 'p' ? ps : [] },
    querySelector(sel) {
      const m = /\[data-anchor="(\d+)"\]/.exec(sel)
      if (!m) return null
      const want = m[1]
      return ps.find((p) => p.attrs['data-anchor'] === want) ?? null
    },
    get innerHTML() { return '' },
    set innerHTML(_v) { for (const p of ps) p.attrs = {} },
  }
}

/* ------------------------------------------------------------------ */
/* 从 client.js 里取纯函数                                              */
/* ------------------------------------------------------------------ */

/**
 * 把 client.js 在沙箱里跑起来，拿到内部函数（stampAnchors / normAnchor）。
 *
 * 为什么走这条路而不是 import：client.js 是浏览器模块，顶层调
 * window.__ModuleLoader__.load，没有 ES 导出。沙箱里跑一遍最接近真实。
 */
function loadInternals() {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  // 把这两个函数抽出来单独求值：它们是纯函数，不依赖 React
  const grab = (name) => {
    const start = src.indexOf(`function ${name}(`)
    assert.ok(start > 0, `client.js 里应该有 ${name}`)
    // 用花括号配平找到函数结尾
    let depth = 0
    let i = src.indexOf('{', start)
    const begin = i
    for (; i < src.length; i++) {
      if (src[i] === '{') depth++
      else if (src[i] === '}') { depth--; if (depth === 0) break }
    }
    return src.slice(start, i + 1)
  }
  const code = grab('normAnchor') + '\n' + grab('stampAnchors') + '\n;({ normAnchor, stampAnchors })'
  const sandbox = { Map, Set, String, Number, Array, Object, console }
  sandbox.globalThis = sandbox
  vm.createContext(sandbox)
  return vm.runInContext(code, sandbox, { filename: 'client-internals.js' })
}

/* ------------------------------------------------------------------ */
/* 用例                                                                */
/* ------------------------------------------------------------------ */

test('T1 锚点晚于渲染到达：必须先空着，随后补上（这正是「点了不跳」）', () => {
  const { stampAnchors } = loadInternals()
  const host = makeHost(['第一章 总则', '正文一段', '第二章 要求'])

  const anchors = [
    { index: 0, norm: '第一章总则', occurrence: 0 },
    { index: 2, norm: '第二章要求', occurrence: 0 },
  ]

  // ① 渲染完成时锚点还没到（真实时序）：一个都打不上
  const before = stampAnchors(host, [])
  assert.equal(before.matched, 0)
  assert.equal(host.querySelector('[data-anchor="0"]'), null, '此时确实没有锚点 → 点不动')

  // ② 锚点到了，补打（对应新加的第二个 effect）
  const after = stampAnchors(host, anchors)
  assert.equal(after.matched, 2, '补打之后必须能对上')
  assert.equal(after.missing.length, 0)
  assert.ok(host.querySelector('[data-anchor="0"]'), '第 0 段可跳')
  assert.ok(host.querySelector('[data-anchor="2"]'), '第 2 段可跳')
})

test('T2 锚点先到：一次到位，不需要补', () => {
  const { stampAnchors } = loadInternals()
  const host = makeHost(['甲', '乙', '丙'])
  const anchors = [
    { index: 0, norm: '甲', occurrence: 0 },
    { index: 1, norm: '乙', occurrence: 0 },
    { index: 2, norm: '丙', occurrence: 0 },
  ]
  const stat = stampAnchors(host, anchors)
  assert.equal(stat.matched, 3)
  assert.equal(stat.missing.length, 0)
})

test('T3 重复打锚点是幂等的（补打不会把已有锚点打乱）', () => {
  const { stampAnchors } = loadInternals()
  const host = makeHost(['甲', '甲', '乙'])
  const anchors = [
    { index: 0, norm: '甲', occurrence: 0 },
    { index: 1, norm: '甲', occurrence: 1 },
  ]
  const a = stampAnchors(host, anchors)
  assert.equal(a.matched, 2)
  const second = stampAnchors(host, anchors)
  // 第二遍仍然对同一批元素给出同样的结果（不会因为已打过就错位）
  assert.equal(second.matched, 2)
  assert.equal(host.querySelector('[data-anchor="0"]').textContent, '甲')
  assert.equal(host.querySelector('[data-anchor="1"]').textContent, '甲')
})

test('T4 源码守卫：渲染 effect 只依赖 docId，锚点走 ref；另有一个补打 effect', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const docxView = src.slice(src.indexOf('function DocxView'), src.indexOf('function OutlineList'))

  // 渲染 effect 不能把 anchors 放进依赖（那会导致锚点一到就整篇重渲染）
  assert.ok(/\}, \[docId\]\)/.test(docxView), '渲染 effect 必须只依赖 docId')
  assert.ok(!/\}, \[docId,\s*anchors\]\)/.test(docxView), 'anchors 不能进渲染 effect 的依赖')

  // 但必须通过 ref 读到最新值
  assert.ok(/anchorsRef/.test(docxView), '必须用 anchorsRef 让 effect 读到最新锚点')
  assert.ok(/anchorsRef\.current/.test(docxView), 'stampAnchors 要读 anchorsRef.current')

  // 且必须有一个「锚点后到就补打」的 effect
  assert.ok(/\}, \[anchors, state\]\)/.test(docxView), '必须有一个依赖 [anchors, state] 的补打 effect')
  // 补打必须是幂等的（已经有锚点就不再打）
  assert.ok(/querySelector\('p\[data-anchor\]'\)/.test(docxView), '补打前要检查是否已经有锚点')
})

test('T5 跳转失败要有可见反馈，不能静默无反应', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  const jump = src.slice(src.indexOf('const jumpTo = useCallback'), src.indexOf('/** 打开文档后统一收尾'))
  // 找不到元素时必须给提示（否则用户看到的就是「点了没反应」）
  assert.ok(/setNote\(/.test(jump), '定位失败要 setNote 提示用户')
  assert.ok(/没能定位到/.test(jump), '提示文案要说明是「定位不到」而不是静默')
})
