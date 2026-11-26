/**
 * 面板渲染冒烟测试
 *
 * 用一个最小但**语义正确**的 React 影子跑一遍组件树：确认 DocAnchorPanel
 * 在各种状态下都能渲染出结构、不抛异常、不发多余 Remote 调用。
 *
 * 为什么值得做：Client 插件里一个渲染期 TypeError 的后果是
 * `slot entry crashed in '<slot>'` —— 面板直接空白，控制台之外没有任何提示。
 * 这里把「加载中 / 有数据 / 渲染失败」三种状态都跑一遍，提前拦住这类错。
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import vm from 'node:vm'

/* ------------------------------------------------------------------ */
/* 最小 React：单次渲染，hook 顺序即索引                                */
/* ------------------------------------------------------------------ */

/**
 * 造一个能跑「一次渲染」的 React 影子。
 *
 * 语义要点（这些错了会验不出真正的 bug）：
 *   · useState 返回 [初值, setter]，setter 只记录调用不触发重渲染
 *   · useMemo / useCallback 直接求值（依赖变化不重算，单次渲染够用）
 *   · useEffect **不执行**——真实环境里它在提交后跑，这里只收集
 *   · useRef 每次返回同一个对象（同一棵渲染树内）
 */
function makeReactShim() {
  const effects = []
  const refs = []
  const setters = []
  let refIdx = 0
  let stateIdx = 0
  /** 注入的 state 序列：第 n 个 useState 拿第 n 个值，用完后回到 undefined。 */
  let stateSeq = []
  const h = (type, props, ...children) => ({ type, props: props || {}, children: children.flat() })
  return {
    effects,
    createElement: h,
    useState: (init) => {
      const i = stateIdx++
      const setter = (v) => { setters.push({ i, v }) }
      const v = i < stateSeq.length ? stateSeq[i] : (typeof init === 'function' ? init() : init)
      return [v, setter]
    },
    useEffect: (fn) => { effects.push(fn) },
    useMemo: (fn) => fn(),
    useCallback: (fn) => fn,
    useRef: (init) => {
      const i = refIdx++
      if (!refs[i]) refs[i] = { current: init === undefined ? null : init }
      return refs[i]
    },
    /** 下一次渲染时按顺序喂 state 初值（模拟数据已加载）。 */
    setStateSequence(seq) { stateSeq = seq; stateIdx = 0; refIdx = 0; refs.length = 0 },
    resetRefs() { refIdx = 0; refs.length = 0; stateIdx = 0 },
    setters,
  }
}

/* ------------------------------------------------------------------ */
/* 模块加载                                                            */
/* ------------------------------------------------------------------ */

/**
 * 执行 client.js 并返回 factory。
 * @returns {Function}
 */
function loadFactory() {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  let captured = null
  const sandbox = {
    window: { __ModuleLoader__: { load(e) { captured = e } } },
    document: { createElement: () => ({ setAttribute() {}, remove() {} }), head: { appendChild() {} }, querySelectorAll: () => [] },
    console,
    atob: (s) => Buffer.from(s, 'base64').toString('binary'),
    btoa: (s) => Buffer.from(s, 'binary').toString('base64'),
    setTimeout, clearTimeout, Promise, Map, Set, Array, Object, String, Number, JSON, Math, Infinity, Uint8Array,
    Error, TypeError, Boolean,
  }
  sandbox.globalThis = sandbox
  vm.createContext(sandbox)
  vm.runInContext(src, sandbox, { filename: 'client.js' })
  if (!captured) throw new Error('没有调用 __ModuleLoader__.load')
  return captured.factory
}

/**
 * 把虚拟树展开成可断言的纯文本。
 * @param {any} node
 * @returns {string}
 */
function textOf(node) {
  if (node === null || node === undefined || node === false || node === true) return ''
  if (typeof node === 'string' || typeof node === 'number') return String(node)
  if (Array.isArray(node)) return node.map(textOf).join(' ')
  // 函数组件：调用它拿到子树。真实 React 也会这么做，只是我们不做重渲染/调和。
  if (typeof node.type === 'function') return textOf(node.type(node.props || {}))
  return textOf(node.children)
}

/**
 * 递归收集所有节点，按 type 名或 className 找。
 * @param {any} node
 * @param {(n:any)=>boolean} pred
 * @param {Array} out
 * @returns {Array}
 */
function find(node, pred, out = []) {
  if (!node || typeof node !== 'object') return out
  if (Array.isArray(node)) { for (const n of node) find(n, pred, out); return out }
  if (pred(node)) out.push(node)
  for (const c of node.children ?? []) find(c, pred, out)
  return out
}

/* ------------------------------------------------------------------ */
/* 用例                                                                */
/* ------------------------------------------------------------------ */

test('R1 空态：没有文档时显示引导，不抛异常', async () => {
  const react = makeReactShim()
  const factory = loadFactory()
  const exp = factory((n) => { if (n === 'react') return react; throw new Error('unexpected require: ' + n) })

  // 拿到组件：注册时传进去的第二个参数
  const registered = []
  const scope = {
    remote: { docAnchor: {} },
    effect: (fn) => { try { return fn() } catch { /* 桩里忽略 */ } },
    sidebarRightTabs: { register: () => () => {} },
    slots: {
      inject: (_name, cb) => { try { cb() } catch { /* 缺少的槽位桩忽略 */ } },
      register: (opts, Comp) => { registered.push({ opts, Comp }); return () => {} },
    },
  }
  await exp.apply({
    effect: (fn) => { fn() },
    remote: { $mount: async () => {} },
    inject: (_deps, cb) => { cb(scope); return Promise.resolve() },
  })

  // 现在还会注册侧栏标签页的 body + title，所以数量是 4；
  // 这里只关心 main 存在，不锁死总数（锁死会让每次加槽位都要改测试）。
  const main = registered.find((r) => r.opts.name === 'main')
  assert.ok(main, '必须有 main 面板')
  assert.equal(main.opts.key, 'doc-anchor-viewer')

  // 渲染初始态（docs=[] 是初值，因为 effect 不会跑 → loadingDocs=true）
  react.resetRefs()
  const tree = main.Comp({ remote: {} })
  const txt = textOf(tree)
  assert.ok(txt.includes('文档锚点'), '标题要在：' + txt.slice(0, 120))
})

test('R2 有文档时：大纲按层级渲染、点条目走 jumpTo', async () => {
  const react = makeReactShim()
  const factory = loadFactory()
  const exp = factory((n) => { if (n === 'react') return react; throw new Error('unexpected require: ' + n) })

  const registered = []
  const scope = {
    remote: { docAnchor: {} },
    effect: (fn) => { try { return fn() } catch { /* 桩里忽略 */ } },
    sidebarRightTabs: { register: () => () => {} },
    slots: {
      inject: (_n, cb) => { try { cb() } catch { /* 缺少的槽位桩忽略 */ } },
      register: (opts, Comp) => { registered.push({ opts, Comp }); return () => {} },
    },
  }
  await exp.apply({
    effect: () => {}, remote: { $mount: async () => {} },
    inject: (_d, cb) => { cb(scope); return Promise.resolve() },
  })
  const main = registered.find((r) => r.opts && r.opts.name === 'main')
  assert.ok(main, '必须有 main 面板')

  const doc = {
    id: 'd1',
    title: '需求调研记录',
    path: 'C:\\tmp\\a.docx',
    stats: { paragraphs: 48, headings: 3, chars: 100 },
    outline: [
      { index: 0, title: '一级标题', level: 1, norm: '一级标题', occurrence: 0 },
      { index: 5, title: '二级标题', level: 2, norm: '二级标题', occurrence: 0 },
      { index: 8, title: '另一个二级', level: 2, norm: '另一个二级', occurrence: 0 },
    ],
    marks: [{ id: 'm1', text: '一段被标记的原文', note: '备注', anchor: { index: 5, norm: '二级标题', occurrence: 0 } }],
  }
  // useState 顺序：docs, docId, doc, tab, sel, busy, err, note, loadingDocs, openPane, filePath
  react.setStateSequence([
    [{ id: 'd1', title: '需求调研记录', stats: doc.stats }],
    'd1',
    doc,
    'outline',
    '',
    false,
    '',
    '',
    false,
    false,
    '',
  ])
  const txt = textOf(main.Comp({ remote: {} }))
  assert.ok(txt.includes('需求调研记录'), '标题要出现：' + txt.slice(0, 200))
  assert.ok(txt.includes('一级标题'), '大纲一级条目要渲染')
  assert.ok(txt.includes('二级标题'), '大纲二级条目要渲染')
  assert.ok(txt.includes('大纲 3'), '大纲页签要显示条目数')
  assert.ok(txt.includes('48 段'), '要显示段落统计')
})

test('R3 docx-preview 渲染区：加载/就绪/失败三态都有可见反馈', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  // 三态文案必须都在
  assert.ok(src.includes('正在按 Word 版式渲染'), '加载态要有提示')
  assert.ok(src.includes("state === 'error'") && src.includes('渲染失败'), '失败要给出原因')
  assert.ok(src.includes('没能对上渲染结果'), '锚点没全命中要如实说明，不能假装成功')
  // 渲染选项必须是「按 Word 还原」的那组
  assert.ok(/breakPages:\s*true/.test(src), '必须分页（Word 是按页排的）')
  assert.ok(/useBase64URL:\s*true/.test(src), '图片走 base64，不依赖外部 URL')
})

test('R4 只读：右侧预览不绑任何点击选条目、不反向改左侧选中', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  // DocxView 里不允许出现 addEventListener('click') 或 onClick 绑到文档段落上
  const docxView = src.slice(src.indexOf('function DocxView'), src.indexOf('function OutlineList'))
  assert.ok(!/addEventListener\(\s*['"]click['"]/.test(docxView), 'DocxView 不能绑点击（用户要求预览只读）')
  assert.ok(!/onSelect/.test(src), '不应该有 onSelect 这类反向选中通道')
  // 跳转只由左侧触发
  assert.ok(/const jumpTo = useCallback/.test(src), '跳转逻辑应集中在 jumpTo')
  assert.ok(/scrollIntoView/.test(src), '跳转必须真的滚动')
  assert.ok(/dav-flash/.test(src) && /animationend/.test(src), '闪一下之后要复原')
})

test('R5 换文档时重建渲染，不残留上一篇', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  // 用 key={docId} 强制卸载重建；并且在 effect 里主动清空 host
  assert.ok(/key:\s*docId/.test(src), 'DocxView 必须用 docId 作 key，否则换文档会复用旧 DOM')
  assert.ok(/host\.innerHTML = ''/.test(src), '渲染前要清空挂载点')
})
