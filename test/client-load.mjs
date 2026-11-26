/**
 * 客户端模块加载冒烟测试
 *
 * 只做一件真实发生的事：**把 lib/client.js 当普通脚本执行一遍**，
 * 用假的 `window.__ModuleLoader__` 收下注册，再用一个最小 React 影子调
 * factory —— 确认它不抛、返回 { inject, apply }、并把 descriptors 交出去。
 *
 * 为什么值得单测：client.js 是手写模块，语法错误只有在浏览器里才会炸成
 * 「slot entry crashed」，用户看到的是一片空白且没有任何提示。这里提前拦住。
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import vm from 'node:vm'

/** 最小 React 影子：够 factory 顶层跑完（组件本身不会被调用）。 */
function fakeReact() {
  const h = (type, props, ...children) => ({ type, props, children })
  const state = () => [undefined, () => {}]
  return {
    createElement: h,
    useState: state,
    useEffect: () => {},
    useMemo: (fn) => fn(),
    useCallback: (fn) => fn,
    useRef: (v) => ({ current: v ?? null }),
  }
}

/**
 * 执行 client.js，返回模块注册结果。
 * @returns {{id:string, factory:Function, exports:any}}
 */
function loadClientModule() {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  let captured = null
  const sandbox = {
    window: {
      __ModuleLoader__: {
        load(entry) { captured = entry },
      },
    },
    document: {
      createElement: () => ({ setAttribute() {}, remove() {} }),
      head: { appendChild() {} },
      querySelectorAll: () => [],
    },
    console,
    atob: (s) => Buffer.from(s, 'base64').toString('binary'),
    btoa: (s) => Buffer.from(s, 'binary').toString('base64'),
    setTimeout,
    clearTimeout,
    Promise,
    Map,
    Set,
    Array,
    Object,
    String,
    Number,
    JSON,
    Math,
    Infinity,
    Uint8Array,
    Error,
    TypeError,
  }
  sandbox.globalThis = sandbox
  vm.createContext(sandbox)
  vm.runInContext(src, sandbox, { filename: 'client.js' })
  if (!captured) throw new Error('client.js 没有调用 window.__ModuleLoader__.load')
  return { id: captured.id, factory: captured.factory }
}

test('C5 client.js 能被当作脚本执行并注册模块（id = 包名）', () => {
  const mod = loadClientModule()
  assert.equal(mod.id, 'dsh-doc-anchor-viewer')
  assert.equal(typeof mod.factory, 'function')
})

test('C6 factory 返回 { inject, apply }，且 inject 只含 remote', () => {
  const mod = loadClientModule()
  const exp = mod.factory((name) => {
    if (name === 'react') return fakeReact()
    throw new Error('客户端不应该 require 其它模块：' + name)
  })
  assert.equal(typeof exp, 'object')
  // vm 里造出来的数组来自另一个 realm，deepEqual 会因原型不同而失败，先摊平成本地数组
  assert.deepEqual([...exp.inject], ['remote'], 'remote.<ns> 是 $mount 后才出现的服务，写进 inject 会永远 pending')
  assert.equal(typeof exp.apply, 'function')
})

test('C7 docx-preview 全局在内联 bundle 里可用（renderAsync 必须存在）', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  // 内联 IIFE 以 var DOCX_PREVIEW = ... 形式定义，必须在同一脚本作用域可见
  assert.ok(/var DOCX_PREVIEW\s*=\s*\(/.test(src), 'BUNDLE 必须以 var 全局形式出现（factory 里直接引用）')
  assert.ok(/renderAsync:/.test(src) || /renderAsync\s*(=|:)/.test(src), 'bundle 必须导出 renderAsync')
})

test('C9 面板按「自由文本 + 哈希配色」渲染，不再依赖枚举', async () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')

  // 面板不能再有 SEVERITIES 白名单数组
  assert.ok(!/const SEVERITIES = \[/.test(src), '面板不应再有 SEVERITIES 白名单')
  // 必须有哈希配色与显式颜色支持
  assert.ok(/function colorOf\(/.test(src), '必须有 colorOf 哈希配色')
  assert.ok(/function markColor\(/.test(src), '必须有 markColor（显式 color 优先）')
  // 必须能原样显示自定义标签
  assert.ok(/SEVERITY_LABEL\[s\] \?\? s/.test(src), '内置值给中文、自定义值原样显示')
  assert.ok(/CATEGORY_LABEL\[s\] \?\? s/.test(src), '分类同理')

  // Host 侧：推荐值仍导出，且额外导出 SUGGESTED_*
  const mod = await import('../lib/index.js')
  assert.ok(mod.SUGGESTED_SEVERITIES && mod.SUGGESTED_CATEGORIES, 'Host 要导出推荐值')
})

test('C8 客户端不发无参方法的参数（wire 与 Host 形参数量必须一一对应）', () => {
  const src = readFileSync(new URL('../lib/client.js', import.meta.url), 'utf8')
  // 无参方法调用点不能传 request（传了会 gateway/arguments-invalid）
  const noArgCalls = [...src.matchAll(/callRemote\(remoteRef\.current,\s*'([A-Za-z0-9_]+)'\s*\)/g)].map((m) => m[1])
  for (const name of new Set(noArgCalls)) {
    assert.ok(['diagnostics', 'listDocs'].includes(name), `${name} 被当无参调用，但只允许 diagnostics / listDocs 无参`)
  }
  // 带参方法必须传 { request: ... }
  const withArg = [...src.matchAll(/callRemote\(remoteRef\.current,\s*'([A-Za-z0-9_]+)',\s*\{\s*request/g)].map((m) => m[1])
  assert.ok(withArg.includes('getDocxBytes'), 'getDocxBytes 必须带 request')
  assert.ok(withArg.includes('openDoc'), 'openDoc 必须带 request')
})
