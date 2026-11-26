/**
 * 把 docx-preview 打成自包含 IIFE，注入到 lib/client.js 的标记区。
 *
 * 为什么内联而不是 npm 依赖 / 外部脚本：
 *   DSH Client 通过 window.__ModuleLoader__.load({factory(require)}) 加载，
 *   require 只提供 react 与平台模块，不解析插件自己的 node_modules；
 *   浏览器页面也没有可引用的插件静态资源 URL。把库内联进 client.js 是
 *   唯一不依赖外部服务、刷新即可生效的方式。
 *
 * 用法：node scripts/build-docx-preview.mjs
 *   （client.js 里必须存在 BEGIN / END 两个标记注释）
 */
import { build } from 'esbuild'
import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const clientPath = join(root, 'lib', 'client.js')

const BEGIN = '/* DOCX_PREVIEW_IIFE_BEGIN */'
const END = '/* DOCX_PREVIEW_IIFE_END */'

const result = await build({
  entryPoints: [join(root, 'node_modules', 'docx-preview', 'dist', 'docx-preview.mjs')],
  bundle: true,
  write: false,
  format: 'iife',
  globalName: 'DOCX_PREVIEW',
  platform: 'browser',
  target: ['chrome100', 'firefox100', 'safari15'],
  minify: true,
  legalComments: 'none',
  logLevel: 'warning',
})
const iife = result.outputFiles[0].text

const src = readFileSync(clientPath, 'utf8')
const i = src.indexOf(BEGIN)
const j = src.indexOf(END)
if (i < 0 || j < 0 || j <= i) throw new Error('client.js 里找不到 DOCX_PREVIEW_IIFE 标记对')
const out = src.slice(0, i + BEGIN.length) + '\n' + iife + '\n' + src.slice(j)
writeFileSync(clientPath, out)
console.log(`docx-preview IIFE 注入完成：${(iife.length / 1024).toFixed(1)} KB`)
