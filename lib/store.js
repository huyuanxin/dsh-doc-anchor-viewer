/**
 * 文档锚点视图 · 持久化
 *
 * 把「打开过哪些文档、每篇的大纲与标记」落在磁盘上，而不是内存——
 * 用户在别的会话里让 Agent 开的文档，回到面板还得在。
 *
 * 三个硬要求（都是 DSH 插件里反复踩到的坑）：
 *  1. **多候选数据目录**：Desktop 一次启动可能组合多个 profile，硬编码一个路径
 *     会在升级后表现为「我开的文档不见了」。
 *  2. **原子写**：先写临时文件再 rename，崩溃不留半截 JSON。
 *  3. **串行化写队列**：两个并发读-改-写会互相覆盖（Agent 连开两篇文档就会出现）。
 *
 * 目录结构：
 *   <root>/doc-anchor-viewer/
 *     docs.json                文档索引（面板列表页只读这一个文件）
 *     docs/<docId>.json        单篇文档：元信息 + 段落文本 + 大纲 + 标记
 */

import { existsSync, statSync } from 'node:fs'
import { mkdir, readFile, rename, writeFile, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'

/** 环境变量覆盖（测试与探针用，避免污染用户数据）。 */
const ENV_OVERRIDE = 'DOC_ANCHOR_VIEWER_DIR'

/**
 * 两个路径是否指向同一目录（Windows 下忽略大小写与末尾分隔符）。
 * @param {string} a
 * @param {string} b
 * @returns {boolean}
 */
function sameDir(a, b) {
  const norm = (p) => {
    const r = resolve(p).replace(/[\\/]+$/, '')
    return process.platform === 'win32' ? r.toLowerCase() : r
  }
  return norm(a) === norm(b)
}

/**
 * 数据目录候选，最具体优先。
 *
 * ⚠️ `DSH_PROFILE_DIR` 只在它**确实是本进程自己的 profile 目录**时才认：
 * 实测它在某些启动方式下仍指向上一个 profile，认了就会跨 profile 读写别人的数据。
 *
 * @returns {string[]} 绝对路径列表
 */
function candidates() {
  const override = process.env[ENV_OVERRIDE]
  if (override) return [override]

  const roots = []
  const home = process.env.DSH_HOME
  const profile = process.env.DSH_PROFILE
  const profileDir = home && profile ? join(home, 'profiles', profile) : null
  if (profileDir) roots.push(profileDir)
  if (home) roots.push(home)
  const dir = process.env.DSH_PROFILE_DIR
  if (dir && !roots.includes(dir) && (profileDir === null || sameDir(dir, profileDir))) {
    roots.push(dir)
  }
  const cwd = process.cwd()
  if (!roots.includes(cwd)) roots.push(cwd)
  return roots.map((r) => join(r, 'doc-anchor-viewer'))
}

/**
 * 解析数据根目录：显式配置优先；否则「已有数据的位置优先」，都没有则用最具体的候选。
 * @param {string} [configured]
 * @returns {string}
 */
export function resolveDataDir(configured) {
  const explicit = String(configured ?? '').trim()
  if (explicit) return resolve(explicit)
  const list = candidates()
  for (const c of list) {
    if (existsSync(join(c, 'docs.json'))) return c
  }
  return list[0]
}

/** 文件名安全化：docId 只允许我们自己生成的字符。 */
const DOC_ID_RE = /^[A-Za-z0-9_-]{1,64}$/

/**
 * 文档库。
 */
export class DocStore {
  /**
   * @param {string} dataDir
   */
  constructor(dataDir) {
    this.root = dataDir
    /** 串行化写队列：所有读-改-写都排在这条链上。 */
    this._chain = Promise.resolve()
  }

  /** 索引文件路径。 */
  get indexPath() { return join(this.root, 'docs.json') }

  /**
   * 单篇文档文件路径。
   * @param {string} docId
   * @returns {string}
   */
  docPath(docId) {
    if (!DOC_ID_RE.test(String(docId ?? ''))) throw new Error('DOC_ID_INVALID')
    return join(this.root, 'docs', docId + '.json')
  }

  /**
   * 原子写：先写临时文件，再 rename 覆盖。
   * @param {string} file
   * @param {any} value
   */
  async _writeJson(file, value) {
    await mkdir(dirname(file), { recursive: true })
    const tmp = file + '.' + process.pid + '.' + Date.now() + '.tmp'
    await writeFile(tmp, JSON.stringify(value), 'utf8')
    await rename(tmp, file)
  }

  /**
   * 在写队列上跑一段读-改-写。
   * @template T
   * @param {() => Promise<T>} fn
   * @returns {Promise<T>}
   */
  _serial(fn) {
    const next = this._chain.then(fn, fn)
    // 链条不能因为单次失败而断掉，否则后续所有写都会 reject
    this._chain = next.catch(() => {})
    return next
  }

  /**
   * 读索引；文件不存在视为空库。
   * @returns {Promise<{docs: Array<object>}>}
   */
  async readIndex() {
    try {
      const raw = await readFile(this.indexPath, 'utf8')
      const v = JSON.parse(raw)
      return { docs: Array.isArray(v?.docs) ? v.docs : [] }
    } catch (e) {
      if (e && (e.code === 'ENOENT' || e.code === 'ENOTDIR')) return { docs: [] }
      // JSON 损坏：不要静默丢数据，抛出去让上层给出可操作的错误
      if (e instanceof SyntaxError) throw new Error('索引文件损坏（docs.json）：' + e.message)
      throw e
    }
  }

  /**
   * 列出全部文档（按最近打开倒序）。
   * @returns {Promise<Array<object>>}
   */
  async listDocs() {
    const { docs } = await this.readIndex()
    return [...docs].sort((a, b) => String(b.openedAt || '').localeCompare(String(a.openedAt || '')))
  }

  /**
   * 读单篇文档。
   * @param {string} docId
   * @returns {Promise<object|null>}
   */
  async readDoc(docId) {
    try {
      const raw = await readFile(this.docPath(docId), 'utf8')
      return JSON.parse(raw)
    } catch (e) {
      if (e && (e.code === 'ENOENT' || e.code === 'ENOTDIR')) return null
      throw e
    }
  }

  /**
   * 写单篇文档（同时更新索引里的摘要字段）。
   * @param {object} doc
   * @returns {Promise<void>}
   */
  async saveDoc(doc) {
    await this._serial(async () => {
      await this._writeJson(this.docPath(doc.id), doc)
      const { docs } = await this.readIndex()
      const meta = {
        id: doc.id,
        path: doc.path,
        title: doc.title,
        openedAt: doc.openedAt,
        size: doc.size,
        mtimeMs: doc.mtimeMs,
        stats: doc.stats,
        markCount: (doc.marks ?? []).length,
      }
      const idx = docs.findIndex((d) => d.id === doc.id)
      if (idx >= 0) docs[idx] = meta
      else docs.push(meta)
      await this._writeJson(this.indexPath, { schemaVersion: 1, docs })
    })
  }

  /**
   * 只更新索引里的摘要（标记增删后同步 markCount，不必重写整篇文档）。
   * @param {string} docId
   * @param {object} patch
   * @returns {Promise<void>}
   */
  async patchIndex(docId, patch) {
    await this._serial(async () => {
      const { docs } = await this.readIndex()
      const idx = docs.findIndex((d) => d.id === docId)
      if (idx < 0) return
      docs[idx] = { ...docs[idx], ...patch }
      await this._writeJson(this.indexPath, { schemaVersion: 1, docs })
    })
  }

  /**
   * 删除一篇文档（连同它的段落文件与索引项）。
   * @param {string} docId
   * @returns {Promise<boolean>} 是否真的删掉了
   */
  async removeDoc(docId) {
    return this._serial(async () => {
      const { docs } = await this.readIndex()
      const next = docs.filter((d) => d.id !== docId)
      const found = next.length !== docs.length
      if (!found) return false
      await this._writeJson(this.indexPath, { schemaVersion: 1, docs: next })
      try { await rm(this.docPath(docId)) } catch { /* 文件已不在，视为已删除 */ }
      return true
    })
  }

  /**
   * 读-改-写一篇文档的标记。
   *
   * @param {string} docId
   * @param {(doc: object) => {doc?: object, result?: any}|Promise<{doc?: object, result?: any}>} fn
   * @returns {Promise<{ok:boolean, error?:string, result?:any}>}
   */
  async mutateDoc(docId, fn) {
    return this._serial(async () => {
      const doc = await this.readDoc(docId)
      if (!doc) return { ok: false, error: '文档不存在或已被删除' }
      const r = await fn(doc)
      if (r && r.doc) {
        await this._writeJson(this.docPath(docId), r.doc)
        const { docs } = await this.readIndex()
        const idx = docs.findIndex((d) => d.id === docId)
        if (idx >= 0) {
          docs[idx] = { ...docs[idx], markCount: (r.doc.marks ?? []).length, title: r.doc.title }
          await this._writeJson(this.indexPath, { schemaVersion: 1, docs })
        }
      }
      return { ok: true, result: r ? r.result : undefined }
    })
  }

  /**
   * 诊断用：数据目录与文件状态。
   * @returns {Promise<object>}
   */
  async diagnostics() {
    const out = { root: this.root, exists: existsSync(this.root), docs: 0, bytes: 0 }
    try {
      const { docs } = await this.readIndex()
      out.docs = docs.length
      for (const d of docs) {
        try { out.bytes += statSync(this.docPath(d.id)).size } catch { /* 单篇缺失不致命 */ }
      }
    } catch (e) {
      out.error = String(e?.message ?? e)
    }
    return out
  }
}
