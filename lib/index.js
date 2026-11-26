/**
 * 文档锚点视图 · Host 服务
 *
 * 职责：
 *   1. 登记一篇 docx（路径或 base64），算出「段落序列 + 大纲」，落盘；
 *   2. 给浏览器返回渲染所需的 docx 字节（docx-preview 在客户端渲染）；
 *   3. 维护用户标记（可带锚点），供面板列表与 Agent 读写；
 *   4. 面对 Agent 的工具接口（open/list/outline/marks），让别的会话也能用。
 *
 * 边界：**只读用户的文档**，不改它；所有写操作只写本插件自己的数据目录。
 */

import fs from 'node:fs'
import { createHash } from 'node:crypto'
import path from 'node:path'

import { readZipBuffer } from './ooxml.js'
import { parseParagraphs, parseStyles, buildOutline, occurrenceMap } from './docparse.js'
import { DocStore, resolveDataDir } from './store.js'

/**
 * 解析一份 docx 时**真正需要**的 ZIP 条目。
 *
 * ⚠️ 刻意不列图片、嵌入对象、theme、字体……它们对「段落序列 + 大纲」毫无贡献，
 *    却可能是文件体积的绝大部分。列进来就等于让宿主把它们全部解压进内存。
 *
 * @type {string[]}
 */
const DOC_PARTS = ['word/document.xml', 'word/styles.xml']

/* ------------------------------------------------------------------ */
/* 配置                                                               */
/* ------------------------------------------------------------------ */

/** 配置校验（Cordis 的 `~standard` 校验协议）。 */
export const Config = {
  '~standard': {
    version: 1,
    vendor: 'dsh-doc-anchor-viewer',
    validate(value) {
      const v = (value && typeof value === 'object') ? { ...value } : {}
      const num = (k, def, min, max) => {
        const n = Number(v[k] ?? def)
        if (!Number.isFinite(n) || n < min || n > max) return def
        return n
      }
      return {
        value: {
          dataDir: typeof v.dataDir === 'string' ? v.dataDir.trim() : '',
          maxDocxMb: num('maxDocxMb', 30, 1, 200),
        },
      }
    },
  },
}

/* ------------------------------------------------------------------ */
/* 失败返回                                                            */
/* ------------------------------------------------------------------ */

/**
 * 统一的失败形状（Remote 与工具都吃这一种）。
 * @param {string} code
 * @param {string} message
 * @returns {{ok:false, error:{code:string, message:string}}}
 */
function fail(code, message) {
  return { ok: false, error: { code, message } }
}

/**
 * Remote 参数在这条链路上会被包成 `{ request: {...} }`；两种形状都接受。
 * @param {any} request
 * @returns {any}
 */
function unwrapArg(request) {
  if (request && typeof request === 'object' && request.request && typeof request.request === 'object') {
    return request.request
  }
  return request && typeof request === 'object' ? request : {}
}

/** 剥掉用户从资源管理器复制的引号。 */
function stripQuotes(v) {
  let s = String(v ?? '').trim()
  if (s.length >= 2) {
    const a = s[0]
    const b = s[s.length - 1]
    if ((a === '"' && b === '"') || (a === "'" && b === "'")) s = s.slice(1, -1).trim()
  }
  return s
}

/**
 * docId 的形状校验。
 *
 * ⚠️ 必须在**进入 store 之前**检查：docId 会被拼进文件路径，
 * `../../x` 这种值可以读到数据目录外的文件。抛异常也不行——Remote 会把
 * 未捕获的异常翻成 gateway/internal，用户看到的是一句莫名的话；
 * 这里统一返回可读的业务失败。
 *
 * @param {any} docId
 * @returns {string|null} 合法时返回原值，否则返回 null
 */
function safeDocId(docId) {
  const s = String(docId ?? '')
  return /^[A-Za-z0-9_-]{1,64}$/.test(s) ? s : null
}

/**
 * docId 由绝对路径 + 修改时间决定：同一份文件重复打开落到同一条记录，
 * 改过之后（mtime 变）自然成为新记录，旧标记不会错误地挂到新内容上。
 * @param {string} filePath
 * @param {number} mtimeMs
 * @returns {string}
 */
function makeDocId(filePath, mtimeMs) {
  const h = createHash('sha1').update(String(filePath).toLowerCase() + '\u0000' + mtimeMs).digest('hex')
  return 'd' + h.slice(0, 16)
}

/* ------------------------------------------------------------------ */
/* 评判字段（结构化，供面板分类着色）                                    */
/* ------------------------------------------------------------------ */

/**
 * 推荐的危险程度取值。
 *
 * ⚠️ 这**不是白名单**，只是给调用方（尤其是模型）的默认词汇建议。
 * 任何字符串都会被接受并按值自动配色——这个插件不想只服务于招投标，
 * 医疗病历、合同审阅、代码评审都能用它打标记，各自有各自的严重程度词汇。
 *
 * 内置值的好处只有一个：面板上有现成的中文标签与语义色；
 * 自定义值则原样显示、颜色由字符串哈希稳定生成（同一个词永远是同一个颜色）。
 */
export const SUGGESTED_SEVERITIES = ['fatal', 'high', 'medium', 'low', 'info']

/** 推荐的分类取值。同样**不是白名单**，任何字符串都接受。 */
export const SUGGESTED_CATEGORIES = [
  'invalid', 'qualification', 'scoring', 'technical', 'commercial', 'contract', 'other',
]

/**
 * 向后兼容的别名。
 *
 * 早期版本把这两个当强制枚举导出，外部可能已经引用了这两个名字；
 * 语义已放宽，但名字保留，避免无谓的破坏性改动。
 */
export const SEVERITIES = SUGGESTED_SEVERITIES
export const CATEGORIES = SUGGESTED_CATEGORIES

/** 标签/分类这类自由文本的上限：够表达语义，又不至于把界面撑爆。 */
const LABEL_MAX = 40

/**
 * 归一一个自由标签（严重程度、分类都用它）。
 *
 * 只做三件事：去首尾空白、去掉控制字符、限长。**不改大小写、不做白名单过滤**——
 * 调用方想用「严重」「P0」「blocker」「待确认」都行，插件不该替它决定词汇表。
 *
 * @param {any} v
 * @returns {string} 空串表示未设置
 */
function normLabel(v) {
  const raw = String(v ?? '')
  // 去掉控制字符（含换行/制表）：它们会破坏徽标的单行布局
  const cleaned = raw.replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim()
  return cleaned.length > LABEL_MAX ? cleaned.slice(0, LABEL_MAX) : cleaned
}

/**
 * 归一一个可选的自定义颜色。
 *
 * 允许的值：十六进制（#rgb/#rrggbb）、rgb()/hsl() 函数、以及常用颜色名。
 * 不合法的一律落成空串——面板会退回「按 severity 哈希配色」，
 * 而不是把一段奇怪的字符串塞进 style 里。
 *
 * @param {any} v
 * @returns {string} 合法时返回原值，否则空串
 */
function normColor(v) {
  const s = String(v ?? '').trim()
  if (!s || s.length > 40) return ''
  if (/^#(?:[0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(s)) return s
  if (/^(?:rgb|hsl)a?\([0-9.,%\s/]+\)$/i.test(s)) return s
  if (/^[a-z]{3,20}$/i.test(s)) return s
  return ''
}

/* ------------------------------------------------------------------ */
/* 服务                                                               */
/* ------------------------------------------------------------------ */

/**
 * 文档锚点视图服务。
 *
 * 公开方法即 Remote 面 + Agent 工具面（两边调用**同一份**实现，不各写一条逻辑）。
 */
export class DocAnchorViewer {
  /**
   * @param {object} ctx Cordis 上下文
   * @param {{dataDir?:string, maxDocxMb?:number}} config
   */
  constructor(ctx, config) {
    this.ctx = ctx
    this.config = config ?? {}
    this.dataDir = resolveDataDir(this.config.dataDir)
    this.store = new DocStore(this.dataDir)

    /**
     * Gateway 的服务发现直接读这个属性，而且必须挂在**实例**上。
     *
     * ⚠️ 只写 prototype 上的方法列表是不够的：Gateway 的 collectSrcClaims /
     * resolveBinding 都先要求 `Reflect.get(original, 'typertRemote')` 是个对象，
     * 并且 service === 实例、serviceKey / namespace 都对得上。少了这一条，
     * 整个命名空间根本不会进路由表，客户端调用表现为 **HTTP 404**——
     * 注意是「整个命名空间不存在」，不是「这个方法不存在」，
     * 所以很容易被误判成插件没装上或宿主没重载。
     */
    this.typertRemote = bindTypertRemote(this, 'docAnchorViewer', { namespace: 'docAnchor' })
  }

  /* ---------------- 内部：读 docx 并算结构 ---------------- */

  /**
   * 从字节解析出段落与大纲。
   *
   * ⚠️ 只解 ZIP 里**需要的两个条目**：解析只要 document.xml 与 styles.xml，
   *    图片、嵌入对象一概不解压——一份几十 MB 的招标文件里图片能占绝大多数体积，
   *    把它们全部 inflate 进内存既慢又危险（docx-preview 在浏览器里渲染时用的
   *    是**原始 docx 字节**，走的是 getDocxBytes，跟这里不是同一条路）。
   *
   * @param {Buffer} buf docx 字节
   * @returns {{paragraphs:Array, outline:Array, occurrence:Object, stats:object}}
   */
  _analyze(buf) {
    const entries = readZipBuffer(buf, { only: DOC_PARTS })
    const docXml = entries.get('word/document.xml')
    if (!docXml) throw new Error('DOCX_NO_DOCUMENT_XML：这不是一个有效的 .docx')
    const stylesXml = entries.get('word/styles.xml')

    const styles = parseStyles(stylesXml ? stylesXml.toString('utf8') : '')
    const paragraphs = parseParagraphs(docXml.toString('utf8'), styles)
    const outline = buildOutline(paragraphs)
    const occ = occurrenceMap(paragraphs)

    // occurrence 用普通对象（段落序号 → 第几次）序列化，体积比数组键值对小
    const occurrence = {}
    for (const [k, v] of occ) occurrence[k] = v

    const textChars = paragraphs.reduce((n, p) => n + (p.text ? p.text.length : 0), 0)
    return {
      paragraphs,
      outline,
      occurrence,
      stats: {
        paragraphs: paragraphs.length,
        headings: outline.length,
        chars: textChars,
        styles: styles.size,
      },
    }
  }

  /**
   * 由一条记录组装落盘文档。
   * @param {object} input
   * @returns {Promise<object>}
   */
  async _persist(input) {
    const { filePath, buf, stat } = input
    const docId = makeDocId(filePath, stat.mtimeMs)
    const analysis = this._analyze(buf)

    const prev = await this.store.readDoc(docId)
    const doc = {
      schemaVersion: 1,
      id: docId,
      path: filePath,
      title: input.title || path.basename(filePath, path.extname(filePath)),
      size: stat.size,
      mtimeMs: stat.mtimeMs,
      openedAt: new Date().toISOString(),
      stats: analysis.stats,
      // 段落只存匹配需要的两个字段，不存样式细节——面板不展示它们，
      // 存了只会让单篇文档 JSON 从几百 KB 涨到几 MB。
      paragraphs: analysis.paragraphs.map((p) => ({
        index: p.index,
        text: p.text,
        norm: p.norm,
        heading: p.heading,
        level: p.level,
      })),
      outline: analysis.outline,
      occurrence: analysis.occurrence,
      // 标记是用户的，重复登记同一份文件时要保住
      marks: Array.isArray(prev?.marks) ? prev.marks : [],
    }
    await this.store.saveDoc(doc)
    return doc
  }

  /* ---------------- Remote / 工具：读 ---------------- */

  /**
   * 诊断：数据目录、配置、能力。
   * @returns {Promise<object>}
   */
  async diagnostics() {
    return {
      ok: true,
      dataDir: this.dataDir,
      maxDocxMb: this.config.maxDocxMb,
      store: await this.store.diagnostics(),
    }
  }

  /**
   * 已登记的文档列表（最近打开在前）。
   * @returns {Promise<object>}
   */
  async listDocs() {
    const docs = await this.store.listDocs()
    return {
      ok: true,
      docs: docs.map((d) => ({
        id: d.id,
        title: d.title,
        path: d.path,
        openedAt: d.openedAt,
        size: d.size,
        stats: d.stats,
        markCount: d.markCount ?? 0,
      })),
    }
  }

  /**
   * 打开/登记一篇 docx。
   *
   * 两种来源：
   *   · filePath —— 服务器本地绝对路径（Agent 最常用）；
   *   · base64   —— 浏览器选文件后传上来的字节（面板「打开本地文件」用）。
   *
   * @param {object} request { filePath?, base64?, fileName?, title? }
   * @returns {Promise<object>}
   */
  async openDoc(request) {
    const input = unwrapArg(request)
    const maxBytes = Number(this.config.maxDocxMb ?? 30) * 1024 * 1024

    let filePath = stripQuotes(input.filePath)
    let buf = null
    let stat = null

    if (filePath) {
      let st
      try { st = fs.statSync(filePath) } catch {
        return fail('doc/file-not-found', '文件不存在：' + filePath)
      }
      if (!st.isFile()) return fail('doc/not-a-file', '路径不是文件：' + filePath)
      if (path.extname(filePath).toLowerCase() !== '.docx') {
        return fail('doc/unsupported-format', '只支持 .docx；旧版 .doc 请先用 Word 另存为 .docx')
      }
      if (st.size > maxBytes) {
        return fail('doc/file-too-large', `文件 ${(st.size / 1048576).toFixed(1)}MB 超过上限 ${this.config.maxDocxMb}MB`)
      }
      buf = fs.readFileSync(filePath)
      stat = { size: st.size, mtimeMs: st.mtimeMs }
    } else {
      const b64 = String(input.base64 ?? '')
      if (!b64) return fail('doc/invalid-input', '需要 filePath 或 base64 之一')
      if (b64.length > maxBytes * 1.4) {
        return fail('doc/file-too-large', `上传内容超过上限 ${this.config.maxDocxMb}MB`)
      }
      buf = Buffer.from(b64, 'base64')
      if (!buf.length) return fail('doc/invalid-input', 'base64 解码为空，文件可能没读全')
      if (buf.length > maxBytes) {
        return fail('doc/file-too-large', `文件 ${(buf.length / 1048576).toFixed(1)}MB 超过上限 ${this.config.maxDocxMb}MB`)
      }
      const name = String(input.fileName ?? input.title ?? 'uploaded.docx')
      // 上传的文件没有稳定路径。用内容哈希当路径标识，保证同一份文件重复打开命中同一条记录。
      const hash = createHash('sha1').update(buf).digest('hex').slice(0, 16)
      filePath = path.join(this.dataDir, 'uploads', hash + '-' + path.basename(name))
      try {
        fs.mkdirSync(path.dirname(filePath), { recursive: true })
        if (!fs.existsSync(filePath) || fs.statSync(filePath).size !== buf.length) {
          fs.writeFileSync(filePath, buf)
        }
      } catch (e) {
        return fail('doc/storage-failed', '保存上传文件失败：' + String(e?.message ?? e))
      }
      stat = { size: buf.length, mtimeMs: 0 }
    }

    let doc
    try {
      doc = await this._persist({ filePath, buf, stat, title: stripQuotes(input.title) })
    } catch (e) {
      const msg = String(e?.message ?? e)
      if (msg.startsWith('DOCX_NO_DOCUMENT_XML')) return fail('doc/not-a-docx', '这不是有效的 .docx（缺少 word/document.xml）')
      // 压缩炸弹 / 条目过大：给一句用户看得懂的话，而不是把 zlib 内部话端出去
      // （条目名来自不可信的 ZIP，绝不能原样透出到界面或日志）。
      if (msg.startsWith('ZIP_ENTRY_TOO_LARGE') || msg.startsWith('ZIP_TOTAL_TOO_LARGE')) {
        return fail('doc/parse-failed', '解析失败：文件里有解压后过大的条目（疑似压缩炸弹），已拒绝读取。')
      }
      if (msg.startsWith('ZIP_BAD_CDH') || msg.startsWith('ZIP_BAD_LFH') || msg.startsWith('ZIP_TRUNCATED')) {
        return fail('doc/parse-failed', '解析失败：文件已损坏或不是有效的 .docx（ZIP 结构不完整）。')
      }
      return fail('doc/parse-failed', '解析失败：' + msg)
    }

    return {
      ok: true,
      doc: {
        id: doc.id,
        title: doc.title,
        path: doc.path,
        openedAt: doc.openedAt,
        stats: doc.stats,
        markCount: doc.marks.length,
      },
    }
  }

  /**
   * 一篇文档的元信息 + 大纲（不含全文段落）。
   * @param {object} request { docId }
   * @returns {Promise<object>}
   */
  async getDoc(request) {
    const docId = safeDocId(unwrapArg(request).docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const doc = await this.store.readDoc(docId)
    if (!doc) return fail('doc/not-found', '文档不存在：' + docId)
    return {
      ok: true,
      doc: {
        id: doc.id,
        title: doc.title,
        path: doc.path,
        size: doc.size,
        openedAt: doc.openedAt,
        stats: doc.stats,
        outline: doc.outline ?? [],
        marks: doc.marks ?? [],
      },
    }
  }

  /**
   * 按**绝对路径**查一篇已登记的文档。
   *
   * 侧栏标签页拿到的只有文件路径（地址里带来的），不知道 docId；
   * 它需要知道「这篇文档在我这儿登记过没有」——登记过才显示大纲与标记，
   * 没登记过就退化成普通预览。没有这个方法的话，标签页就得先 listDocs
   * 再自己比对路径，多一次往返、还要在客户端复制 Windows 路径规范化逻辑。
   *
   * @param {object} request { filePath }
   * @returns {Promise<object>} 找到时同 getDoc；找不到时 { ok:true, found:false }
   */
  async getDocByPath(request) {
    const filePath = stripQuotes(unwrapArg(request).filePath)
    if (!filePath) return fail('doc/invalid-input', '缺少 filePath')
    // 路径比较：Windows 忽略大小写与分隔符差异；两边都归一成 / 再比
    const norm = (v) => {
      const t = String(v ?? '').replace(/\\/g, '/').replace(/\/+$/, '')
      return process.platform === 'win32' ? t.toLowerCase() : t
    }
    const docs = await this.store.listDocs()
    const hit = docs.find((d) => norm(d.path) === norm(filePath))
    if (!hit) return { ok: true, found: false, filePath }
    const full = await this.getDoc({ docId: hit.id })
    if (!full.ok) return { ok: true, found: false, filePath }
    return { ...full, found: true }
  }

  /**
   * 只取大纲（Agent 规划阅读顺序时最常用，比 getDoc 更省 token）。
   * @param {object} request { docId }
   * @returns {Promise<object>}
   */
  async getOutline(request) {
    const docId = safeDocId(unwrapArg(request).docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const doc = await this.store.readDoc(docId)
    if (!doc) return fail('doc/not-found', '文档不存在：' + docId)
    return { ok: true, docId: doc.id, title: doc.title, outline: doc.outline ?? [], stats: doc.stats }
  }

  /**
   * 按锚点/关键词取原文段落。
   *
   * 这是「锚点」对 Agent 的那一面：给定某条大纲的 index，或一段文字，
   * 取回它前后若干段——模型据此读上下文，而不是把整篇几百万字塞进上下文。
   *
   * @param {object} request { docId, index?, query?, before?, after?, maxChars? }
   * @returns {Promise<object>}
   */
  async readSection(request) {
    const input = unwrapArg(request)
    const docId = safeDocId(input.docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const doc = await this.store.readDoc(docId)
    if (!doc) return fail('doc/not-found', '文档不存在：' + docId)

    const paras = doc.paragraphs ?? []
    const before = Math.max(0, Math.min(200, Number(input.before ?? 2) || 0))
    const after = Math.max(0, Math.min(400, Number(input.after ?? 12) || 0))
    const maxChars = Math.max(200, Math.min(40000, Number(input.maxChars ?? 6000) || 6000))

    // 定位起点：优先 index，其次 query（精确 → 归一化包含 → 包含）
    let start = -1
    if (input.index !== undefined && input.index !== null && input.index !== '') {
      const n = Number(input.index)
      if (Number.isFinite(n)) start = paras.findIndex((p) => p.index === n)
    }
    if (start < 0 && input.query) {
      const q = String(input.query)
      const qn = q.replace(/\s+/g, '').trim()
      start = paras.findIndex((p) => p.text === q)
      if (start < 0) start = paras.findIndex((p) => p.norm === qn)
      if (start < 0) start = paras.findIndex((p) => p.norm.includes(qn))
      if (start < 0) start = paras.findIndex((p) => p.text.includes(q))
    }
    if (start < 0) {
      if (input.query || input.index !== undefined) {
        return fail('doc/anchor-not-found', '在这篇文档里没找到该锚点：' + (input.query ?? input.index))
      }
      start = 0
    }

    const from = Math.max(0, start - before)
    const to = Math.min(paras.length, start + after + 1)
    const picked = []
    let used = 0
    let truncated = false
    for (let i = from; i < to; i++) {
      const p = paras[i]
      const line = p.text ?? ''
      if (used + line.length > maxChars && picked.length) { truncated = true; break }
      picked.push({ index: p.index, text: line, heading: !!p.heading, level: p.level ?? 0 })
      used += line.length
    }
    return {
      ok: true,
      docId: doc.id,
      title: doc.title,
      anchorIndex: paras[start].index,
      total: paras.length,
      truncated,
      paragraphs: picked,
    }
  }

  /**
   * 取原始 docx 字节（base64），供浏览器端 docx-preview 渲染。
   *
   * 为什么必须是原始字节：docx-preview 自己解 ZIP、还原字体/编号/分页/表格，
   * 给它任何「已解析的中间结构」都还原不出 Word 的样子。
   *
   * @param {object} request { docId }
   * @returns {Promise<object>}
   */
  async getDocxBytes(request) {
    const docId = safeDocId(unwrapArg(request).docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const doc = await this.store.readDoc(docId)
    if (!doc) return fail('doc/not-found', '文档不存在：' + docId)
    if (!doc.path || !fs.existsSync(doc.path)) {
      return fail('doc/file-missing', '源文件已不在原位置：' + (doc.path || '(未记录)'))
    }
    const maxBytes = Number(this.config.maxDocxMb ?? 30) * 1024 * 1024
    let buf
    try { buf = fs.readFileSync(doc.path) } catch (e) {
      return fail('doc/read-failed', '读取源文件失败：' + String(e?.message ?? e))
    }
    if (buf.length > maxBytes) {
      return fail('doc/file-too-large', `文件 ${(buf.length / 1048576).toFixed(1)}MB 超过上限 ${this.config.maxDocxMb}MB`)
    }
    const hash = createHash('sha1').update(buf).digest('hex')
    return { ok: true, docId: doc.id, title: doc.title, size: buf.length, sha1: hash, base64: buf.toString('base64') }
  }

  /* ---------------- Remote / 工具：标记 ---------------- */

  /**
   * 取一份「可以拿来总结」的压缩底稿。
   *
   * 为什么不是让模型自己 read_section 十几次：
   *   一份几千段的技术文档，逐段读会把上下文吃满，而且每一段都是同等的
   *   篇幅——但总结真正需要的是**骨架 + 每节的开头**。这里按大纲把文档切成
   *   若干节，每节只取前若干段（截断到 maxCharsPerSection），整体再按
   *   maxChars 收敛。模型拿到它就能写出结构化的总结，并**顺手拿到每节的
   *   段落序号**——那正是打锚点要用的 index。
   *
   * 返回的 sections[].anchor 可以直接喂给 set_marks，
   * 不用模型自己猜序号或复制原文。
   *
   * @param {object} request { docId, maxChars?, maxCharsPerSection?, maxSections? }
   * @returns {Promise<object>}
   */
  async getDigest(request) {
    const input = unwrapArg(request)
    const docId = safeDocId(input.docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const doc = await this.store.readDoc(docId)
    if (!doc) return fail('doc/not-found', '文档不存在：' + docId)

    const maxChars = Math.max(1000, Math.min(60000, Number(input.maxChars ?? 12000) || 12000))
    const perSection = Math.max(120, Math.min(4000, Number(input.maxCharsPerSection ?? 600) || 600))
    const maxSections = Math.max(1, Math.min(400, Number(input.maxSections ?? 120) || 120))

    const paras = doc.paragraphs ?? []
    const outline = doc.outline ?? []

    // 章节边界：大纲条目各自的 index。没有大纲时整篇当一节。
    const bounds = outline.length
      ? outline.map((o) => o.index)
      : [paras[0]?.index ?? 0]

    const sections = []
    let used = 0
    for (let s = 0; s < bounds.length && sections.length < maxSections; s++) {
      const start = bounds[s]
      const end = s + 1 < bounds.length ? bounds[s + 1] : paras.length
      const head = outline[s] ?? null
      const body = []
      let secUsed = 0
      for (let i = start; i < end; i++) {
        const p = paras[i]
        if (!p) continue
        // 跳过标题本身在正文里的重复
        if (head && p.index === head.index) continue
        const line = String(p.text ?? '').trim()
        if (!line) continue
        if (secUsed + line.length > perSection) {
          if (!body.length) body.push(line.slice(0, perSection))
          break
        }
        body.push(line)
        secUsed += line.length
      }
      const text = body.join('\n')
      if (used + text.length > maxChars && sections.length) {
        sections.push({ truncatedAt: sections.length })
        break
      }
      used += text.length
      sections.push({
        index: sections.length,
        anchorIndex: head ? head.index : start,
        level: head ? head.level : 0,
        title: head ? head.title : '（无标题正文）',
        norm: head ? head.norm : '',
        occurrence: head ? head.occurrence : 0,
        text,
      })
    }

    return {
      ok: true,
      docId: doc.id,
      title: doc.title,
      path: doc.path,
      stats: doc.stats,
      sectionCount: sections.length,
      outlineCount: outline.length,
      sections,
    }
  }

  /**
   * 新增一条标记。
   *
   * 标记 = 「用户/AI 认为重要的一处原文」，带锚点（段落序号 + 归一化文本 +
   * 同文第几次出现），面板据此跳转，Agent 据此知道该回去看哪一段。
   *
   * @param {object} request { docId, text, note?, anchor? }
   * @returns {Promise<object>}
   */
  async addMark(request) {
    const input = unwrapArg(request)
    const docId = safeDocId(input.docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const text = String(input.text ?? '').trim()
    if (!text) return fail('doc/invalid-input', '标记内容不能为空')

    const r = await this.store.mutateDoc(String(docId), (doc) => {
      const marks = Array.isArray(doc.marks) ? doc.marks : []
      const anchor = this._normalizeAnchor(doc, input.anchor, text)
      const mark = {
        id: 'm' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        text,
        note: String(input.note ?? '').trim(),
        // 评判三件套：严重程度 / 条款类别 / 建议怎么改。
        // 都留成可选：用户自己的标记不需要「严重程度」，只有 AI 评判时才带。
        severity: normLabel(input.severity),
        category: normLabel(input.category),
        // 可选：调用方直接指定颜色（十六进制/rgb/hsl/颜色名）。
        // 留空则由面板按 severity 文案哈希生成——同一个词永远同一个颜色。
        color: normColor(input.color),
        suggestion: String(input.suggestion ?? '').trim(),
        anchor,
        createdAt: new Date().toISOString(),
      }
      return { doc: { ...doc, marks: [...marks, mark] }, result: mark }
    })
    if (!r.ok) return fail('doc/not-found', r.error)
    return { ok: true, mark: r.result }
  }

  /**
   * 批量写标记：一次提交一组锚点。
   *
   * 为什么需要一个**批量**入口，而不是让调用方循环 addMark：
   *   1. 「总结一篇文档并标出要点」是一次语义动作，模型的产出天然是一组
   *      （要点 + 每条对应的原文），不是一条一条挤出来的。逐条调用会让一次
   *      总结变成十几次往返，且中途失败会留下一半的锚点。
   *   2. 批量写是**原子的**：任何一条锚点解析失败，整组不落盘，用户不会看到
   *      「标了一半」的文档。
   *   3. 它还能一次回传逐条的命中情况（哪条锚到了第几段、哪条没找到原文），
   *      调用方据此决定要不要重试那几条——这是逐条调用给不了的信息。
   *
   * @param {object} request { docId, marks: Array<{text, note?, anchor?}>, replace? }
   * @returns {Promise<object>}
   */
  async setMarks(request) {
    const input = unwrapArg(request)
    const docId = safeDocId(input.docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')

    const list = Array.isArray(input.marks) ? input.marks : null
    if (!list) return fail('doc/invalid-input', 'marks 必须是数组')
    if (list.length > 200) return fail('doc/invalid-input', '一次最多写 200 条标记')

    const replace = input.replace === true
    const now = new Date().toISOString()

    const r = await this.store.mutateDoc(docId, (doc) => {
      const existing = Array.isArray(doc.marks) ? doc.marks : []
      const added = []
      const rejected = []

      for (let i = 0; i < list.length; i++) {
        const item = list[i] ?? {}
        const text = String(item.text ?? '').trim()
        if (!text) { rejected.push({ i, reason: 'text 为空' }); continue }

        // 锚点两种写法都接受：嵌套 { anchor: {index,norm,occurrence} } 或平铺的三字段。
        // 平铺是工具面的自然写法（参数表里逐项列出），嵌套是存储形状；
        // 只认一种的话，另一种写法会静默退化成「按文字找第一处」——同文多处时直接锚错。
        const rawAnchor = item.anchor ?? (
          (item.index !== undefined || item.norm !== undefined || item.occurrence !== undefined)
            ? { index: item.index, norm: item.norm, occurrence: item.occurrence }
            : undefined
        )
        const anchor = this._normalizeAnchor(doc, rawAnchor, text)
        // 锚点没落到任何一段 → 这条标记点了会跳不动，宁可退回让调用方改
        if (anchor.index < 0) {
          rejected.push({ i, text: text.slice(0, 60), reason: '在文档里找不到这段原文（锚点只做精确匹配）' })
          continue
        }
        added.push({
          id: 'm' + Date.now().toString(36) + i.toString(36) + Math.random().toString(36).slice(2, 5),
          text,
          note: String(item.note ?? '').trim(),
          severity: normLabel(item.severity),
          category: normLabel(item.category),
          color: normColor(item.color),
          suggestion: String(item.suggestion ?? '').trim(),
          anchor,
          createdAt: now,
        })
      }

      // 有任何一条被拒 → 整组不落盘（原子性），把原因如实带回去
      if (rejected.length) return { result: { added: [], rejected, applied: false } }

      const merged = replace ? added : [...existing, ...added]
      return {
        doc: { ...doc, marks: merged },
        result: { added, rejected: [], applied: true, total: merged.length },
      }
    })

    if (!r.ok) return fail('doc/not-found', r.error)
    const out = r.result
    if (!out.applied) {
      return {
        ok: false,
        error: {
          code: 'doc/marks-partially-invalid',
          message: `有 ${out.rejected.length} 条标记对不上原文，整组未写入：`
            + out.rejected.map((x) => (x.text ? `「${x.text}」${x.reason}` : x.reason)).join('；'),
        },
        rejected: out.rejected,
      }
    }
    return { ok: true, added: out.added.length, total: out.total, marks: out.added }
  }

  /**
   * 删掉一条标记。
   * @param {object} request { docId, markId }
   * @returns {Promise<object>}
   */
  async removeMark(request) {
    const args = unwrapArg(request)
    const docId = safeDocId(args.docId)
    const markId = String(args.markId ?? '').trim()
    if (!docId || !markId) return fail('doc/invalid-input', '需要 docId 与 markId')
    const r = await this.store.mutateDoc(String(docId), (doc) => {
      const marks = Array.isArray(doc.marks) ? doc.marks : []
      const next = marks.filter((m) => m.id !== markId)
      if (next.length === marks.length) return { result: { removed: false } }
      return { doc: { ...doc, marks: next }, result: { removed: true } }
    })
    if (!r.ok) return fail('doc/not-found', r.error)
    if (!r.result?.removed) return fail('doc/mark-not-found', '标记不存在：' + markId)
    return { ok: true, removed: true }
  }

  /**
   * 列出标记。
   * @param {object} request { docId }
   * @returns {Promise<object>}
   */
  async listMarks(request) {
    const docId = safeDocId(unwrapArg(request).docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const doc = await this.store.readDoc(docId)
    if (!doc) return fail('doc/not-found', '文档不存在：' + docId)
    return { ok: true, docId: doc.id, marks: doc.marks ?? [] }
  }

  /**
   * 把不完整/缺失的锚点补齐。
   *
   * 锚点三要素缺一不可（序号未命中时客户端会退化为全文精确匹配，但同文
   * 出现多次就必须靠 occurrence 区分）。补齐时**只认精确匹配**——模糊匹配
   * 会把「联系人」这种短词锚到一条长段上，这是明确的错位来源。
   *
   * @param {object} doc
   * @param {any} given
   * @param {string} text
   * @returns {{index:number, norm:string, occurrence:number}}
   */
  _normalizeAnchor(doc, given, text) {
    const paras = doc.paragraphs ?? []
    const g = (given && typeof given === 'object') ? given : {}

    let index = Number.isFinite(Number(g.index)) ? Number(g.index) : null
    let norm = String(g.norm ?? '').trim() || String(text).replace(/\s+/g, '').trim()
    let occurrence = Number.isFinite(Number(g.occurrence)) ? Number(g.occurrence) : null

    // 序号有效且文本一致 → 直接用
    const at = index === null ? -1 : paras.findIndex((p) => p.index === index)
    const consistent = at >= 0 && (!norm || paras[at].norm === norm)

    if (!consistent) {
      // 按归一化全文精确匹配；有 occurrence 时取**第 n 次**出现，
      // 否则同文多处的标记会全部挤到第一处。
      const want = occurrence ?? 0
      let seen = -1
      let hit = -1
      for (let i = 0; i < paras.length; i++) {
        if (paras[i].norm !== norm) continue
        seen++
        if (seen === want) { hit = i; break }
      }
      if (hit < 0 && seen >= 0) hit = paras.findIndex((p) => p.norm === norm) // 出现次数不足时退回第一处
      if (hit >= 0) {
        index = paras[hit].index
        norm = paras[hit].norm
        occurrence = want
      } else if (at >= 0) {
        // 序号有效但文本对不上（用户手改了标记文字）：保留序号，文本以原文为准
        norm = paras[at].norm
        occurrence = doc.occurrence?.[paras[at].index] ?? 0
      }
    }
    if (occurrence === null) occurrence = index !== null ? (doc.occurrence?.[index] ?? 0) : 0
    return { index: index === null ? -1 : index, norm, occurrence }
  }

  /**
   * 删掉一篇文档（只删本插件登记，不动源文件）。
   * @param {object} request { docId }
   * @returns {Promise<object>}
   */
  async removeDoc(request) {
    const docId = safeDocId(unwrapArg(request).docId)
    if (!docId) return fail('doc/invalid-input', '缺少或非法的 docId')
    const removed = await this.store.removeDoc(docId)
    if (!removed) return fail('doc/not-found', '文档不存在：' + docId)
    return { ok: true, removed: true }
  }
}

/* ------------------------------------------------------------------ */
/* 插件入口                                                            */
/* ------------------------------------------------------------------ */

/** 暴露给浏览器的 Remote 方法名（必须与 lib/client.js 的 descriptors 逐位一致）。 */
export const REMOTE_METHODS = [
  'diagnostics',
  'listDocs',
  'openDoc',
  'getDoc',
  'getDocByPath',
  'getOutline',
  'readSection',
  'getDocxBytes',
  'getDigest',
  'addMark',
  'setMarks',
  'removeMark',
  'listMarks',
  'removeDoc',
]

const REMOTE_METHODS_KEY = '@deepseek-ai/dsh-typert-protocol/remote-methods'

/**
 * 把一个 Service 绑定到它的 Remote 命名空间。
 *
 * 协议包存在时用官方的 `bindTypertRemote`；这里手写同样三个字段——
 * Gateway 的发现逻辑（collectSrcClaims / readBinding）读的就是这几处，
 * 所以契约仍然成立。
 *
 * @param {object} service 服务实例
 * @param {string} serviceKey Cordis 服务键
 * @param {{namespace:string}} options 线上市名空间
 * @returns {{service:object, serviceKey:string, namespace:string}}
 */
function bindTypertRemote(service, serviceKey, options) {
  return { service, serviceKey, namespace: options.namespace }
}

/**
 * 把方法标记为可 Remote 调用（纯 JS 没有装饰器，手写同样的两个字段）。
 * @param {object} instance
 * @param {string[]} methods
 */
function markRemoteMethods(instance, methods) {
  const proto = Object.getPrototypeOf(instance)
  if (!proto) return
  Object.defineProperty(proto, REMOTE_METHODS_KEY, {
    configurable: true,
    value: Object.freeze({
      version: 1,
      methods: Object.freeze(methods.map((m) => Object.freeze({ method: m, invocation: Object.freeze({ kind: 'direct' }) }))),
    }),
  })
}

/**
 * 插件入口：服务 + Remote 面 + Agent 工具。
 *
 * @param {object} ctx
 * @param {object} config
 */
export function apply(ctx, config) {
  const svc = new DocAnchorViewer(ctx, config)
  ctx.provide('docAnchorViewer')
  ctx.set('docAnchorViewer', svc)
  markRemoteMethods(svc, REMOTE_METHODS)

  // Agent 工具：可选注册（宿主没有 tools 服务时插件其余部分照常工作）
  ctx.inject(['tools'], (toolsCtx) => {
    const disposers = buildDocTools(toolsCtx, svc)
    toolsCtx.effect(() => () => {
      for (const d of disposers) { try { d() } catch { /* ignore */ } }
    }, 'docanchor:tools')
  })

  return () => { /* ctx.set 由 Cordis 在 fiber 销毁时回收 */ }
}

/* ------------------------------------------------------------------ */
/* Agent 工具                                                          */
/* ------------------------------------------------------------------ */

/**
 * 构造并注册工具。
 *
 * 为什么是**一个**工具 + action 参数，而不是五六个小工具：
 * 这些动作共享同一组参数语义（docId/锚点）与同一份失败含义
 * （「文档不存在」该去先 open_doc），拆成多个只会让模型在每个工具上都学一遍。
 *
 * @param {object} ctx 已含 tools 的上下文
 * @param {DocAnchorViewer} svc
 * @returns {Array<() => void>}
 */
export function buildDocTools(ctx, svc) {
  const tools = ctx.get('tools')
  if (!tools) return []

  const compile = (spec) => {
    const properties = {}
    const required = []
    for (const [k, v] of Object.entries(spec ?? {})) {
      const { required: req, ...rest } = v ?? {}
      properties[k] = rest
      if (req === true) required.push(k)
    }
    return { type: 'object', properties, ...(required.length ? { required } : {}) }
  }
  const text = (t) => [{ type: 'text', text: t }]

  /**
   * 把标记的标签压成一行，如 "[fatal/qualification]"。
   *
   * ⚠️ 不做白名单：severity / category 是**自由文本**，原样显示。
   * 模型写「废标」「P0」「blocker」都能读懂，插件不该替调用方决定词汇表。
   *
   * @param {{severity?:string, category?:string}} m
   * @returns {string} 没有标签时返回空串
   */
  function fmtTag(m) {
    const sev = String(m?.severity ?? '').trim()
    const cat = String(m?.category ?? '').trim()
    if (!sev && !cat) return ''
    return '[' + [sev, cat].filter(Boolean).join('/') + ']'
  }

  /**
   * 把结果渲染成给人/模型看的纯文本。
   * @param {any} v
   * @returns {string}
   */
  const render = (v) => {
    if (!v || v.ok === false) {
      return '失败：' + (v?.error?.message ?? '未知错误') + (v?.error?.code ? `（${v.error.code}）` : '')
    }
    const lines = []
    if (Array.isArray(v.docs)) {
      lines.push(`共 ${v.docs.length} 篇文档：`)
      for (const d of v.docs) {
        lines.push(`- ${d.id}  ${d.title}  [${d.stats?.paragraphs ?? '?'} 段 / ${d.stats?.headings ?? '?'} 标题 / ${d.markCount ?? 0} 标记]`)
      }
    } else if (v.doc) {
      // open 的返回同时带 doc 与 outline：先报文档，再附大纲（模型下一步就要用）
      const d = v.doc
      lines.push(`已登记 ${d.id}：${d.title}`)
      lines.push(`路径：${d.path}`)
      lines.push(`统计：${d.stats?.paragraphs ?? '?'} 段 / ${d.stats?.headings ?? '?'} 标题 / ${(v.outline ?? []).length} 条大纲`)
      const ol = v.outline ?? []
      lines.push(ol.length ? '大纲：' : '大纲：这篇文档没有识别到标题样式。')
      for (const o of ol) {
        lines.push(`${'  '.repeat(Math.max(0, (o.level ?? 1) - 1))}- [${o.index}] ${o.title}`)
      }
    } else if (v.outline) {
      lines.push(`${v.title ?? v.docId}：${v.stats?.paragraphs ?? '?'} 段，${v.outline.length} 条大纲`)
      for (const o of v.outline) {
        lines.push(`${'  '.repeat(Math.max(0, (o.level ?? 1) - 1))}- [${o.index}] ${o.title}`)
      }
    } else if (Array.isArray(v.paragraphs)) {
      lines.push(`${v.title}（锚点第 ${v.anchorIndex} 段，共 ${v.total} 段${v.truncated ? '，已截断' : ''}）`)
      for (const p of v.paragraphs) {
        lines.push(`[${p.index}]${p.heading ? ' #' : ''} ${p.text}`)
      }
    } else if (Array.isArray(v.sections)) {
      lines.push(`${v.title}：${v.stats?.paragraphs ?? '?'} 段，按大纲切成 ${v.sectionCount} 节（共 ${v.outlineCount} 条大纲）`)
      lines.push('每节形如：- [段落序号] 标题 … 正文节选。要把某节标成锚点，用 set_marks 并带上该节的 anchorIndex/norm/occurrence。')
      for (const sec of v.sections) {
        if (sec.truncatedAt !== undefined) { lines.push(`…（超出篇幅上限，后面的节已省略）`); break }
        lines.push('')
        lines.push(`- [${sec.anchorIndex}] L${sec.level} ${sec.title}`)
        if (sec.text) lines.push(sec.text)
      }
    } else if (v.added !== undefined) {
      // ⚠️ 必须排在 marks 数组分支**前面**：set_marks 的返回也带 marks，
      //    顺序反了会把「已写入 N 条」误显示成「共 N 条标记」。
      lines.push(`已写入 ${v.added} 条标记（该文档现有 ${v.total} 条）。`)
      for (const m of v.marks ?? []) {
        lines.push('- ' + [m.id, '@' + (m.anchor?.index ?? '?'), fmtTag(m)].filter(Boolean).join(' ') + '：' + m.text
          + (m.note ? `（${m.note}）` : '')
          + (m.suggestion ? ` → ${m.suggestion}` : ''))
      }
    } else if (Array.isArray(v.marks)) {
      lines.push(`共 ${v.marks.length} 条标记：`)
      for (const m of v.marks) {
        lines.push('- ' + [m.id, '@' + (m.anchor?.index ?? '?'), fmtTag(m)].filter(Boolean).join(' ') + '：' + m.text
          + (m.note ? `（${m.note}）` : '')
          + (m.suggestion ? ` → ${m.suggestion}` : ''))
      }
    } else if (v.mark) {
      lines.push('已加标记 ' + [v.mark.id, '@' + (v.mark.anchor?.index ?? '?'), fmtTag(v.mark)].filter(Boolean).join(' ')
        + '：' + v.mark.text
        + (v.mark.note ? `（${v.mark.note}）` : '')
        + (v.mark.suggestion ? ` → ${v.mark.suggestion}` : ''))
    } else {
      lines.push(JSON.stringify(v))
    }
    return lines.join('\n')
  }

  const disposers = []

  disposers.push(tools.register({
    name: 'doc_viewer',
    description: '文档锚点视图（通用）：打开 docx 并按 Word 版式在面板里预览，附带可跳转的大纲锚点与自定义标记。'
      + '标记是通用的——严重程度与分类都是**自由文本**，不限于招投标；合同审阅、方案评审、病历批注、'
      + '会议纪要都能用同一套机制。'
      + '通读一篇文档的标准走法：open → get_digest（按大纲分节、每节带好锚点的压缩底稿）→ 写出要点/判定 → '
      + 'set_marks（批量写入，用户在面板里点一条就跳到对应段落）。'
      + '也可以用 get_outline / read_section 按锚点读原文，不必把整篇文档读进上下文。'
      + '只读用户的文档，只写本插件自己的数据目录。',
    parameters: compile({
      action: {
        type: 'string',
        required: true,
        enum: ['open', 'list_docs', 'get_outline', 'get_digest', 'read_section', 'add_mark', 'set_marks', 'list_marks', 'remove_mark', 'remove_doc', 'diagnostics'],
        description: 'open=登记并预览一篇 docx；list_docs=已登记文档；get_outline=取大纲锚点；'
          + 'get_digest=取用于总结的压缩底稿（按大纲分段、每节截断，带每节的段落序号）；'
          + 'read_section=按锚点读原文段落；add_mark=加一条标记；set_marks=批量写标记（总结后一次落锚用这个）；'
          + 'list_marks/remove_mark=读写标记；remove_doc=移除登记；diagnostics=诊断',
      },
      filePath: { type: 'string', description: 'open 时必填：docx 的绝对路径' },
      docId: { type: 'string', description: '除 open / list_docs / diagnostics 外必填：文档 id（由 open 返回）' },
      title: { type: 'string', description: 'open 时可选：显示名，默认取文件名' },
      index: { type: 'number', description: 'read_section 时可选：从哪一段开始（大纲里的 index）' },
      query: { type: 'string', description: 'read_section 时可选：用一段原文定位（与 index 二选一）' },
      before: { type: 'number', description: 'read_section：向前取几段，默认 2' },
      after: { type: 'number', description: 'read_section：向后取几段，默认 12' },
      maxChars: { type: 'number', description: 'read_section：返回字符上限，默认 6000；get_digest：整份底稿上限，默认 12000' },
      maxCharsPerSection: { type: 'number', description: 'get_digest：每节最多取多少字符，默认 600' },
      maxSections: { type: 'number', description: 'get_digest：最多返回几节，默认 120' },
      text: { type: 'string', description: 'add_mark 时必填：要标记的原文文字' },
      note: { type: 'string', description: 'add_mark 时可选：备注（评判时写结论）' },
      severity: {
        type: 'string',
        description: '这条标记的严重程度，**自由文本**：写什么就显示什么（推荐 fatal/high/medium/low/info，'
          + '面板自带中文标签；也可以写「P0」「严重」「blocker」等任意词）。颜色按词自动生成，也可用 color 指定。不评判时省略。',
      },
      category: {
        type: 'string',
        description: '分类标签，**自由文本**：写什么就显示什么（推荐 invalid/qualification/scoring/technical/'
          + 'commercial/contract/other，面板自带中文标签；也可以写「风险点」「待确认」「第三章」等任意词）。',
      },
      color: {
        type: 'string',
        description: '可选：这条标记的颜色，十六进制（#e0574b）或 rgb()/hsl()/颜色名。省略则按 severity 自动配色。',
      },
      suggestion: { type: 'string', description: '可选：建议怎么改 / 该注意什么（与 note 分开，note 写结论，这里写行动）' },
      markId: { type: 'string', description: 'remove_mark 时必填：标记 id' },
      marks: {
        type: 'array',
        description: 'set_marks 时必填：标记数组。每项 { text, note?, anchor? }。'
          + 'text 必须是文档里**逐字存在**的一段原文（只做精确匹配，不做模糊包含）；'
          + 'anchor 可选，写 { index, norm, occurrence } 可精确指定第几段。'
          + 'get_digest 返回的 sections[].anchorIndex / norm / occurrence 可以直接拿来填。',
        items: {
          type: 'object',
          properties: {
            text: { type: 'string', description: '原文文字（逐字）' },
            note: { type: 'string', description: '这条锚点的结论（评判时就是判定理由）' },
            severity: {
              type: 'string',
              description: '严重程度，自由文本（推荐 fatal/high/medium/low/info，也可自定义）',
            },
            category: {
              type: 'string',
              description: '分类标签，自由文本（推荐 invalid/qualification/scoring/technical/commercial/contract/other，也可自定义）',
            },
            color: { type: 'string', description: '可选：颜色（#hex / rgb() / hsl() / 颜色名），省略则按 severity 自动配色' },
            suggestion: { type: 'string', description: '建议怎么改 / 该注意什么' },
            index: { type: 'number', description: '可选：段落序号' },
            norm: { type: 'string', description: '可选：去掉空白后的全文' },
            occurrence: { type: 'number', description: '可选：同文第几次出现（0 起）' },
          },
          required: ['text'],
        },
      },
      replace: { type: 'boolean', description: 'set_marks 时可选：true 表示替换掉已有标记，默认追加' },
    }),
    output: {
      schema: { type: 'object' },
      render: (_args, v) => text(render(v)),
    },
    async execute(args = {}) {
      switch (args.action) {
        case 'open': {
          if (!args.filePath) return fail('doc/invalid-input', 'open 需要 filePath')
          const r = await svc.openDoc({ filePath: args.filePath, title: args.title })
          if (r.ok) {
            // 打开后顺手把大纲带回去：模型下一步多半就要用它，省一次往返
            const o = await svc.getOutline({ docId: r.doc.id })
            return { ...r, outline: o.ok ? o.outline : [] }
          }
          return r
        }
        case 'list_docs': return svc.listDocs()
        case 'get_outline': return svc.getOutline({ docId: args.docId })
        case 'read_section':
          return svc.readSection({
            docId: args.docId, index: args.index, query: args.query,
            before: args.before, after: args.after, maxChars: args.maxChars,
          })
        case 'get_digest':
          return svc.getDigest({
            docId: args.docId, maxChars: args.maxChars,
            maxCharsPerSection: args.maxCharsPerSection, maxSections: args.maxSections,
          })
        case 'add_mark':
          return svc.addMark({
            docId: args.docId, text: args.text, note: args.note,
            severity: args.severity, category: args.category,
            color: args.color, suggestion: args.suggestion,
          })
        case 'set_marks': {
          const items = Array.isArray(args.marks) ? args.marks : []
          return svc.setMarks({
            docId: args.docId,
            replace: args.replace === true,
            marks: items.map((m) => ({
              text: m?.text,
              note: m?.note,
              severity: m?.severity,
              category: m?.category,
              color: m?.color,
              suggestion: m?.suggestion,
              // 扁平参数（index/norm/occurrence）收成 anchor 对象，与 Host 的锚点形状一致
              anchor: (m && (m.index !== undefined || m.norm !== undefined || m.occurrence !== undefined))
                ? { index: m.index, norm: m.norm, occurrence: m.occurrence }
                : undefined,
            })),
          })
        }
        case 'list_marks': return svc.listMarks({ docId: args.docId })
        case 'remove_mark': return svc.removeMark({ docId: args.docId, markId: args.markId })
        case 'remove_doc': return svc.removeDoc({ docId: args.docId })
        case 'diagnostics': return svc.diagnostics()
        default:
          return fail('doc/invalid-input', '未知的 action：' + args.action)
      }
    },
  }))

  return disposers
}
