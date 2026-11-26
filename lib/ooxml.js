/**
 * 文档锚点视图 · OOXML 底层工具
 *
 * 只做两件事：解 DOCX 的 ZIP 容器、把 document.xml 拆成「按文档顺序的段落序列」。
 * 刻意不引第三方依赖（ZIP 用 node:zlib），这样整个插件可以整包搬走、离线可用。
 *
 * 为什么不用 DOM 解析器：一份 3.8MB 的 document.xml 用 DOM 峰值 37MB，
 * 而这里的单遍扫描只与嵌套深度有关。招标文件动辄几千段，内存差一个量级。
 */

import { inflateRawSync } from 'node:zlib'

/* ------------------------------------------------------------------ */
/* 一、ZIP 读取                                                         */
/* ------------------------------------------------------------------ */

/**
 * 单个条目**解压后**的字节上限。
 *
 * ⚠️ 这个上限不是「防御性编程」，是**必须**的：docx 就是 ZIP，而 ZIP 的压缩比
 * 可以高得离谱。实测一个 65KB 的条目能解压出 64MB（1024×）——「文件 ≤30MB」
 * 这道检查完全挡不住它，因为被撑爆的是**解压后**的内存。
 *
 * Node 的 `maxOutputLength` 是**边解边停**：超限时立刻抛 ERR_BUFFER_TOO_LARGE，
 * 不会先把 64MB 分配出来（实测 6ms 抛错，堆增量 ≈ 0）。
 */
export const MAX_ENTRY_BYTES = 64 * 1024 * 1024

/**
 * 一次解压**全部条目**的总字节上限。
 *
 * 单条目上限挡不住「很多个中等条目累加」这一种：1000 个 1MB 的条目也能撑爆。
 * 有了总预算，一次读取的内存占用就有了确定上界。
 */
export const MAX_TOTAL_BYTES = 256 * 1024 * 1024

/**
 * 解压一个条目，超限就抛可读的错误。
 *
 * @param {Buffer} raw 压缩后的字节
 * @param {string} name 条目名（只进错误信息）
 * @param {number} maxBytes 上限
 * @returns {Buffer}
 */
function inflateEntry(raw, name, maxBytes) {
  try {
    return inflateRawSync(raw, { maxOutputLength: maxBytes })
  } catch (e) {
    // Node 在超限时抛 ERR_BUFFER_TOO_LARGE。这里换成自己的错误码，
    // 让上层能给出「文件有压缩炸弹」这种可操作的提示，而不是一句 zlib 内部话。
    if (e && (e.code === 'ERR_BUFFER_TOO_LARGE' || /maxOutputLength|larger than/i.test(String(e?.message ?? '')))) {
      throw new Error('ZIP_ENTRY_TOO_LARGE:' + name)
    }
    throw e
  }
}

/**
 * 把上限选项规整成一个正整数。
 *
 * ⚠️ 不能直接用 `Number(v)`：传进来 NaN / 0 / 负数时，比较式 `size > NaN`
 *    恒为 false —— **上限会彻底失效**，正好是这里最不能出错的地方。
 *
 * @param {any} v
 * @param {number} def 兜底值
 * @returns {number}
 */
function positiveInt(v, def) {
  const n = Number(v)
  return Number.isFinite(n) && n >= 1 ? Math.floor(n) : def
}

/**
 * 解析 ZIP 中央目录，按需解压条目。
 *
 * 两条安全约束（都不是「可选的加固」，是**不做就会打死宿主进程**）：
 *   1. **只解需要的条目**（`options.only`）。解析只需要 document.xml 与 styles.xml；
 *      其余条目（图片、嵌入对象）连解压都不做——不分配、不占 CPU。
 *      注意这不影响浏览器端渲染：渲染用的是**原始 docx 字节**（走 getDocxBytes），
 *      走的是另一条路。
 *   2. **解压有上限**（单条目 + 总预算）。挡住压缩炸弹。
 *
 * 中央目录本身也来自不可信输入，所以每一步都做越界检查——恶意构造的偏移量
 * 以前会变成 RangeError 之类的天书错误，现在统一翻成 ZIP_BAD_CDH / ZIP_TRUNCATED。
 *
 * @param {Buffer} buf zip 文件字节
 * @param {{only?: string[]|Set<string>, maxOutputLength?: number, maxTotalBytes?: number}} [options]
 * @returns {Map<string, Buffer>} 条目名 → 内容（只含被解出来的）
 */
export function readZipBuffer(buf, options = {}) {
  const only = options.only ? new Set(options.only) : null
  const maxEntryBytes = positiveInt(options.maxOutputLength, MAX_ENTRY_BYTES)
  const maxTotalBytes = positiveInt(options.maxTotalBytes, MAX_TOTAL_BYTES)

  // 从尾部找 EOCD（0x06054b50），允许尾部有 0-65535 字节注释
  let eocd = -1
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 22 - 65536; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) throw new Error('ZIP_EOCD_NOT_FOUND')
  const count = buf.readUInt16LE(eocd + 10)
  let off = buf.readUInt32LE(eocd + 16)
  const entries = new Map()
  let total = 0
  for (let n = 0; n < count; n++) {
    // 中央目录记录至少 46 字节；越界说明目录被截断或伪造
    if (off + 46 > buf.length) throw new Error('ZIP_BAD_CDH')
    if (buf.readUInt32LE(off) !== 0x02014b50) throw new Error('ZIP_BAD_CDH')
    const method = buf.readUInt16LE(off + 10)
    const csize = buf.readUInt32LE(off + 20)
    const nameLen = buf.readUInt16LE(off + 28)
    const extraLen = buf.readUInt16LE(off + 30)
    const cmtLen = buf.readUInt16LE(off + 32)
    const lho = buf.readUInt32LE(off + 42)
    const name = buf.toString('utf8', off + 46, off + 46 + nameLen)
    const next = off + 46 + nameLen + extraLen + cmtLen
    if (next > buf.length) throw new Error('ZIP_BAD_CDH')
    off = next

    // ⚠️ 只需要这两个条目。其余的**不解压**：一张几百 KB 的图片能吃掉几十 MB 内存，
    //    而解析根本用不到它。不在 only 白名单里的条目直接跳过（连 inflate 都不做）。
    if (only && !only.has(name)) continue

    if (lho + 30 > buf.length) throw new Error('ZIP_BAD_LFH')
    const lnameLen = buf.readUInt16LE(lho + 26)
    const lextraLen = buf.readUInt16LE(lho + 28)
    const dataStart = lho + 30 + lnameLen + lextraLen
    if (dataStart + csize > buf.length) throw new Error('ZIP_TRUNCATED')
    const raw = buf.subarray(dataStart, dataStart + csize)

    let content
    if (method === 0) {
      // 存储法（不压缩）：大小就是 csize，同样要过上限
      if (csize > maxEntryBytes) throw new Error('ZIP_ENTRY_TOO_LARGE:' + name)
      content = raw
    } else if (method === 8) {
      content = inflateEntry(raw, name, maxEntryBytes)
    } else {
      throw new Error('ZIP_UNSUPPORTED_METHOD_' + method)
    }

    total += content.length
    if (total > maxTotalBytes) throw new Error('ZIP_TOTAL_TOO_LARGE')
    entries.set(name, content)
  }
  return entries
}

/* ------------------------------------------------------------------ */
/* 二、XML 实体解码                                                     */
/* ------------------------------------------------------------------ */

const NAMED = {
  lt: '<', gt: '>', quot: '"', apos: "'", amp: '&',
  nbsp: '\u00a0', copy: '©', reg: '®', trade: '™',
  hellip: '…', mdash: '—', ndash: '–',
  lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”',
}

/**
 * 反转义 XML 实体。
 *
 * ⚠️ `&amp;` 必须**最后**处理，否则 `&amp;lt;`（字面量 `&lt;`）会被二次解码。
 * 这个顺序坑在投标工作台里已经踩过一次，这里沿用同样的实现。
 *
 * @param {string} s
 * @returns {string}
 */
export function decodeXml(s) {
  if (!s || s.indexOf('&') < 0) return s || ''
  return s.replace(/&(#[xX][0-9a-fA-F]+|#\d+|[a-zA-Z][a-zA-Z0-9]*);/g, (m, g) => {
    if (g[0] === '#') {
      const hex = g[1] === 'x' || g[1] === 'X'
      const code = hex ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10)
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return m
      try { return String.fromCodePoint(code) } catch { return m }
    }
    const v = NAMED[g]
    return v === undefined ? m : v
  })
}

/* ------------------------------------------------------------------ */
/* 三、标签扫描                                                         */
/* ------------------------------------------------------------------ */

/**
 * 找标签的结束 `>`，跳过引号内的内容。
 *
 * XML 允许属性值里出现裸 `>`，OOXML 的 VML `style` 属性尤其常见
 * （`style="…width:49.85pt;…"`）。直接 indexOf('>') 会截在属性中间，
 * 整段 XML 从此错位。
 *
 * @param {string} xml
 * @param {number} from ` < ` 之后的位置
 * @returns {number} `>` 的下标；找不到返回 -1
 */
function tagEnd(xml, from) {
  let i = from
  let quote = ''
  while (i < xml.length) {
    const c = xml[i]
    if (quote) {
      if (c === quote) quote = ''
    } else if (c === '"' || c === "'") {
      quote = c
    } else if (c === '>') {
      return i
    }
    i++
  }
  return -1
}

/**
 * 从标签原文里取属性值（支持带命名空间前缀的属性名）。
 *
 * @param {string} raw 标签内部文本（不含 < >）
 * @param {string} name 属性名，如 `w:val`
 * @returns {string} 属性值；缺省返回空串
 */
export function attr(raw, name) {
  if (!raw) return ''
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const m = new RegExp('\\s' + escaped + '\\s*=\\s*"([^"]*)"').exec(raw)
  return m ? decodeXml(m[1]) : ''
}

/**
 * 去命名空间前缀：`w:p` → `p`。
 * @param {string} name
 * @returns {string}
 */
export function localOf(name) {
  const s = String(name ?? '')
  const i = s.indexOf(':')
  return i < 0 ? s : s.slice(i + 1)
}

/**
 * 单遍扫描 XML，产出 start / end / text 三类事件。
 *
 * @param {string} xml
 * @returns {Generator<{t:'start'|'end'|'text', name?:string, local?:string, self?:boolean, raw?:string, s?:string}>}
 */
export function* scanXml(xml) {
  const n = xml.length
  let i = 0
  while (i < n) {
    const lt = xml.indexOf('<', i)
    if (lt < 0) {
      if (i < n) yield { t: 'text', s: xml.slice(i) }
      break
    }
    if (lt > i) yield { t: 'text', s: xml.slice(i, lt) }

    // 注释
    if (xml.startsWith('<!--', lt)) {
      const e = xml.indexOf('-->', lt)
      i = e < 0 ? n : e + 3
      continue
    }
    // CDATA：内容当文本
    if (xml.startsWith('<![CDATA[', lt)) {
      const e = xml.indexOf(']]>', lt)
      yield { t: 'text', s: xml.slice(lt + 9, e < 0 ? n : e) }
      i = e < 0 ? n : e + 3
      continue
    }
    // 处理指令
    if (xml.startsWith('<?', lt)) {
      const e = xml.indexOf('?>', lt)
      i = e < 0 ? n : e + 2
      continue
    }
    // DOCTYPE
    if (xml.startsWith('<!', lt)) {
      const e = tagEnd(xml, lt + 2)
      i = e < 0 ? n : e + 1
      continue
    }

    const e = tagEnd(xml, lt + 1)
    if (e < 0) break
    const inner = xml.slice(lt + 1, e)
    i = e + 1

    if (inner[0] === '/') {
      const name = inner.slice(1).trim()
      yield { t: 'end', name, local: localOf(name) }
      continue
    }
    const self = inner.endsWith('/')
    const body = self ? inner.slice(0, -1) : inner
    const sp = body.search(/[\s]/)
    const name = sp < 0 ? body : body.slice(0, sp)
    yield { t: 'start', name, local: localOf(name), self, raw: body }
  }
}
