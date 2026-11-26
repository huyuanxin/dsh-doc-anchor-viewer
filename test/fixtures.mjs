/**
 * 测试用：造一个最小的 docx（ZIP + OOXML），不依赖任何第三方库。
 *
 * 为什么要自己造而不是塞一个二进制样本到仓库：
 *   测试要断言的是「标题层级识别」「同文多处锚点」「样式继承」这些**结构**行为，
 *   只有自己造才能精确控制每一段的样式引用。二进制样本既看不出在测什么，
 *   也没法针对边界改。
 */

import { deflateRawSync } from 'node:zlib'

/** CRC32（ZIP 需要）。 */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256)
  for (let i = 0; i < 256; i++) {
    let c = i
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[i] = c >>> 0
  }
  return t
})()

function crc32(buf) {
  let c = 0xffffffff
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

/**
 * 造一个 ZIP（只用 deflate，够 docx 用）。
 * @param {Array<{name:string, data:Buffer|string}>} entries
 * @returns {Buffer}
 */
export function makeZip(entries) {
  const locals = []
  const centrals = []
  let offset = 0

  for (const e of entries) {
    const nameBuf = Buffer.from(e.name, 'utf8')
    const raw = Buffer.isBuffer(e.data) ? e.data : Buffer.from(e.data, 'utf8')
    const comp = deflateRawSync(raw)
    const crc = crc32(raw)

    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)      // version needed
    local.writeUInt16LE(0, 6)       // flags
    local.writeUInt16LE(8, 8)       // method = deflate
    local.writeUInt16LE(0, 10)      // time
    local.writeUInt16LE(0, 12)      // date
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(comp.length, 18)
    local.writeUInt32LE(raw.length, 22)
    local.writeUInt16LE(nameBuf.length, 26)
    local.writeUInt16LE(0, 28)

    locals.push(local, nameBuf, comp)

    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(20, 4)    // version made by
    central.writeUInt16LE(20, 6)    // version needed
    central.writeUInt16LE(0, 8)
    central.writeUInt16LE(8, 10)
    central.writeUInt16LE(0, 12)
    central.writeUInt16LE(0, 14)
    central.writeUInt32LE(crc, 16)
    central.writeUInt32LE(comp.length, 20)
    central.writeUInt32LE(raw.length, 24)
    central.writeUInt16LE(nameBuf.length, 28)
    central.writeUInt16LE(0, 30)    // extra len
    central.writeUInt16LE(0, 32)    // comment len
    central.writeUInt16LE(0, 34)    // disk
    central.writeUInt16LE(0, 36)    // internal attrs
    central.writeUInt32LE(0, 38)    // external attrs
    central.writeUInt32LE(offset, 42)

    centrals.push(central, nameBuf)
    offset += local.length + nameBuf.length + comp.length
  }

  const cdBuf = Buffer.concat(centrals)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(0x06054b50, 0)
  eocd.writeUInt16LE(0, 4)
  eocd.writeUInt16LE(0, 6)
  eocd.writeUInt16LE(entries.length, 8)
  eocd.writeUInt16LE(entries.length, 10)
  eocd.writeUInt32LE(cdBuf.length, 12)
  eocd.writeUInt32LE(offset, 16)
  eocd.writeUInt16LE(0, 20)

  return Buffer.concat([...locals, cdBuf, eocd])
}

/**
 * 造一个样式表。
 * @param {Array<{id:string, name?:string, outlineLvl?:number, basedOn?:string}>} styles
 * @returns {string}
 */
export function stylesXml(styles) {
  const body = styles.map((s) => {
    const parts = [`<w:style w:type="paragraph" w:styleId="${s.id}">`]
    if (s.name) parts.push(`<w:name w:val="${s.name}"/>`)
    if (s.basedOn) parts.push(`<w:basedOn w:val="${s.basedOn}"/>`)
    if (s.outlineLvl !== undefined) parts.push(`<w:pPr><w:outlineLvl w:val="${s.outlineLvl}"/></w:pPr>`)
    parts.push('</w:style>')
    return parts.join('')
  }).join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + body + '</w:styles>'
}

/**
 * 造一段。
 * @param {string} text
 * @param {string} [styleId]
 * @returns {string}
 */
export function para(text, styleId) {
  const pPr = styleId ? `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>` : ''
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
}

/** XML 转义。 */
export function escapeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/**
 * 造一个完整 docx：正文段落 + 可选样式表。
 * @param {string[]} paragraphs 每项是已构造好的 <w:p> 串
 * @param {string} [styles]
 * @returns {Buffer}
 */
export function makeDocx(paragraphs, styles) {
  const doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:body>' + paragraphs.join('') + '</w:body></w:document>'
  const entries = [
    { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>' },
    { name: 'word/document.xml', data: doc },
    { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
  ]
  if (styles) entries.push({ name: 'word/styles.xml', data: styles })
  return makeZip(entries)
}
