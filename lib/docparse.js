/**
 * 文档锚点视图 · DOCX 结构化解析
 *
 * 只做预览需要的那一层：把 document.xml 拆成**按文档顺序**的段落序列，
 * 并解析 styles.xml，判断每段是不是标题、几级标题。
 *
 * 为什么自己解析、不用现成的 docx 解析库：
 *   1. 预览渲染已经交给 docx-preview（浏览器端），但**锚点必须由服务端
 *      按同一份 XML 算出来**——否则「Agent 看到的第 37 段」和「浏览器渲染的
 *      第 37 段」会各说各话。两边都读 document.xml 才可能对得上。
 *   2. 只需要段落文本与标题层级，用不上完整的样式/表格/图片模型；
 *      单遍扫描对 3.8MB 的 document.xml 是毫秒级，DOM 方案要 37MB 内存。
 *
 * 与 docx-preview 的对齐策略：**不用段落序号做主键**（两个渲染器对空段、
 * 域代码、文本框的处理不可能保证一致），而是用「归一化的段落全文 + 出现序号」。
 * 客户端拿这个去渲染结果里精确匹配，匹配不上就如实报「未定位到」。
 */

import { scanXml, attr, decodeXml } from './ooxml.js'

/** 归一化：去掉所有空白，用于跨渲染器的文本匹配。 */
export function normText(s) {
  return String(s ?? '').replace(/\s+/g, '').trim()
}

/**
 * 解析 styles.xml，得到 styleId → { name, outlineLvl, basedOn }。
 *
 * 真实招标文件里 9105 处 pStyle 引用，标题层级定义全在 styles 里，
 * 段落上什么都没有——不解析 styles 就一个标题都认不出来。
 *
 * @param {string} xml styles.xml 文本
 * @returns {Map<string, {name:string, outlineLvl:number|null, basedOn:string}>}
 */
export function parseStyles(xml) {
  const styles = new Map()
  if (!xml) return styles

  let cur = null
  let depth = 0
  for (const ev of scanXml(xml)) {
    if (ev.t === 'start' && ev.local === 'style') {
      const id = attr(ev.raw, 'w:styleId')
      cur = { id, type: attr(ev.raw, 'w:type'), name: '', outlineLvl: null, basedOn: '' }
      depth = 0
      if (ev.self) {
        if (cur.id) styles.set(cur.id, cur)
        cur = null
      }
      continue
    }
    if (!cur) continue
    if (ev.t === 'end' && ev.local === 'style') {
      if (cur.id && cur.type === 'paragraph') styles.set(cur.id, cur)
      cur = null
      continue
    }
    if (ev.t === 'start' && ev.local === 'name') {
      cur.name = attr(ev.raw, 'w:val')
      continue
    }
    if (ev.t === 'start' && ev.local === 'basedOn') {
      cur.basedOn = attr(ev.raw, 'w:val')
      continue
    }
    if (ev.t === 'start' && ev.local === 'outlineLvl') {
      const v = Number(attr(ev.raw, 'w:val'))
      if (Number.isFinite(v)) cur.outlineLvl = v
      continue
    }
    depth++ // 只为可读性保留，不做语义判断
    void depth
  }
  return styles
}

/**
 * 沿 basedOn 链回溯标题级别。
 *
 * Word 里「标题 1」常常 basedOn 到一个自定义样式，outlineLvl 只写在祖先上；
 * 不回溯就会漏掉整棵标题树。
 *
 * @param {string} styleId
 * @param {Map<string, any>} styles
 * @returns {{outlineLvl:number|null, name:string, chain:string[]}}
 */
export function resolveStyle(styleId, styles) {
  const chain = []
  let cur = styleId
  let outlineLvl = null
  let name = ''
  let guard = 0
  while (cur && styles.has(cur) && guard++ < 20) {
    const s = styles.get(cur)
    chain.push(cur)
    if (!name && s.name) name = s.name
    if (outlineLvl === null && s.outlineLvl !== null && s.outlineLvl !== undefined) outlineLvl = s.outlineLvl
    if (outlineLvl !== null) {
      // 继续回溯只为拿 name，不再覆盖层级
    }
    cur = s.basedOn
  }
  return { outlineLvl, name, chain }
}

/** 内置标题样式名（中英文 Word 默认名）。 */
const HEADING_NAME = /^(heading\s*[1-9]|标题\s*[1-9]|标题[一二三四五六七八九十]|title|标题)$/i

/**
 * 判定一段是不是标题，并给出 1 起的级别。
 *
 * 两条判据，任一成立即可：outlineLvl（0 起）或样式名形如 Heading N / 标题 N。
 * 两者都拿不到时按正文处理——**不猜**。
 *
 * @param {{outlineLvl:number|null, name:string, styleId:string}} st
 * @returns {{heading:boolean, level:number}}
 */
export function headingLevel(st) {
  const lvl = st.outlineLvl
  if (lvl !== null && lvl !== undefined && lvl >= 0 && lvl <= 8) {
    return { heading: true, level: lvl + 1 }
  }
  const name = String(st.name || '')
  const id = String(st.styleId || '')
  const m = /^(?:heading|标题)\s*([1-9])/i.exec(name) || /^heading\s*([1-9])$/i.exec(id)
  if (m) return { heading: true, level: Number(m[1]) }
  if (HEADING_NAME.test(name.trim())) return { heading: true, level: 1 }
  return { heading: false, level: 0 }
}

/**
 * 把收口后的段落帧转成最终记录（样式解析 + 标题判定）。
 * @param {{text:string[], styleId:string, outlineLvl:number|null}} frame
 * @param {Map<string,any>} styles
 * @param {number} index 段落序号
 * @returns {object}
 */
function emit(frame, styles, index) {
  const text = frame.text.join('')
  const st = resolveStyle(frame.styleId, styles)
  const lvl = frame.outlineLvl !== null ? frame.outlineLvl : st.outlineLvl
  const hl = headingLevel({ outlineLvl: lvl, name: st.name, styleId: frame.styleId })
  return {
    index,
    text,
    norm: normText(text),
    heading: hl.heading,
    level: hl.level,
    styleId: frame.styleId,
    styleName: st.name,
  }
}

/**
 * 解析 document.xml，产出按文档顺序的段落列表。
 *
 * 每段记录：序号、文本、是否标题、级别、样式 id/名。
 * 文本只收 `w:t`，并跳过域代码（instrText）与修订删除（delText）——
 * 这两类在 Word 里要么不显示、要么显示成别的东西，收进来会让锚点对不上。
 *
 * @param {string} xml document.xml 文本
 * @param {Map<string, any>} styles parseStyles 的产物
 * @returns {Array<{index:number, text:string, norm:string, heading:boolean, level:number, styleId:string, styleName:string}>}
 */
export function parseParagraphs(xml, styles = new Map()) {
  const out = []
  /** 段落栈：文本框里的 w:p 会嵌在外层 w:p 里。 */
  const stack = []
  let textDepth = 0

  for (const ev of scanXml(xml)) {
    if (ev.t === 'start') {
      if (ev.local === 'p') {
        const frame = { text: [], styleId: '', outlineLvl: null, outer: stack.length > 0 }
        // ⚠️ 自闭合 <w:p/> 是**空段落**，必须当场收口。
        //    压进栈里而没有对应的 </w:p> 的话，它之后的所有内容都会被并进
        //    栈顶这一"段"——整篇文档塌成几段、大纲全空（实测 9189 个 w:p
        //    里有 8 个自闭合，就足以毁掉整篇解析）。
        if (ev.self) {
          if (frame.outer) stack[stack.length - 1].text.push('')
          else out.push(emit(frame, styles, out.length))
          continue
        }
        stack.push(frame)
        continue
      }
      if (ev.local === 'pStyle' && stack.length) {
        stack[stack.length - 1].styleId = attr(ev.raw, 'w:val')
        continue
      }
      if (ev.local === 'outlineLvl' && stack.length) {
        const v = Number(attr(ev.raw, 'w:val'))
        if (Number.isFinite(v)) stack[stack.length - 1].outlineLvl = v
        continue
      }
      if (ev.local === 't') { textDepth++; continue }
      if (ev.local === 'tab' && stack.length && textDepth > 0) {
        stack[stack.length - 1].text.push(' ')
        continue
      }
      continue
    }

    if (ev.t === 'end') {
      if (ev.local === 't') { if (textDepth > 0) textDepth--; continue }
      if (ev.local === 'p' && stack.length) {
        const frame = stack.pop()
        const text = frame.text.join('')
        // 嵌套段落（文本框）不单独成段，文本折进外层——与「一个 w:p 一块」的口径一致，
        // 避免同一段文字在列表里出现两次。
        if (frame.outer) { stack[stack.length - 1].text.push(text); continue }
        out.push(emit(frame, styles, out.length))
        continue
      }
      continue
    }

    // 文本
    if (textDepth > 0 && stack.length && ev.s) {
      const s = ev.s
      // 纯空白节点（xml:space 保留）对匹配无意义，归一化时会被去掉，这里照收即可
      stack[stack.length - 1].text.push(decodeXml(s))
    }
  }

  // 文档不闭合（截断的 XML）时把栈里的段落补出来，不让用户白屏
  while (stack.length) {
    const frame = stack.pop()
    const text = frame.text.join('')
    if (frame.outer) { stack[stack.length - 1].text.push(text); continue }
    out.push(emit(frame, styles, out.length))
  }

  return out
}

/**
 * 给每段标上「同文第几次出现」（从 0 起），跨**全部段落**统计。
 *
 * ⚠️ 这个序号必须按全文统计、不能只在大纲里数：客户端是在**所有渲染出的
 * 段落**里按「归一化文本 + 第几次出现」找目标的。若这里只数标题，
 * 「1.1」这种既是标题又出现在正文里的文本就会全被定位到错误的一处。
 *
 * @param {Array<object>} paragraphs
 * @returns {Map<number, number>} 段落序号 → occurrence
 */
export function occurrenceMap(paragraphs) {
  const seen = new Map()
  const out = new Map()
  for (const p of paragraphs) {
    const k = p.norm || ('\u0000' + p.index) // 空段落各自独立，不参与同文计数
    const n = seen.get(k) ?? 0
    seen.set(k, n + 1)
    out.set(p.index, n)
  }
  return out
}

/**
 * 由段落列表生成大纲条目。
 *
 * 同名文本（如两个「附件一」）按出现顺序编号，客户端据此定位到**各自那一处**，
 * 而不是都跳到第一个。
 *
 * @param {Array<object>} paragraphs
 * @returns {Array<{index:number, title:string, text:string, occurrence:number, level:number}>}
 */
export function buildOutline(paragraphs) {
  const occ = occurrenceMap(paragraphs)
  const out = []
  for (const p of paragraphs) {
    if (!p.heading) continue
    if (!p.norm) continue
    out.push({
      index: p.index,
      title: String(p.text ?? '').trim(),
      text: p.text,
      norm: p.norm,
      occurrence: occ.get(p.index) ?? 0,
      level: p.level,
    })
  }
  return out
}
