/**
 * 评判 + 锚点：测试
 *
 * 覆盖「AI 读若干段 → 给出评判 → 按评判结果打锚点」这条链路。
 *
 * 与 summary.mjs 的分工：那边测「总结」（要点是什么），这边测「评判」
 * （这句话有风险吗、多严重、该怎么改）。两者用的是同一组锚点机制，
 * 但评判多出 severity / category / suggestion 三个结构化字段——
 * 它们决定了面板能不能按「废标 / 高风险 / 需澄清」分类着色。
 */

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { makeDocx, stylesXml, para } from './fixtures.mjs'
import { DocAnchorViewer, buildDocTools, SEVERITIES, CATEGORIES } from '../lib/index.js'

/** 临时数据目录。 */
async function withTemp(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'dav-judge-'))
  const prev = process.env.DOC_ANCHOR_VIEWER_DIR
  process.env.DOC_ANCHOR_VIEWER_DIR = dir
  try { await fn(dir) } finally {
    if (prev === undefined) delete process.env.DOC_ANCHOR_VIEWER_DIR
    else process.env.DOC_ANCHOR_VIEWER_DIR = prev
    rmSync(dir, { recursive: true, force: true })
  }
}

/** 一份有废标条款、资格条款、评分办法的招标文件样本。 */
function tenderDocx(dir) {
  const p = join(dir, 'tender.docx')
  writeFileSync(p, makeDocx([
    para('第三章 投标人资格要求', 'H1'),
    para('3.1 投标人须具备独立法人资格，注册资金不低于 500 万元。'),
    para('3.2 投标人须提供近三年内至少两个同类项目业绩。'),
    para('3.3 本项目不接受联合体投标。'),
    para('第四章 评标办法', 'H1'),
    para('4.1 技术方案完整性：优 10 分，良 7 分，一般 4 分。'),
    para('4.2 报价得分按最低价法计算。'),
  ], stylesXml([{ id: 'H1', name: 'heading 1', outlineLvl: 0 }])))
  return p
}

/* ------------------------------------------------------------------ */
/* J：评判字段                                                         */
/* ------------------------------------------------------------------ */

test('J1 推荐值是稳定约定，但**不是白名单**', async () => {
  // 推荐值只是给调用方的默认词汇（面板自带中文标签），改动它们要谨慎；
  // 但任意自定义词都必须能用——这个插件是通用的，不该被招投标词汇绑死。
  assert.deepEqual(SEVERITIES, ['fatal', 'high', 'medium', 'low', 'info'])
  assert.deepEqual(CATEGORIES, ['invalid', 'qualification', 'scoring', 'technical', 'commercial', 'contract', 'other'])

  // 语义自证：导出的名字是「推荐」而不是「允许」
  const mod = await import('../lib/index.js')
  assert.ok(mod.SUGGESTED_SEVERITIES, '应该同时导出 SUGGESTED_SEVERITIES')
  assert.ok(mod.SUGGESTED_CATEGORIES, '应该同时导出 SUGGESTED_CATEGORIES')
})

test('J2 addMark 带评判：严重程度 / 类别 / 建议都能存下并回读', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })

    const m = await svc.addMark({
      docId: r.doc.id,
      text: '3.3 本项目不接受联合体投标。',
      note: '若以联合体形式投标将直接废标',
      severity: 'fatal',
      category: 'invalid',
      suggestion: '确认投标主体为单一法人，联合体成员不得参与',
    })
    assert.equal(m.ok, true, JSON.stringify(m))
    assert.equal(m.mark.severity, 'fatal')
    assert.equal(m.mark.category, 'invalid')
    assert.equal(m.mark.suggestion, '确认投标主体为单一法人，联合体成员不得参与')

    const l = await svc.listMarks({ docId: r.doc.id })
    assert.equal(l.marks[0].severity, 'fatal', '回读必须保留评判字段')
    assert.equal(l.marks[0].category, 'invalid')
  })
})

test('J3 自定义标签原样保留（不是白名单）', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })

    // 换一套完全不同的词汇：这正是「通用工具」要的
    const m = await svc.addMark({
      docId: r.doc.id,
      text: '3.2 投标人须提供近三年内至少两个同类项目业绩。',
      severity: 'P0',
      category: '风险点',
    })
    assert.equal(m.ok, true)
    assert.equal(m.mark.severity, 'P0', '自定义 severity 必须原样保留')
    assert.equal(m.mark.category, '风险点', '自定义 category 必须原样保留')

    // 空值与纯空白仍然归一成空（表示未设置）
    const e = await svc.addMark({ docId: r.doc.id, text: '3.3 本项目不接受联合体投标。', severity: '   ' })
    assert.equal(e.mark.severity, '', '空白归成空串')
  })
})

test('J3b 标签有长度上限，且控制字符被清理（防止把界面撑爆）', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })

    const long = 'X'.repeat(200)
    const m = await svc.addMark({ docId: r.doc.id, text: '3.1 投标人须具备独立法人资格，注册资金不低于 500 万元。', severity: long })
    assert.equal(m.mark.severity.length, 40, '超长标签要截断到 40 字')

    const nl = await svc.addMark({ docId: r.doc.id, text: '第四章 评标办法', category: '第一行\n第二行\t制表' })
    assert.ok(!/[\n\t]/.test(nl.mark.category), '换行/制表要清掉（徽标是单行的）')
  })
})

test('J3c color 字段：合法值保留，非法值归一成空串', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })

    const ok = await svc.addMark({ docId: r.doc.id, text: '3.3 本项目不接受联合体投标。', severity: 'high', color: '#e0574b' })
    assert.equal(ok.mark.color, '#e0574b')

    const rgb = await svc.addMark({ docId: r.doc.id, text: '第四章 评标办法', color: 'rgb(255, 100, 0)' })
    assert.equal(rgb.mark.color, 'rgb(255, 100, 0)')

    const named = await svc.addMark({ docId: r.doc.id, text: '3.2 投标人须提供近三年内至少两个同类项目业绩。', color: 'crimson' })
    assert.equal(named.mark.color, 'crimson')

    // 非法：不能把任意字符串塞进 style
    const bad = await svc.addMark({ docId: r.doc.id, text: '3.1 投标人须具备独立法人资格，注册资金不低于 500 万元。', color: 'javascript:alert(1)' })
    assert.equal(bad.mark.color, '', '非法 color 要归空，退回自动配色')
  })
})

test('J4 不传评判字段：普通标记照常工作（不必强制 AI 评判）', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })
    const m = await svc.addMark({ docId: r.doc.id, text: '第四章 评标办法', note: '记一下' })
    assert.equal(m.ok, true)
    assert.equal(m.mark.severity, '')
    assert.equal(m.mark.category, '')
    assert.equal(m.mark.suggestion, '')
  })
})

/* ------------------------------------------------------------------ */
/* K：按段落评判的完整链路                                             */
/* ------------------------------------------------------------------ */

test('K1 「读某几段 → 逐段评判 → 批量落锚」全链路', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })
    const docId = r.doc.id

    // 1) 读资格要求那一段区间（模型实际就是这么读的：锚点 + 条数）
    //    before: 0 是刻意的——要判的就是这三条，不想把上一节的标题也读进来
    const sec = await svc.readSection({ docId, index: 1, before: 0, after: 2 })
    assert.equal(sec.ok, true)
    assert.equal(sec.paragraphs.length, 3, '3.1 / 3.2 / 3.3 三段')
    assert.deepEqual(sec.paragraphs.map((p) => p.index), [1, 2, 3])

    // 2) 逐段给出评判（模拟模型产出）
    const verdicts = [
      { text: sec.paragraphs[0].text, severity: 'high', category: 'qualification', note: '注册资金门槛 500 万，需核对我司是否达标', suggestion: '核验营业执照与实缴资本' },
      { text: sec.paragraphs[1].text, severity: 'medium', category: 'qualification', note: '要求近三年两个同类业绩', suggestion: '整理可引用的合同清单' },
      { text: sec.paragraphs[2].text, severity: 'fatal', category: 'invalid', note: '不接受联合体投标', suggestion: '不得以联合体形式参与' },
    ]

    // 3) 一次批量落锚，带上原始序号（最精确）
    const res = await svc.setMarks({
      docId,
      marks: verdicts.map((v, i) => ({ ...v, index: sec.paragraphs[i].index })),
    })
    assert.equal(res.ok, true, JSON.stringify(res))
    assert.equal(res.added, 3)

    // 4) 每条都必须落在**它自己那一段**上
    assert.deepEqual(res.marks.map((m) => m.anchor.index), [1, 2, 3])
    assert.deepEqual(res.marks.map((m) => m.severity), ['high', 'medium', 'fatal'])
    assert.deepEqual(res.marks.map((m) => m.category), ['qualification', 'qualification', 'invalid'])
  })
})

test('K2 只给文字也能评判：Host 按精确匹配找段，锚点不丢', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })

    const res = await svc.setMarks({
      docId: r.doc.id,
      marks: [
        { text: '4.2 报价得分按最低价法计算。', severity: 'medium', category: 'scoring', note: '最低价法意味着价格战' },
        { text: '3.3 本项目不接受联合体投标。', severity: 'fatal', category: 'invalid' },
      ],
    })
    assert.equal(res.ok, true, JSON.stringify(res))
    assert.deepEqual(res.marks.map((m) => m.anchor.index), [6, 3], '按文档顺序匹配到各自那一段')
    assert.deepEqual(res.marks.map((m) => m.severity), ['medium', 'fatal'])
  })
})

test('K3 评判报告：按严重程度聚合，行为符合评标复核习惯', async () => {
  await withTemp(async (dir) => {
    const svc = new DocAnchorViewer({}, {})
    const r = await svc.openDoc({ filePath: tenderDocx(dir) })
    const docId = r.doc.id

    await svc.setMarks({
      docId,
      marks: [
        { text: '3.1 投标人须具备独立法人资格，注册资金不低于 500 万元。', severity: 'high', category: 'qualification' },
        { text: '3.2 投标人须提供近三年内至少两个同类项目业绩。', severity: 'medium', category: 'qualification' },
        { text: '3.3 本项目不接受联合体投标。', severity: 'fatal', category: 'invalid' },
        { text: '4.2 报价得分按最低价法计算。', severity: 'low', category: 'scoring' },
      ],
    })

    const l = await svc.listMarks({ docId })
    // 按严重程度排序：废标在前，参考在后
    const rank = { fatal: 0, high: 1, medium: 2, low: 3, info: 4, '': 5 }
    const sorted = [...l.marks].sort((a, b) => (rank[a.severity] ?? 9) - (rank[b.severity] ?? 9))
    assert.equal(sorted[0].severity, 'fatal')
    assert.equal(sorted[0].anchor.index, 3)
    assert.deepEqual(sorted.map((m) => m.severity), ['fatal', 'high', 'medium', 'low'])
  })
})

/* ------------------------------------------------------------------ */
/* T：工具面                                                           */
/* ------------------------------------------------------------------ */

test('T2 工具面上能拿到评判参数，纯文本输出带严重程度标签', async () => {
  await withTemp(async (dir) => {
    const reg = []
    buildDocTools(
      { get: (k) => (k === 'tools' ? { register: (d) => { reg.push(d); return () => {} } } : undefined) },
      new DocAnchorViewer({}, {}),
    )
    const def = reg[0]
    const show = (v) => def.output.render({}, v).map((b) => b.text).join('\n')

    // 参数表里必须有这几个字段，否则模型根本不知道能评判
    const props = def.parameters.properties
    assert.ok(props.severity, '要有 severity')
    assert.ok(props.category, '要有 category')
    assert.ok(props.suggestion, '要有 suggestion')
    assert.ok(props.color, '要有 color')
    // 关键：**不能再有 enum** —— 分类是自由文本，枚举会把插件锁死在招投标词汇上
    assert.equal(props.severity.enum, undefined, 'severity 不能是枚举')
    assert.equal(props.category.enum, undefined, 'category 不能是枚举')
    assert.equal(props.marks.items.properties.severity.enum, undefined, '批量里也不能是枚举')
    assert.equal(props.marks.items.properties.category.enum, undefined, '批量里也不能是枚举')

    const open = await def.execute({ action: 'open', filePath: tenderDocx(dir) })
    const docId = open.doc.id

    const res = await def.execute({
      action: 'set_marks', docId,
      marks: [
        { text: '3.3 本项目不接受联合体投标。', severity: 'fatal', category: 'invalid', note: '不接受联合体', suggestion: '确认单一主体投标' },
        { text: '4.2 报价得分按最低价法计算。', severity: 'medium', category: 'scoring', note: '价格竞争激烈' },
      ],
    })
    assert.equal(res.ok, true, JSON.stringify(res))

    const out = show(res)
    // 纯文本输出原样显示标签（不再翻译成中文），格式是 [severity/category]
    assert.ok(out.includes('[fatal/invalid]'), '输出要带标签：' + out.slice(0, 200))
    assert.ok(out.includes('→') || out.includes('确认单一主体投标'), '输出要带建议：' + out.slice(0, 200))

    const listOut = show(await def.execute({ action: 'list_marks', docId }))
    assert.ok(listOut.includes('fatal'), '列表里要能看出严重程度：' + listOut.slice(0, 200))
  })
})

test('T3 单条 add_mark 也支持评判字段', async () => {
  await withTemp(async (dir) => {
    const reg = []
    buildDocTools(
      { get: (k) => (k === 'tools' ? { register: (d) => { reg.push(d); return () => {} } } : undefined) },
      new DocAnchorViewer({}, {}),
    )
    const def = reg[0]
    const show = (v) => def.output.render({}, v).map((b) => b.text).join('\n')
    const open = await def.execute({ action: 'open', filePath: tenderDocx(dir) })

    const res = await def.execute({
      action: 'add_mark',
      docId: open.doc.id,
      text: '3.3 本项目不接受联合体投标。',
      severity: 'fatal',
      category: 'invalid',
      note: '会废标',
      suggestion: '不要组联合体',
    })
    assert.equal(res.ok, true)
    assert.equal(res.mark.severity, 'fatal')
    const out = show(res)
    assert.ok(out.includes('fatal'), '单条输出也要带严重程度：' + out)
  })
})
