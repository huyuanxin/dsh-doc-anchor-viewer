# 文档锚点视图 · dsh-doc-anchor-viewer

[![npm](https://img.shields.io/npm/v/dsh-doc-anchor-viewer.svg)](https://www.npmjs.com/package/dsh-doc-anchor-viewer)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![dsh-plugin](https://img.shields.io/badge/topic-dsh--plugin-2ea44f.svg)](https://github.com/topics/dsh-plugin)
[![Listed on dsh-plugin.org](https://dsh-plugin.org/badges/listed.svg)](https://dsh-plugin.org/plugins/huyuanxin/dsh-doc-anchor-viewer)

把 `.docx` 按 **Word 自己的版式**渲染出来，并把「大纲」与「自定义标记」做成可跳转的锚点：点左边、右边滚到位并闪一下高亮。Agent 也能用同一套锚点读原文，不必把整篇文档塞进上下文。

> Render a `.docx` with Word's own layout and turn its outline plus your own marks into jump targets — clickable for you in a panel, and readable by the agent through the `doc_viewer` tool.

**这是一个通用工具，不绑定任何行业。** 标记的「严重程度」与「分类」都是**自由文本**：
合同审阅可以写「高风险 / 付款条款」，病历批注可以写「危急值 / 检验结果」，代码评审可以写「P0 / 空指针」。
内置的 `fatal/high/medium/low/info` 只是**推荐词汇**（面板自带中文标签），不是白名单。

**目录**：[安装](#安装) · [30 秒上手](#30-秒上手) · [Agent 怎么用](#agent-怎么用doc_viewer) · [数据与边界](#数据与边界) · [权限与风险](#权限外部服务与风险请先读这一节) · [配置](#配置) · [兼容性](#兼容性) · [常见问题](#常见问题) · [开发](#开发) · [English](#english)

---

## 安装

一行装进 profile，装完刷新页面：

```bash
dsh plugin --profile web add dsh-doc-anchor-viewer
```

DSH Desktop 用的是 `desktop` 这个 profile 名，命令一样、只换名字：

```bash
dsh plugin --profile desktop add dsh-doc-anchor-viewer
```

也可以从源码装（仓库根就有 `package.json`，`dsh.bundle.patch` 会被自动识别并加进 profile 的 bundle 列表）：

```bash
dsh plugin --profile web add github:huyuanxin/dsh-doc-anchor-viewer
```

不想敲命令也行：在 DSH 的插件面板里添加插件，包名填 `dsh-doc-anchor-viewer`。

装完需要**刷新一次页面**（或重启 DSH）让浏览器加载面板代码。之后左侧栏会出现 **「文档锚点」** 入口。

卸载：

```bash
dsh plugin --profile web remove dsh-doc-anchor-viewer
```

> NPM：<https://www.npmjs.com/package/dsh-doc-anchor-viewer>

## 30 秒上手

1. 左侧栏点 **「文档锚点」**；
2. 点「打开」→ 粘贴本地绝对路径（如 `D:\合同.docx`），或「从本机选择…」；
3. 左栏出现大纲，点一条 → 右侧 Word 文档滚到那一段并闪一下；
4. 滚到想记住的位置，点左下角「标记当前段落」，写一句备注，它会出现在「标记」页签里。

会话里也能用同一份数据：直接对助手说「把这篇合同的付款条款挑出来打上标记」，它会调用 `doc_viewer`。

## 它解决什么

| 没有它 | 有它 |
| --- | --- |
| 助手读一篇 300 页的 docx，要么读全文撑爆上下文，要么只能采样 | 先用 `get_digest` 拿一份按大纲切好的压缩底稿，再按锚点只读需要的那几段 |
| 助手说「第 3.2 条有问题」，你得自己 Ctrl+F 找半天 | 结论连同锚点一起写回，面板里点一条就跳到那一段 |
| 预览器只负责「看」，看出来的判断留在聊天记录里 | 判断变成带严重程度/分类/建议的结构化标记，和文档一起存着 |
| 不同行业的批注词汇互不兼容 | 严重程度与分类都是自由文本，一个机制服务所有场景 |

## 面板在哪

侧栏（左侧导航）里的 **「文档锚点」** 入口。打开后：

- 左栏：文档列表 → 大纲 / 标记两个页签
- 右栏：docx-preview 渲染的 Word 文档

## 两种打开方式

1. **在面板里打开**：左栏「打开」→ 粘贴本地绝对路径，或「从本机选择…」选一个文件（会复制到插件数据目录）。
2. **在会话里让助手打开**：`doc_viewer` 工具，`action: "open"`，`filePath` 给绝对路径。打开后面板里立刻能看到。

## 侧栏标签页（另一种打开方式）

同一份 docx 有两种看法，入口不同、用途不同：

| | 「文档锚点」全局面板 | 侧栏标签页 |
| --- | --- | --- |
| 位置 | 主区域（占满） | 右侧栏 |
| 布局 | 左右分栏 | **左栏固定 240px + 右文档** |
| 适合 | 专心读一篇、批量打标记 | 边看会话边看文件 |

侧栏标签页**会自动识别**这篇文档登记过没有：

- **登记过**（在面板或 `doc_viewer` 工具里打开过）→ 左边出「大纲 / 标记」竖栏，右边 Word 渲染，点击照常跳转与闪烁
- **没登记过** → 就是干净的 Word 预览，不摆一个空侧栏

它用 `priority: 'extension'` 认领 `*.docx`，压过产品自带的 office 预览——因为自带的走 LibreOffice 转 PDF，装不下大纲与锚点。

**布局为什么是左右而不是上下**：侧栏本来就矮，上下切一刀两边都剩不下几行；左右分栏让标记铺满一条竖列，文档也不被压扁。侧栏窄于 460px 时左栏自动收到 168px。

**已知限制**：只有**绝对路径**的 `.docx` 能关联到已登记的大纲/标记。地址是工作区相对路径时（`dsh-resource://file/session/<id>/docs/a.docx`），侧栏 body 拿不到会话的 workspace root，无法反推绝对路径，于是退化成普通预览——这是**如实降级**，不是 bug。

## Agent 怎么用：`doc_viewer`

一个工具 + `action` 参数，不是一个动作一个工具——这些动作共享同一组参数语义（`docId` / 锚点）与同一份失败含义，拆开只会让模型在每个工具上都学一遍。

| `action` | 作用 |
| --- | --- |
| `open` | 登记并预览一篇 docx（`filePath` 绝对路径），顺手返回大纲 |
| `list_docs` | 已登记文档 |
| `get_outline` | 大纲锚点（`index` / 层级 / 标题） |
| `get_digest` | **通读用**：按大纲切节、每节截断，每节都带好锚点三要素 |
| `read_section` | 按 `index` 或一段原文读回前后若干段 |
| `add_mark` | 加一条标记 |
| `set_marks` | **批量**写标记（一条总结产出的一组要点，一次落盘） |
| `list_marks` / `remove_mark` | 读写标记 |
| `remove_doc` | 移除登记（不删源文件） |
| `diagnostics` | 数据目录、能力自检 |

### 让会话总结文档并打上锚点

一条完整的、给 Agent 用的两步链路：

1. **`doc_viewer action=get_digest`** —— 取一份「能拿来总结」的压缩底稿。
   按大纲把文档切成若干节，每节只取开头那一小段正文（`maxCharsPerSection`，默认 600 字），
   整份底稿再按 `maxChars`（默认 12000 字）收敛。**每节都带好锚点三要素**
   （`anchorIndex` / `norm` / `occurrence`），模型不用自己猜序号、也不用复制原文。

2. **`doc_viewer action=set_marks`** —— 把总结出的要点连同锚点一次写进去。
   `marks` 是数组，每项 `{ text, note?, index?, norm?, occurrence? }`；
   不传 `index` 时由 Host 按 `text` **精确匹配**回原文那一段。

为什么用 `set_marks` 而不是循环 `add_mark`：

- **原子性**：任何一条对不上原文，整组不落盘——不会留下「标了一半」的文档
- **一次往返**：一次总结产出的是一组要点，不是一条
- **逐条回执**：返回每条的命中段落，没命中的会说明原因

`replace: true` 表示覆盖原有标记，默认追加。

### 一个真实的调用序列

```
doc_viewer action=open       filePath=D:\招标文件.docx
doc_viewer action=get_digest docId=d3f6...              → 拿到分节 + 每节的锚点
（模型据此写出 N 条要点）
doc_viewer action=set_marks  docId=d3f6... marks=[{text,note,index,norm,occurrence}, ...]
```

写完后用户在面板「标记」页签里就能看到，点一条跳到对应段落。

### 让 AI 评判某几段并打上锚点（通用）

上面那条链路是「总结」，评判是它的近亲——**读若干段 → 给出判定 → 按判定落锚**。
评判比总结多几个结构化字段，它们决定面板怎么显示：

| 字段 | 取值 | 作用 |
| --- | --- | --- |
| `severity` | **任意文本**（推荐 `fatal/high/medium/low/info`） | 决定左边框颜色。内置值有中文标签；自定义值原样显示，颜色按词哈希生成 |
| `category` | **任意文本**（推荐 `invalid/qualification/...`） | 面板上显示成分类型徽标。自定义值原样显示 |
| `color` | `#hex` / `rgb()` / `hsl()` / 颜色名 | 可选。不填则按 severity 自动配色 |
| `suggestion` | 任意文本 | 「该怎么改」，与 `note`（判定理由）分开显示 |

为什么允许自由文本：写死枚举等于替用户决定词汇表。招投标要「废标/资格」，合同审阅要
「付款条款/违约责任」，病历要「危急值」——一套枚举服务不了所有场景。

面板的应对是**字符串哈希配色**：同一个词永远得到同一个颜色，不需要预定义色板。
内置值额外带中文标签，让常见场景读起来更顺；自定义值原样显示。

标签会做两道清理：去掉控制字符（换行会破坏徽标的单行布局）、截断到 40 字（防止撑爆界面）。
`color` 字段会校验格式，非法值归空并退回自动配色——不会把任意字符串塞进 style。

### 一个评判的调用序列

```
doc_viewer action=open        filePath=D:\招标文件.docx
doc_viewer action=read_section docId=... index=1 before=0 after=2   → 拿到这 3 段与它们的 index
（模型逐段判定）
doc_viewer action=set_marks   docId=... marks=[
  { text: "3.3 本项目不接受联合体投标。", severity: "fatal", category: "invalid",
    note: "联合体投标将被否决", suggestion: "确认以单一法人主体投标",
    index: 3 },
  ...
]
```

`index` 可以省略，Host 会按 `text` 精确匹配；但**从 `read_section` 拿到的
`index` 直接带上更稳**——同名段落（比如两处「附件一」）只有靠它才能对上各自那一处。

## 真实跑一遍长什么样

拿一份 29 段的《设备采购合同（示例）》跑通全程，工具面的实际输出（为篇幅起见只保留前几节；路径已换成示例路径，其余逐字保留）：

```
doc_viewer action=open filePath=D:\样例\设备采购合同（示例）.docx

已登记 dc6eacfa9303a204d：设备采购合同（示例）
路径：D:\样例\设备采购合同（示例）.docx
统计：29 段 / 15 标题 / 15 条大纲
大纲：
- [0] 设备采购合同（示例）
- [1] 第一章 总则
  - [2] 1.1 合同双方
  - [5] 1.2 合同标的
- [8] 第二章 交付与验收
  - [9] 2.1 交付时间
  - [11] 2.2 验收标准
  ...
```

```
doc_viewer action=set_marks docId=dc6eacfa9303a204d replace=true marks=[...]

已写入 4 条标记（该文档现有 4 条）。
- mmwx0v1e40o3x @6 [medium/commercial]：乙方向甲方供应数控加工中心 3 台，型号 CX-880，详见附件一《供货清单》。（标的物只写了型号，未写清配置与备件清单，交付时容易扯皮。） → 把附件一的配置表与技术参数一并写进合同正文，注明「附件与正文同等效力」。
- mmwx0v1e51ei5 @13 [high/contract]：验收不合格的，乙方应在 15 日内免费更换或修复；逾期未处理的，甲方有权解除合同并要求赔偿。（解除权写得清楚，但「要求赔偿」没有约定赔偿范围与上限。） → 补充「赔偿包括甲方因此产生的直接损失与第三方检测费用」。
- mmwx0v1e52dol @18 [low/commercial]：设备验收合格后 30 日内支付 60%，余款 10% 作为质保金，质保期满且无质量问题的，一次性无息退还。（质保金比例合理，但未写质保期长度，需要在第四章或附件里补齐。） → 明确质保期为验收合格之日起 12 个月。
- mmwx0v1e53n7v @23 [fatal/risk]：乙方逾期交付的，每逾期一日按合同总价的 0.5‰ 支付违约金，累计不超过合同总价的 5%。（违约金上限 5% 约为 10 天，逾期两个月也只能拿到 5%，对甲方约束力偏弱。） → 上限提到 10%，并约定逾期超过 30 日甲方可单方解除合同。
```

`@6` 是段落序号——面板里点这一条，右侧文档就滚到那一段。

## 数据与边界

- 只**读**你的文档，不改它。
- 只写插件自己的数据目录：`<profile>/doc-anchor-viewer/`（`docs.json` 索引 + `docs/<docId>.json`）。
- 「从本机选择…」上传的文件会复制到该目录下的 `uploads/`；移除登记不会删除源文件。
- 单文件上限 `maxDocxMb`（默认 30MB）——预览要把整个文件 base64 送进浏览器。
- **解析时只解 ZIP 里需要的两个条目**（`word/document.xml`、`word/styles.xml`）：图片、嵌入对象一概不解压。一份 8MB 的文档，省下的是 8MB 常驻内存与一次全文 inflate。
- **解压有上限**：单条目 64MB、全部条目合计 256MB。docx 就是 ZIP，而 ZIP 能被造出来骗过「文件 ≤30MB」这道检查——实测一个 65KB 的条目能膨胀出 64MB。超限的文件会被拒绝并给出可读提示，不会把宿主进程撑爆。

## 权限、外部服务与风险（请先读这一节）

| 项目 | 实际情况 |
| --- | --- |
| 文件读取 | 只读你用面板或 `doc_viewer` 明确指出的那一个 `.docx`；不会扫盘 |
| 文件写入 | 只写 `<profile>/doc-anchor-viewer/` 下的索引、段落缓存与 `uploads/`；不碰源文件 |
| 网络 | 插件自身**不发起任何网络请求**，无遥测、无上报 |
| 数据流向 | 文档字节通过本机 loopback（`127.0.0.1`）送给浏览器渲染，不出本机 |
| **会进模型上下文的部分** | 你让助手执行 `get_digest` / `read_section` / `get_outline` 时，**文档正文会作为工具结果进入该会话的模型上下文**，随后按你 DSH 的模型配置发往对应模型服务商。不需要让文档离开本机时，就别让助手去读它 |
| 关闭插件 | 移除或停用后，面板与 `doc_viewer` 一起消失；已登记的数据目录保留，删不删由你决定 |

不需要任何 API Key，也不依赖任何外部服务。

## 配置

`cordis.patch.yml` 里两行：

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `dataDir` | `''` | 留空按 profile → DSH_HOME → 当前目录顺序自动选址 |
| `maxDocxMb` | `30` | 单个 docx 大小上限（MB） |

## 兼容性

- **DSH**：在 DSH Desktop（Windows）上开发与验证，其内置的 `@deepseek-ai/dsh` 为 `0.2.2-rc.5`；更早的版本没有保证。
- **宿主能力**：Host 半边是纯 Node，`Windows / macOS / Linux` 通用；面板与侧栏标签页需要 Web GUI（`dsh --profile web` 或 Desktop）。
- **文件格式**：只支持 `.docx`。`.doc`（97-2003 二进制）需要先另存为 `.docx`。
- **可选依赖**：宿主没有 `tools` 服务时插件其余部分照常工作，只是不注册 `doc_viewer`；宿主没有右侧栏时侧栏标签页自动跳过并打一条 warn，不影响全局面板。

## 常见问题

**点了左边条目，右边不跳 / 提示「这一段没能定位到」。**

先确认右侧文档渲染完了。匹配不上时插件会**如实提示**，不会跳到错误的位置——原因可能是这一段在表格里（docx-preview 会拆表），或者它和相邻段落的文字完全相同。换个段落试试；同一篇文档重新打开后锚点仍然有效。

**支持 `.doc` 吗？**

不支持。只认 `.docx`；97-2003 的二进制 `.doc` 请先在 Word 里另存为 `.docx`。

**文档会离开我的电脑吗？**

插件本身不发起任何网络请求。但**你让助手读文档时**，`get_digest` / `read_section` 的结果会作为工具结果进入会话上下文，按你的模型配置发往对应模型服务商。不想让文档出本机，就不要让助手去读它——面板里的浏览、标记、跳转全在本机完成。

**能开两个 profile 共用一份数据吗？**

数据目录按 profile → `DSH_HOME` → 当前目录的顺序自动选址（见「配置」）。要让多个 profile 共用，在 `cordis.patch.yml` 里把 `dataDir` 写成同一个绝对路径。

**装完左侧栏没有「文档锚点」入口？**

1. 刷新一次页面——面板代码要浏览器重新加载才会出现；
2. 确认包进的是当前 profile（`dsh plugin --profile web list | grep doc-anchor` 之类）；
3. 若从 GitHub 装，`pnpm` 可能因为构建脚本把安装拦下了，按它打印的提示在 `pnpm-workspace.yaml` 里加 `allowBuilds` 后重跑。

## 实现笔记（给开发者）

这一节记录两个必须记住的坑，改代码前先看，能省半天。

### 锚点晚于渲染到达

`DocxView` 的渲染 effect **只能依赖 `docId`**：

```js
useEffect(() => { /* renderAsync + stampAnchors */ }, [docId])
```

因为 `anchors` 来自父组件异步加载的 `doc`（`getDoc` 那个 Remote 回来才有值），
而渲染一篇几万段的文档要几秒。两者到达顺序不稳定：

- **若 anchors 在渲染时还是空数组** → 一个锚点都打不上 → 点了不跳（实测踩过）
- **若把 anchors 加进依赖** → 锚点一到就整篇重渲染一遍（几秒白屏）

正确做法：渲染只按 `docId` 做一次，锚点通过 **ref** 读；另设一个依赖
`[anchors, state]` 的 effect，锚点后到就**只补打锚点**、不动已渲染的 DOM。
补打前检查 `p[data-anchor]`，保证幂等。

测试 `T1`～`T4` 钉死了这条：T1 模拟「先渲染后到锚点」，T4 检查 effect 依赖没被改坏。

### `typertRemote` 绑定

客户端报 `docAnchor/xxx failed: ... HTTP 404` 时，先查这一条，
不要去怀疑「插件没装」或「宿主没重载」——fiber 可能明明是 `active`。

Gateway 的服务发现要求两处**都**存在：

```js
// ① 实例上的绑定（构造函数里）——缺它就整条路由 404
this.typertRemote = { service: this, serviceKey: 'docAnchorViewer', namespace: 'docAnchor' }

// ② prototype 上的方法列表（apply 里）——缺它会端点级失败
Object.defineProperty(proto, '@deepseek-ai/dsh-typert-protocol/remote-methods', { value: { version: 1, methods: [...] } })
```

只写 ② 时，命名空间根本不会进路由表，表现为**整个命名空间 404**（不是「方法不存在」）。
这个区别很重要：方法级失败会提示 `gateway/method-unavailable`，而命名空间级失败只有一句
HTTP 404，看起来就像插件压根没起来。测试 C1b 钉死了这条契约。

### 为什么锚点不按段落序号匹配

Host 用自己的解析器算段落（9 千段级的文件是毫秒级），渲染交给浏览器的 docx-preview。两个实现**不可能**对空段落、域代码、文本框的处理完全一致，所以序号一定会错位。

锚点因此由三部分组成：`index`（给人和 Agent 看的稳定编号）+ `norm`（去掉全部空白后的全文）+ `occurrence`（这串文字在全文里第几次出现）。客户端渲染完按后两个键**精确匹配**，匹配不上就如实提示「这一段没能定位到」，不会跳到错误的位置。

匹配**绝不用 `includes`**：docx-preview 会渲染出大量表头与序号（「序号」「条款号」「1」「1.1」），包含匹配会把它们全锚到某条长段上，表现为「点哪都跳到同一处」。

## 开发

```bash
npm install            # 安装 docx-preview 与 esbuild
npm run build:client   # 把 docx-preview 打成 IIFE 注入 lib/client.js 的标记区
npm test               # 8 个文件、77 项：解析 / 锚点规则 / Host 服务 / 契约 / 渲染 / 侧栏 / ZIP 容器安全
```

`lib/client.js` 是一份手写的 Client 模块，docx-preview 以 IIFE 内联在标记注释之间（DSH 的浏览器模块表只提供 `react`，不解析插件的 node_modules，页面也没有可引用的静态资源 URL，内联是唯一刷新即生效的方式）。改完 client.js 记得跑一次 `build:client`。

目录：

```
lib/index.js     Host 半边：登记、解析、大纲、标记、Remote 面、doc_viewer 工具
lib/client.js    浏览器半边：全局面板 + 侧栏标签页（含内联的 docx-preview）
lib/docparse.js  段落 / 样式 / 大纲 / occurrence 的确定性解析
lib/ooxml.js     最小 ZIP + OOXML 读取
lib/store.js     数据目录选址与落盘
locale/          zh / en 文案（Plugin Manager 与详情页读这里）
cordis.patch.yml 把 Host 那一行插进 profile
```

入口就是 `lib/index.js` 里的 `export function apply(ctx, config)`（外加 `Config` 的 `~standard` 校验），符合 DSH 插件规范；`cordis.patch.yml` 负责把 Host 那一行插进 profile，浏览器半边由 `dsh.client.platform = "web"` 声明。

**发布**：`.github/workflows/npm-publish.yml` 在 GitHub Release 发布时用 Trusted Publishing（OIDC）跑 `npm ci && npm test && npm publish`，不需要 `NPM_TOKEN`。

## 分类与许可

- **分类**：Integrations & Connections → Office & docs（同时提供 Tools & Capabilities 侧的 `doc_viewer` 工具）
- **许可证**：[MIT](LICENSE)
- **DSH 插件标记**：仓库 topic `dsh-plugin`

---

## English

A DSH plugin that renders `.docx` files with Word's own layout and turns the outline plus your own marks into jump targets.

- **Panel**: the left rail lists the outline and the marks; clicking one scrolls the document to that paragraph and flashes it.
- **Agent tool** `doc_viewer`: `open` → `get_digest` (a sectioned digest whose sections already carry anchors) → `set_marks` (batch-write the findings, atomically). Or `read_section` to read just a few paragraphs by anchor.
- **Free-text severity/category**: `fatal/high/medium/low/info` are suggestions, not an enum. Colors are hashed from the word, so any vocabulary works (contract review, clinical notes, code review…).
- **Read-only**: documents are never modified. The plugin writes only to its own data directory and makes no network requests of its own. Note that text you ask the agent to read becomes part of that conversation's model context.

```bash
dsh plugin --profile web add dsh-doc-anchor-viewer
```

Requires a DSH Web GUI (or DSH Desktop). `.docx` only — convert `.doc` first. MIT licensed; not affiliated with DeepSeek.

---

## 免责声明

这是一个社区插件，**与 DeepSeek 官方无隶属关系**，不代表官方立场，也未获得官方背书。
DeepSeek Harness 与 DSH 是其各自权利人的名称与标识。
