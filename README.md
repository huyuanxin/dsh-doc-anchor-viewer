# 文档锚点视图（dsh-doc-anchor-viewer）

把 `.docx` 按 **Word 自己的版式**渲染出来，并把「大纲」与「自定义标记」做成可跳转的锚点：点左边、右边滚到位并闪一下高亮。Agent 也能用同一套锚点读原文，不必把整篇文档塞进上下文。

**这是一个通用工具，不绑定任何行业。** 标记的「严重程度」与「分类」都是**自由文本**：
合同审阅可以写「高风险 / 付款条款」，病历批注可以写「危急值 / 检验结果」，代码评审可以写「P0 / 空指针」。
内置的 `fatal/high/medium/low/info` 只是**推荐词汇**（面板自带中文标签），不是白名单。

## 面板在哪

侧栏（左侧导航）里的 **「文档锚点」** 入口。打开后：

- 左栏：文档列表 → 大纲 / 标记两个页签
- 右栏：docx-preview 渲染的 Word 文档

## 两种打开方式

1. **在面板里打开**：左栏「打开」→ 粘贴本地绝对路径，或「从本机选择…」选一个文件（会复制到插件数据目录）。
2. **在会话里让助手打开**：`doc_viewer` 工具，`action: "open"`，`filePath` 给绝对路径。打开后面板里立刻能看到。

## 锚点怎么用

- **大纲锚点**：来自 Word 的标题级别（`styles.xml` 的 `outlineLvl`，含 `basedOn` 继承；也认 `Heading N` / `标题 N` 样式名）。点一下就跳到那一段。
- **标记锚点**：把文档滚到想记住的位置，点左栏底部「标记当前段落」，可以加备注。标记会记住它是**哪一段**（段落序号 + 归一化全文 + 同文第几次出现），即使文档重新打开也能跳回去。
- **Agent 的读法**：`doc_viewer` 的 `get_outline` 拿大纲，`read_section` 按 `index` 或一段原文取回前后若干段。这就是「锚点」对模型的意义——按需读一小段，而不是全文。

## 让会话总结文档并打上锚点

这是一条完整的、给 Agent 用的两步链路：

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

## 让 AI 评判某几段并打上锚点（通用）

上面那条链路是「总结」，评判是它的近亲——**读若干段 → 给出判定 → 按判定落锚**。
评判比总结多三个结构化字段，它们决定面板怎么显示：

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

## 一个必须记住的坑：锚点晚于渲染到达

如果表现为「文档能正常显示，但点左边什么都不跳」，先查这条。

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

测试 `T1`～`T4` 钉死了这条：T1 模拟"先渲染后到锚点"，T4 检查 effect 依赖没被改坏。

## 一个必须记住的坑：typertRemote 绑定

如果客户端报 `client api: docAnchor/xxx failed: ... HTTP 404`，**先查这一条**，
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

## 为什么锚点不按段落序号匹配

Host 用自己的解析器算段落（9 千段级的文件是毫秒级），渲染交给浏览器的 docx-preview。两个实现**不可能**对空段落、域代码、文本框的处理完全一致，所以序号一定会错位。

锚点因此由三部分组成：`index`（给人和 Agent 看的稳定编号）+ `norm`（去掉全部空白后的全文）+ `occurrence`（这串文字在全文里第几次出现）。客户端渲染完按后两个键**精确匹配**，匹配不上就如实提示「这一段没能定位到」，不会跳到错误的位置。

匹配**绝不用 `includes`**：docx-preview 会渲染出大量表头与序号（「序号」「条款号」「1」「1.1」），包含匹配会把它们全锚到某条长段上，表现为「点哪都跳到同一处」。

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

## 数据与边界

- 只**读**你的文档，不改它。
- 只写插件自己的数据目录：`<profile>/doc-anchor-viewer/`（`docs.json` 索引 + `docs/<docId>.json`）。
- 上传的文件会复制到该目录下的 `uploads/`；移除登记不会删除源文件。
- 单文件上限 `maxDocxMb`（默认 30MB）——预览要把整个文件 base64 送进浏览器。

## 配置

`cordis.patch.yml` 里两行：

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `dataDir` | `''` | 留空按 profile → DSH_HOME → 当前目录顺序自动选址 |
| `maxDocxMb` | `30` | 单个 docx 大小上限（MB） |

## 开发

```bash
pnpm install          # 安装 docx-preview 与 esbuild
pnpm run build:client # 把 docx-preview 打成 IIFE 注入 lib/client.js 的标记区
pnpm test             # 18 项：解析 / 锚点规则 / Host 服务 / 契约
```

`lib/client.js` 是一份手写的 Client 模块，docx-preview 以 IIFE 内联在标记注释之间（DSH 的浏览器模块表只提供 `react`，不解析插件的 node_modules，页面也没有可引用的静态资源 URL，内联是唯一刷新即生效的方式）。改完 client.js 记得跑一次 `build:client`。
