# PPTX 在线编辑器 + Hyperframe 视频导出（开发交接文档）

> 面向下一个开发者。读完这份文档 + 配合 Claude Code，你应该能直接继续开发。
> 建议先读「一、目标」「三、核心设计原理」「六、当前进度与覆盖度」「九、后续路线」。

---

## 一、目标

在现有的 `@wisdomgarden/pptx-renderer`（把 `.pptx` 高保真渲染成 HTML/SVG 的库）之上：

1. 让渲染出来的元素能像 Office 一样**编辑**：选中、移动、缩放、旋转、翻转、改文字、改字体/颜色/填充、新增/删除/调层级；
2. 把编辑后的结果**导出给 hyperframe** 生成视频。

已与产品确认的关键决策：

- 编辑能力做到「完整（含增删/层级）」；
- 编辑器放在**独立的 app**（不是塞进库里）；
- 先做**可跑通的 POC**，再逐步补齐。

---

## 二、Hyperframe 是什么（导出目标）

仓库：`github.com/heygen-com/hyperframes`。要点（决定了导出方案）：

- **它消费的是 HTML/CSS，不是 JSON。** 一个 composition 形如：
  ```html
  <div id="stage" data-composition-id="deck1" data-start="0" data-width="1920" data-height="1080">
    <div class="clip" data-start="0" data-duration="5" data-track-index="0">...slide1...</div>
    <div class="clip" data-start="5" data-duration="5" data-track-index="0">...slide2...</div>
  </div>
  ```
- 动画是**暂停且可 seek** 的 GSAP/CSS/Lottie 时间线，注册到 `window.__timelines[compositionId]`；hyperframe 逐帧 seek 渲染成 MP4。
- CLI：`npx hyperframes init | preview | render`；包 `@hyperframes/core`、`@hyperframes/producer`。

**为什么天然契合**：pptx-renderer 本来就输出自包含内联样式的 HTML，所以导出 ≈「再渲染一次编辑后的模型 + 套上 stage/clip/data-\* + 加默认入场动画」，不是格式转换。（导出尚未实现，见 Phase 5。）

---

## 三、核心设计原理（最重要）

**一句话：编辑不是改渲染出来的 DOM，而是改「模型」，然后让原渲染器重新画一遍。**

库本身是一个近似纯函数的 `模型(PresentationData) → DOM` 渲染器。我们只往回补了一条 `DOM → 模型` 的映射线，其余全靠改模型 + 重渲染。

三个「支点」（都是库原本的行为，没改）：

1. **模型可变、可重渲染**：`renderSlide(presentation, slide)` 每次都从普通对象图 `PresentationData` 重新画。改完模型再调一次就更新，不需要增量 patch。
2. **几何是 typed 的**：每个节点的 `position/size/rotation/flipH/flipV`（`src/model/nodes/BaseNode.ts`）是像素/角度普通数字。移动/缩放/旋转 = 直接改这些数。
3. **渲染器「直接读」模型上的引用**（样式能改的关键）：
   - 文字内容读 `run.text`；
   - 文字样式读 `run.properties`（即 `<a:rPr>`，`SafeXmlNode`）；
   - 形状填充读 `node.fill`（`src/renderer/ShapeRenderer.ts` 约 1689 行，且有 `resolveFill(spPr)` 的兜底）。
   - `SafeXmlNode` 通过 `.element`（`src/parser/XmlParser.ts:105`）暴露底层 DOM 元素。

由此推出每个操作的实现：

| 操作                | 原理                                                                                                |
| ------------------- | --------------------------------------------------------------------------------------------------- |
| 选中                | 点击 → `el.closest('[data-node-id]')` → 查 `onNodeRendered` 建的 `Map<id,{node,element}>`           |
| 移动/缩放/旋转/翻转 | 覆盖层手柄算新几何 → 写 `node.position/size/rotation/flip` → 重渲染                                 |
| 改字号/颜色/加粗    | **新建** `<a:rPr>`/`<a:solidFill>` 元素 → **重新赋值** `run.properties` / `node.fill` 引用 → 重渲染 |
| 改文字              | 直接改 typed 的 `run.text`（不碰 XML）                                                              |
| 新增元素            | 拼一小段 `<sp>` OOXML → 复用 `parseShapeNode` 解析成真节点 → push 进 `slide.nodes`                  |
| 删除/层级           | `slide.nodes` 数组即 z 序，splice / 重排                                                            |

**为什么不需要 OOXML 序列化器**：改样式时不去改原始 XML 子树，而是把模型上的 `SafeXmlNode` 引用换成新建的游离元素。渲染器读的是这个引用，所以重渲染即生效——既不用把模型写回 OOXML，也没有「缓存旧引用导致改了不生效」的坑。库里那个单向、有损的 `serializePresentation` 完全没用到。

> ⚠️ 这个「引用被换成游离元素」的事实对**持久化**很关键：编辑分散在 (a) typed 字段 和 (b) 被换掉的 `SafeXmlNode` 引用（不在 `node.source` 里）。所以直接存 `node.source.outerHTML` 会丢样式。见「八、持久化」。

---

## 四、对库（pptx-renderer）的改动（很小，约 15 行 + 几个导出）

1. **`src/renderer/RenderContext.ts`**：`RenderContext` 增加可选字段
   `onNodeRendered?: (node: BaseNodeData, element: HTMLElement, ctx: RenderContext) => void`。
2. **`src/renderer/SlideRenderer.ts`**：
   - `SlideRendererOptions` 增加同名 `onNodeRendered?`；
   - `renderSlide` 里把它挂到 `ctx`（`ctx.onNavigate` 赋值附近）；
   - 分发函数 `renderNode`（约 85 行）在返回元素前：若 `ctx.onNodeRendered` 存在，则给元素盖 `dataset.nodeId/nodeType` 并回调。**仅在传了回调时生效**，普通渲染零影响；
   - 母版/版式模板形状用的 `masterCtx/layoutCtx` 里显式 `onNodeRendered: undefined`，避免给不可编辑的模板形状盖标记 / 触发回调。
   - 因为组子节点也走同一个 `renderNode`（`GroupRenderer` 用它做递归），所以**整棵树都被覆盖**，`GroupRenderer.ts` 不用改。
3. **`src/core/Viewer.ts`**：`renderSlideToContainer` 增加可选第 4 参 `options?: Pick<SlideRendererOptions,'onNodeRendered'>` 并转发（可选便利，编辑器其实直接用 `renderSlide`）。
4. **`src/index.ts`** 新增导出（编辑/新增元素用的原语）：
   `parseXml`、`SafeXmlNode`、`parseShapeNode`、`parseTextBody`。

单测：`test/unit/renderer/SlideRenderer.test.ts` 里加了一组 `onNodeRendered editing hook` 用例（标记 data-node-id、逐节点触发、覆盖组子节点、不触发模板形状、无回调时不加属性）。全库单测 3186 通过，typecheck/lint 干净。

---

## 五、编辑器工程结构（`apps/editor/`）

独立 Vite + React + pnpm app，通过 Vite alias 直接吃库的**源码**（永远拿到最新钩子）。

```
apps/editor/
  package.json            # react/vite；含 pnpm.onlyBuiltDependencies:["esbuild"]
  vite.config.ts          # alias '@wisdomgarden/pptx-renderer' -> ../../src/index.ts；server.fs.allow 到仓库根；端口 5273
  tsconfig.json           # 注意：未开 noUnusedLocals/Parameters（因为要 typecheck 库源码，避免库里 _ 前缀噪音）
  index.html
  editor.md               # 本文档
  src/
    main.tsx              # 挂载。⚠️ 故意不套 <StrictMode>（见坑位）
    EditorApp.tsx         # React 外壳：工具栏（打开/翻页/+Text/+Rect/层级/删除）、舞台容器、右侧面板；
                          #   创建 RenderHost；把 host 挂到 window.__pptxHost（调试用）；?sample= 自动加载
    styles.css            # 深色 UI；.app 设 user-select:none（拖手柄不误选文字），输入框恢复可选
    panels/
      SelectionPanel.tsx  # 右侧 Inspector + 样式面板（填充色/文字色/字号/字体/加粗斜体 + ✏️Edit text 按钮）
    editor/
      RenderHost.ts       # 【核心控制器，框架无关】加载/渲染/缩放/选中/变换/增删/层级/文字编辑/样式，全在这里
      TransformController.ts  # 选择框 + 8 缩放手柄 + 旋转手柄，旋转感知的变换数学，回调 onPreview/onCommit/onBoxPointerDown
      StyleFacade.ts      # 读/写样式：读 rPr/fill 显示；写=新建元素并重新赋值 run.properties/node.fill；setPlainText 改文字
      nodeOps.ts          # 新增(createTextBox/createRectangle，拼 OOXML→parseShapeNode)、删除、层级(reorder)、freshId
```

**RenderHost 的 DOM 结构**：`mount > stageWrap > [slideEl(缩放后的渲染结果), overlayLayer(覆盖层，放选择框/手柄/文字编辑 textarea)]`。overlayLayer 用内在坐标（幻灯片 px），随 slideEl 同一个 `scale(k)`。

**关键约定**：所有几何用内在幻灯片像素；屏幕像素增量要 `/scale`。选中只认**顶层** `slide.nodes`（组按整体，命中测试向上回溯到最近的顶层 `data-node-id`）。

---

## 六、当前进度与覆盖度

### 已完成阶段

- **Phase 0 库钩子**：`onNodeRendered` + `data-node-id`（见上）。✅ 有单测。
- **Phase 1 骨架 + 渲染宿主**：加载 pptx、fit 渲染、翻页、点击选中、Inspector。
- **Phase 2 变换**：移动（拖框体，实时平移元素）/ 8 向缩放（旋转感知，保持对边固定）/ 旋转 / 翻转，指针抬起提交 + 整页重渲染。
- **Phase 3 结构**：+文本框 / +矩形（种子 OOXML→真解析器）、删除（Del 键 / 按钮）、层级（前移/后移）、Esc 取消选中。
- **Phase 4 文字 + 样式**：双击（或 ✏️Edit text 按钮）内联编辑文字；属性面板改 填充色 / 文字色 / 字号 / 字体 / 加粗斜体。

### 覆盖度矩阵（诚实版）

「变换类」只碰每个节点都有的 `position/size/rotation`，所以全类型生效；「内容/样式类」目前仅 shape。

| 能力                     | 覆盖                                                         |
| ------------------------ | ------------------------------------------------------------ |
| 选中/移动/缩放/旋转/翻转 | ✅ 所有顶层类型（shape/picture/table/group/chart），组按整体 |
| 删除 / 层级              | ✅ 所有顶层类型                                              |
| 改文字内容               | ⚠️ 仅 shape                                                  |
| 字体/字号/颜色/加粗斜体  | ⚠️ 仅 shape，且整形状所有 run 一起改                         |
| 填充                     | ⚠️ 仅 shape，且仅纯色                                        |
| 新增元素                 | ⚠️ 仅 文本框 + 矩形                                          |

### renderer 能渲染、但还不能编辑（明确缺口）

- **表格**：单元格文字/样式、增删行列、列宽、合并（数据在 `TableNodeData.rows[].cells[].textBody`）
- **图片**：裁剪、替换图片、描边
- **图表**：任何数据/系列/类型编辑（chart 数据在外部 XML，最复杂）
- **组内子元素**单独编辑（现为原子）
- **文字**：按选区/按 run 局部格式、项目符号、对齐、行距、超链接
- **线条/箭头**：线宽、颜色、虚线、箭头端
- **形状 adj 手柄**（黄色菱形）、自定义几何
- **幻灯片级**：背景、增删页、备注；母版/版式（刻意不开放）

---

## 七、踩过的坑（务必知道，避免回退）

1. **不要套 React `<StrictMode>`**：本编辑器用命令式 DOM + `useEffect` 里的 RenderHost 生命周期；StrictMode 开发期双挂载会创建**两个** RenderHost（两个 textarea、重复监听器），导致内联编辑提交读到错误的编辑框、内容丢失。`main.tsx` 已去掉。
2. **原生 `dblclick` 不可靠**：首次点击选中后会盖上选择框，双击两次点击落在不同元素上，浏览器不触发 dblclick。改成**手动双击检测**（按节点 id + 时间间隔，覆盖舞台和选择框两条路径，见 `RenderHost.registerClick` 与 `TransformController.onBoxPointerDown`）。
3. **编辑框会被双击的尾随 click 立刻 blur 关闭**：加了 blur 保护（打开后 350ms 内的 blur 忽略并重新聚焦）。
4. **点进编辑框会清空选中**：textarea 的 `pointerdown` 冒泡到舞台 → `select(null)`。已对 textarea 的 `pointerdown/mousedown` `stopPropagation`；且提交改用**记住的 `editingShape`**（不依赖当前选中），提交后保持该形状选中。
5. **中文输入法回车**：IME 合成中的 Enter 只是确认候选，不能当提交；handler 里 `if (ev.isComposing || keyCode===229) return`，并同时认 `key==='Enter' || keyCode===13`。（自动化工具发的 Return 会变成空 key，这是自动化特有现象，真实键盘正常。）
6. **图片 blob URL 抖动**：每次重渲染各自建/撤销 blob 会报 `ERR_FILE_NOT_FOUND`。已改为**共享一个持久 `mediaCache`**（`renderSlide({mediaUrlCache})`），跨重渲染复用、不撤销；`destroy` 时统一撤销。
7. **样式回写要「换引用」而不是「改原 XML 子树」**：因为 `node.fill` 等是解析时缓存的引用；替换原子树会让引用悬空、改了不生效。正确做法见「三」。
8. **内置浏览器面板会自己变尺寸**：调试时坐标点击很飘。用 `window.__pptxHost` 直接驱动做确定性验证更稳。

---

## 八、持久化（**尚未实现**，给出方案，建议先做 A）

不能直接 `JSON.stringify` 模型：模型带活的 DOM 引用；且编辑分散在 typed 字段 + 被换掉的游离 `SafeXmlNode` 引用（不在 `node.source` 里）。

- **方案 A（推荐给 POC）：编辑操作日志 + 原 pptx**。存「原始 pptx 字节 + 一份 JSON 操作记录（按节点 id：transform/text/style/added/deleted/order）」。加载 = `parseZip(原pptx)`→`buildPresentation`→对每条记录**重放**已有的 mutator（`applyTextStyle`/`applyFillColor`/`setPlainText`/`createRectangle`/`reorderNode`…）。JSON 安全、体积小、和「编辑=改模型」架构天然契合。依赖 id 稳定（pptx 的 `cNvPr id` 稳定；新增用生成的 `new-*` id）。
  建议 JSON 形状：
  ```jsonc
  { "sourceRef":"deck.pptx",
    "slides": { "0": {
      "transforms": {"2":{"x":80,"y":47,"w":1120,"h":44,"rot":0,"flipH":false,"flipV":false}},
      "text":  {"2":"标题"},
      "style": {"51":{"fill":"12B886"}, "2":{"color":"C00000","bold":true,"size":44}},
      "added": [{"id":"new-1","kind":"rect","box":{"x":..,"y":..,"w":..,"h":..},"fill":"4472C4"}],
      "deleted": ["9"],
      "order": ["new-1","2","51"]
    }}}
  ```
- **方案 B（更稳，工作量中）：无损模型快照**。把 `serializePresentation` 扩成可逆：每节点存 typed 字段 + 当前生效 rPr/fill 的 `outerHTML` + presetGeometry/adjustments；加载用 `parseXml` 还原挂回。自包含、不依赖原 pptx，但要为每类节点写序列化/反序列化。
- **方案 C（真正可移植，工作量大）：写回 .pptx**。需要一个 OOXML writer，才能在 PowerPoint 里也能打开。库目前**没有** writer，是独立大工程。

> 落地建议：先 A（存 localStorage 或下载/上传 `.json`，进来自动重放恢复），要「导出标准 pptx / PowerPoint 打开」时再上 C。

---

## 九、后续路线（下一个 dev 从这里挑）

1. **Phase 5：Hyperframe 导出**（原计划的收尾）
   - 每张幻灯片 `await handle.ready` 后取 `element.outerHTML`；
   - 包成 `class="clip"` + `data-start=i*SLIDE_SECONDS` + `data-duration` + `data-track-index=0`；
   - 外层套 `<div id="stage" data-composition-id data-width="1920" data-height="1080">`；
   - 每个 clip `transform: scale(1920/presentation.width)`；
   - 默认入场动画：内联 `<script>` 建暂停 GSAP 时间线挂 `window.__timelines[id]`（整片淡入 + 逐元素借 `data-node-id` 错峰滑入，标题优先，用 `node.placeholder` 判断）；
   - 产出可被 `npx hyperframes preview|render` 跑的 `index.html`。
   - 新文件建议：`apps/editor/src/export/toHyperframe.ts`。
2. **持久化方案 A**（见八），让「下次进来接着编辑」。
3. **补编辑覆盖度**（优先级建议）：表格单元格文字/样式 → 图片裁剪/替换 → 文字按选区格式/对齐/项目符号 → 渐变/图片填充、线条样式 → 组内子元素 → 图表。
4. 体验：撤销/重做（快照几何 + 逐节点当前 rPr/fill 的 outerHTML）、多选、键盘微调。

---

## 十、如何运行与调试

```bash
# 依赖（首次）：仓库根先装库依赖；再装 app 依赖
cd /path/to/pptx-renderer && pnpm install
cd apps/editor && pnpm install     # 若报 esbuild 构建被忽略：package.json 已加 pnpm.onlyBuiltDependencies

# 起开发服务器（app 目录）
pnpm dev                            # http://localhost:5273
```

- 打开后「Open .pptx」上传，或拖拽 `.pptx` 进舞台。
- **开发期快速加载样例**：`http://localhost:5273/?sample=/@fs/绝对路径/source.pptx`（`vite.config` 已放开 fs.allow 到仓库根）。仓库自带样例在 `docs/example/*/source.pptx`。
- **确定性调试**：`window.__pptxHost` 暴露了 RenderHost，可在控制台直接调：
  ```js
  const h = window.__pptxHost;
  h.currentSlide.nodes.map((n) => ({ id: n.id, type: n.nodeType, name: n.name }));
  h.select('2');
  h.applyTextStyle({ color: '#C00000', bold: true, fontSize: 44 });
  h.applyFillColor('#12B886');
  h.setSelectedText('新文字'); // 或 h.editSelectedText() 弹内联编辑框
  h.addRectangle();
  h.addTextBox('Hi');
  h.reorderSelected('backward');
  h.deleteSelected();
  ```
- 校验：`cd apps/editor && ./node_modules/.bin/tsc --noEmit`（app 类型检查）；库侧 `pnpm typecheck && pnpm lint && npx vitest run`。

### 交互操作一览（当前）

- 单击选中；`Esc` 取消选中；`Del/Backspace` 删除选中。
- 拖框体移动；拖 8 个手柄缩放；拖顶部圆点旋转。
- 双击文字 或 面板「✏️ Edit text」进入内联编辑：`Enter` 提交（`Shift+Enter` 换行）、`Esc` 取消、点别处也提交。
- 工具栏：翻页、+Text、+Rect、后移/前移、删除。

---

## 十一、关键文件/行号索引

库（改动/依赖点）：

- `src/renderer/SlideRenderer.ts` — `renderNode`(~85)、`SlideRendererOptions`(~34)、`renderSlide` 挂 ctx(~287)、模板 ctx strip
- `src/renderer/RenderContext.ts` — `onNodeRendered` 字段
- `src/parser/XmlParser.ts:105` — `SafeXmlNode.element`（样式回写的底层）
- `src/model/nodes/ShapeNode.ts` — `parseShapeNode`(202)、`findFill`(165)、`parseTextBody`(141)
- `src/renderer/ShapeRenderer.ts` — 读 `node.fill`(~1689)
- `src/index.ts` — 导出 `parseXml/SafeXmlNode/parseShapeNode/parseTextBody`

编辑器：

- `apps/editor/src/editor/RenderHost.ts` — 控制器（选中/变换/增删/层级/文字/样式/媒体缓存/双击检测）
- `apps/editor/src/editor/TransformController.ts` — 选择框与手柄、旋转感知变换数学
- `apps/editor/src/editor/StyleFacade.ts` — 样式读写（换引用）、`setPlainText`
- `apps/editor/src/editor/nodeOps.ts` — 新增/删除/层级/`freshId`
- `apps/editor/src/panels/SelectionPanel.tsx` — 属性面板
- `apps/editor/src/EditorApp.tsx` — 外壳、`window.__pptxHost`、`?sample=`
