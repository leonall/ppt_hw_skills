# html2pptx 增强版技术文档

> 基于原版 `html2pptx.js` 的四轮功能增强：表格支持、CSS 渐变支持、Canvas/SVG 截图嵌入、Overflow 处理模式。
>
> 文档版本：2026-03-18

---

## 目录

1. [项目概述](#1-项目概述)
2. [文件目录结构](#2-文件目录结构)
3. [快速开始](#3-快速开始)
4. [增强特性一：HTML 表格支持](#4-增强特性一html-表格支持)
5. [增强特性二：CSS 渐变支持](#5-增强特性二css-渐变支持)
6. [增强特性三：Canvas / SVG 截图嵌入](#6-增强特性三canvas--svg-截图嵌入)
7. [增强特性四：Overflow 处理模式](#7-增强特性四overflow-处理模式)
8. [完整 API 参考](#8-完整-api-参考)
9. [测试体系](#9-测试体系)
10. [环境配置与运行方式](#10-环境配置与运行方式)
11. [已知限制与降级规则](#11-已知限制与降级规则)

---

## 1. 项目概述

### 原始版本能力

原版 `html2pptx.js` 通过 Playwright 加载 HTML，在浏览器环境中用 `getBoundingClientRect()` + `getComputedStyle()` 提取每个元素的像素坐标和计算样式，转换为 pptxgenjs API 调用，生成 PPTX 文件。

**已支持**：`<div>` 色块、`<p>/<h1>-<h6>` 文字（含 inline 格式）、`<ul>/<ol>` 列表、`<img>` 图片、CSS 旋转、box-shadow、border-radius、placeholder 机制。

**不支持**：`<table>`、CSS 渐变背景、`<canvas>` / `<svg>` 图表、HTML 内容超出画布的处理。

### 增强版新增能力

| 特性 | 解决的问题 |
|---|---|
| **表格支持** | `<table>` 直接映射为 PPT 原生表格，保留样式 |
| **CSS 渐变** | `linear-gradient` / `radial-gradient` 写入真实 OOXML `gradFill` |
| **Canvas/SVG 截图** | ECharts / Chart.js / SVG 图形以 2× 高清 PNG 嵌入 |
| **Overflow 处理** | 内容超出画布时可选等比缩放（expand）或裁剪（clip） |

---

## 2. 文件目录结构

```
html2pptx/
│
├── html2pptx.js               # 主入口（大幅增强，向后兼容原接口）
├── html2pptx-cli.js           # CLI 命令行入口
│
├── overflow-handler.js        # ★ 新增：Overflow 处理模块
├── element-screenshot.js      # ★ 新增：Canvas/SVG 截图模块
├── gradient-parser.js         # ★ 新增：CSS 渐变解析器
├── gradient-xml.js            # ★ 新增：OOXML gradFill XML 生成器
├── gradient-postprocess.js    # ★ 新增：PPTX 后处理（JSZip 替换）
│
├── test-gradient.js           # 渐变功能单元测试（44 用例）
├── test-screenshot-logic.js   # 截图逻辑单元测试（17 用例）
├── test-overflow.js           # Overflow 单元 + 集成测试（22 用例）
├── test-e2e-charts.js         # 图表截图端到端测试
│
├── test-overflow.html         # Overflow 测试 HTML（内容 ≈ 1.5× 画布高度）
└── test-charts.html           # 图表测试 HTML（Canvas + SVG × 4 个图表）
```

### 模块依赖关系

```
html2pptx.js
├── overflow-handler.js        （overflow 处理）
├── element-screenshot.js      （canvas/svg 截图）
│   └── sharp                  （图像处理/降采样）
├── gradient-parser.js         （CSS 渐变解析）
├── gradient-postprocess.js    （PPTX XML 后处理）
│   ├── gradient-xml.js        （OOXML XML 生成）
│   └── jszip                  （ZIP 解压/重打包）
└── playwright                 （浏览器自动化）
```

---

## 3. 快速开始

### 安装依赖

```bash
npm install pptxgenjs playwright jszip sharp
```

### 配置浏览器路径

Playwright 默认下载 Chromium。若环境受限（如容器内），可指定已有 Chrome/Chromium：

```bash
export PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH=/opt/pw-browsers/chromium-1194/chrome-linux/chrome
```

macOS 会自动使用系统 Chrome，无需配置。

### CLI 使用

```bash
# 单个 slide
node html2pptx-cli.js slide.html -o output.pptx

# 多个 slide（按顺序组成演示文稿）
node html2pptx-cli.js slide1.html slide2.html slide3.html -o deck.pptx
```

### 编程 API 使用

```javascript
const pptxgen = require('pptxgenjs');
const { html2pptx, GradientRegistry, postprocess } = require('./html2pptx');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

const registry = new GradientRegistry();   // 渐变注册表（多 slide 共享）

// 处理每张 slide
for (let i = 0; i < slides.length; i++) {
  await html2pptx(slides[i], pptx, {
    slideIndex:       i + 1,          // 1-based，渐变定位用
    registry,                         // 渐变注册表
    overflowMode:     'expand',       // 'expand' | 'clip' | 'error'
    captureCanvas:    true,           // 截图 <canvas>
    captureSvg:       true,           // 截图 <svg>
    screenshotScale:  2,              // 2× 高清截图
    screenshotTimeout: 8000,          // 图表渲染等待超时 ms
  });
}

// 写出 PPTX
await pptx.writeFile({ fileName: 'output.pptx' });

// 渐变后处理（把占位纯色替换为真实 gradFill XML）
await postprocess('output.pptx', registry);
```

### 返回值

```javascript
const { slide, placeholders, screenshots, overflow } = await html2pptx(...);

// slide        — pptxgenjs Slide 对象
// placeholders — [{id, x, y, w, h}] placeholder 元素坐标（用于 addChart）
// screenshots  — [{tag, objectName, position, pngBuffer, dataUrl, error}]
// overflow     — {mode, didTransform, transformInfo, scale, offsetX}
```

---

## 4. 增强特性一：HTML 表格支持

### 问题背景

原版完全不处理 `<table>` 元素，内部的 `<td>/<th>/<tr>` 会被当作普通 div/text 处理，导致表格内容丢失或错位。

### 技术实现

**位置**：`html2pptx.js` 内的 `extractTable()` 函数（浏览器侧）+ `addTable()` 函数（Node.js 侧）

#### 提取流程

```
<table> 被扫描到
    ↓
把整棵子树（td/th/tr/p/...）全部加入 processed 集合
防止被后续文字提取逻辑重复处理
    ↓
extractTable(tableEl) 逐行逐格提取：
  - 行高：tr.getBoundingClientRect().height
  - 列宽：第一行各 td 的 getBoundingClientRect().width
  - 每格：text / fontSize / fontFace / bold / italic / color /
          align / valign / fill / border / margin /
          colspan / rowspan
    ↓
occupiedMap[row][col] 追踪 colspan/rowspan 覆盖区域
被覆盖的格插入 _spanPlaceholder: true 占位对象
    ↓
addTable() 把结构转为 pptxgenjs addTable() 格式
```

#### colspan / rowspan 处理

pptxgenjs `addTable` 要求被合并单元格覆盖的位置填入空字符串占位，而非省略。实现用 `occupiedMap` 二维矩阵跟踪哪些格被占用：

```
表格 DOM（colspan=2 例子）：
  tr[0]: [TD colspan=2] [TD]
  tr[1]: [TD] [TD] [TD]

occupiedMap 标记：
  [0][0] = 占用者自身（不标记）
  [0][1] = true ← colspan 覆盖

最终提取结果：
  row[0]: [{text:'A', colspan:2}, {_spanPlaceholder:true}, {text:'B'}]
  row[1]: [{text:'C'}, {text:'D'}, {text:'E'}]
```

#### 样式降级规则

| HTML 特性 | pptxgenjs 现实 | 降级策略 |
|---|---|---|
| `linear-gradient` 背景 | 不支持渐变 | 取第一个 color-stop 为纯色 |
| `rgba(0,0,0,0)` 透明背景 | 单元格必须有实色 | 返回 `null`，不传 `fill`（白色） |
| `box-shadow` on table | 无表格阴影 | 忽略 |
| `border-radius` on cell | 无单元格圆角 | 忽略 |
| 四边统一边框 | 直接支持 | `{type, pt, color}` 单对象 |
| 混合边框（各边不同） | 支持四边数组 | `[top, right, bottom, left]` |
| 无边框（`border:none`） | 支持 | 四元 `{type:'none'}` 数组 |
| `text-transform` | 无此属性 | 浏览器侧 JS 直接转换文字内容 |
| 极细边框（< 0.5px） | PPT 最小可见 | 最小 `0.25pt` |

---

## 5. 增强特性二：CSS 渐变支持

### 问题背景

pptxgenjs 的 `fill: { type: 'gradient', ... }` API 看似支持渐变，实测证明参数被完全忽略——生成的 XML 里没有任何 `gradFill` 元素。唯一可行的方案是绕过 pptxgenjs 高层 API，在生成 PPTX 后直接修改 OOXML。

### 架构设计：两阶段处理

```
阶段一：html2pptx 正常运行
  渐变元素 → 提取第一个 stop 颜色 → 作为纯色占位交给 pptxgenjs
  同时把 { objectName, gradientCss } 注册到 GradientRegistry

                    ↓

  pptx.writeFile('output.pptx')  ← 含纯色占位的 PPTX

                    ↓

阶段二：postprocess(pptxPath, registry)
  JSZip 解压 PPTX
  按 slideIndex 处理每张幻灯片 XML：
    - 背景渐变：替换 <p:bgPr> 内的 <a:solidFill>
    - Shape 渐变：按 objectName 定位 <p:cNvPr name="...">
                  → 找到所属 <p:sp> → 替换 <p:spPr> 内的 <a:solidFill>
  JSZip 重新打包写回
```

### 模块详解

#### gradient-parser.js — CSS 渐变解析器

解析浏览器 `getComputedStyle` 返回的渐变字符串（与手写 CSS 格式不同，浏览器已将关键字转为数值）：

```
输入：linear-gradient(135deg, rgb(255, 59, 79) 0%, rgb(192, 32, 48) 100%)

输出：{
  type: 'linear',
  angle: 135,
  angOoxml: 18900000,   ← OOXML ang 单位（60000ths of a degree）
  stops: [
    { hex: 'FF3B4F', alpha: 100000, position: 0 },
    { hex: 'C02030', alpha: 100000, position: 100000 },
  ]
}
```

**角度转换公式**：CSS 和 OOXML 的角度参考点不同。

```
CSS:   0deg = 向上，顺时针增大
OOXML: 0 = 向右，顺时针增大，单位 1/60000 度

转换：ooxmlDeg = ((90 - cssDeg) % 360 + 360) % 360
     ooxmlAng = ooxmlDeg × 60000

示例：
  CSS 90deg（向右） → OOXML ang=0
  CSS 180deg（向下）→ OOXML ang=16200000（270° × 60000）
  CSS 135deg       → OOXML ang=18900000（315° × 60000）
```

**stop 透明度**：`rgba(255, 59, 79, 0.5)` 的 alpha 映射为 OOXML `<a:alpha val="50000"/>`（范围 0-100000，100000 = 完全不透明）。

**支持的渐变类型**：

| CSS 类型 | 处理方式 |
|---|---|
| `linear-gradient` | 完整映射，含角度、多 stop、透明度 |
| `repeating-linear-gradient` | 当作 linear-gradient 处理 |
| `radial-gradient` | 映射为 `<a:path path="circle">` + `fillToRect` |
| `repeating-radial-gradient` | 当作 radial 处理 |
| `conic-gradient` | 取第一个 stop 纯色，记录 warning |
| 多层叠加（`url(), linear-gradient(...)`）| 取第一个渐变层 |
| stop 无位置标注 | 线性插值自动填充 |

#### gradient-xml.js — OOXML XML 生成器

生成符合 ECMA-376 §20.1.8.33 规范的 `<a:gradFill>` XML：

```xml
<!-- 线性渐变（135deg CSS = ang=18900000） -->
<a:gradFill rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="FF3B4F"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="C02030"/></a:gs>
  </a:gsLst>
  <a:lin ang="18900000" scaled="0"/>
</a:gradFill>

<!-- 径向渐变（偏心 cx=30% cy=40%） -->
<a:gradFill rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="1A4A8A"/></a:gs>
  </a:gsLst>
  <a:path path="circle">
    <a:fillToRect l="30000" t="40000" r="70000" b="60000"/>
  </a:path>
</a:gradFill>

<!-- 带透明度的 stop -->
<a:gs pos="0">
  <a:srgbClr val="FF3B4F">
    <a:alpha val="50000"/>   <!-- 50% 透明度 -->
  </a:srgbClr>
</a:gs>
```

#### gradient-postprocess.js — XML 后处理器

**Shape 渐变定位逻辑**：pptxgenjs 为每个 shape 生成 `<p:cNvPr name="OBJECT_NAME">` 属性。通过 objectName 精确定位目标 `<p:sp>` 块，只替换其 `<p:spPr>` 段内的 `<a:solidFill>`，不影响 `<p:txBody>` 里的文字颜色。

```
XML 中 shape 的结构：
<p:sp>
  <p:nvSpPr>
    <p:cNvPr name="shape-grad-0">  ← 通过 name 定位
  </p:nvSpPr>
  <p:spPr>
    <a:solidFill>...</a:solidFill>  ← 替换这里
  </p:spPr>
  <p:txBody>
    <a:solidFill>...</a:solidFill>  ← 这里不动（文字颜色）
  </p:txBody>
</p:sp>
```

### 调用方式

```javascript
// 需要 registry 和 postprocess
const { html2pptx, GradientRegistry, postprocess } = require('./html2pptx');
const registry = new GradientRegistry();

await html2pptx('slide.html', pptx, { slideIndex: 1, registry });
await pptx.writeFile({ fileName: 'output.pptx' });
await postprocess('output.pptx', registry);  // ← 这一行写入真实渐变
```

**不使用渐变时**：`registry` 为空，`postprocess` 直接返回，零开销，完全向后兼容。

---

## 6. 增强特性三：Canvas / SVG 截图嵌入

### 问题背景

`<canvas>` 元素（ECharts / Chart.js / D3 等图表库的渲染目标）和 `<svg>` 元素无法逐元素映射为 pptxgenjs 对象——前者是像素画布，后者虽是矢量但 pptxgenjs 不支持 SVG 输入。唯一可行方案是截图后以 PNG 图片嵌入。

### 核心模块：element-screenshot.js

#### 两阶段浏览器会话

```
第一阶段（已有 1× page）：
  用 getBoundingClientRect() 收集所有 canvas/svg 的 CSS 坐标
  → 坐标精确（不受 deviceScaleFactor 影响）

第二阶段（新建 2× context）：
  browser.newContext({ deviceScaleFactor: 2 })
  重新加载同一 URL → page.screenshot({ clip: ... })
  → 截图分辨率是 CSS 尺寸的 2 倍
  sharp.resize(targetW, targetH, { kernel: 'lanczos3' })
  → 降采样回 1× 尺寸，清晰度等同 Retina 渲染

两阶段在同一 browser 进程内，避免重启开销。
```

#### 三级等待策略

图表库通常有渲染延迟（动画、异步数据加载等），过早截图只能拍到空白。

```
优先级 1：window._h2p_ready === true
  → HTML 作者在图表完成后手动设置，最精确
  → ECharts: chart.on('finished', () => window._h2p_ready = true)
  → 触发条件：canvas 设置 data-h2p-wait="flag"

优先级 2：Canvas 像素检测（默认 CANVAS 行为）
  → 轮询 canvas.getImageData() 中心 5×5 区域
  → 有非透明像素即判定就绪
  → 适配没有回调的图表库（Chart.js 等）

优先级 3：双 rAF + 固定延迟（SVG 默认行为及兜底）
  → requestAnimationFrame × 2 确保绘制完成
  → 再等最多 1500ms
  → 无需任何配置，自动生效
```

#### HTML 侧控制接口

通过 `data-*` 属性精细控制截图行为，无需修改 JS 代码：

```html
<!-- 策略1：用 _h2p_ready 标志精确控制 -->
<canvas id="my-chart" data-h2p-wait="flag"></canvas>
<script>
  const chart = echarts.init(document.getElementById('my-chart'));
  chart.setOption({ ... });
  chart.on('finished', () => { window._h2p_ready = true; });
</script>

<!-- 策略2：固定延迟（知道渲染时间时使用） -->
<canvas data-h2p-delay="1000"></canvas>   <!-- 等 1000ms 后截图 -->

<!-- 策略3：什么都不写，自动检测（推荐 Chart.js） -->
<canvas id="bar-chart"></canvas>
```

#### 截图参数计算

```
CSS 坐标（1× page 采集）：
  rect = { x: 100px, y: 50px, width: 300px, height: 200px }

2× page 截图参数：
  clip = { x: 98, y: 48, width: 304, height: 204 }  (加 2px padding 防裁切)

sharp extract（从 2× 图裁掉 padding）：
  extract = { left: 4, top: 4, width: 600, height: 400 }  (padding × scale = 2×2 = 4px)

sharp resize（降采样回 CSS 尺寸）：
  resize(300, 200, { kernel: 'lanczos3' })

最终：300×200 的高清 PNG，清晰度等同 600×400 的原始截图
```

#### 嵌入 PPT

截图结果以 `data URI` 方式嵌入：

```javascript
targetSlide.addImage({
  data: 'image/png;base64,' + pngBuffer.toString('base64'),
  x: position.x,   // 英寸，与 HTML 坐标完全一致
  y: position.y,
  w: position.w,
  h: position.h,
});
```

**注意**：pptxgenjs `addImage` 的 `data` 参数不需要 `data:` 前缀，直接用 `image/png;base64,...`。

### 跳过截图

```javascript
await html2pptx('slide.html', pptx, {
  captureCanvas: false,   // 跳过所有 <canvas>
  captureSvg:    false,   // 跳过所有 <svg>
});
```

---

## 7. 增强特性四：Overflow 处理模式

### 问题背景

原版在内容超出画布时直接抛出错误。实际场景中，HTML slide 的内容高度常常超出 PPT 画布（尤其是信息密集的汇报 slide），需要一种无损转化方式。

### 三种模式

```javascript
await html2pptx('slide.html', pptx, {
  overflowMode: 'expand',  // 等比缩放（推荐）
  // overflowMode: 'clip', // 截断（保留精确尺寸）
  // overflowMode: 'error',// 原有行为（默认，严格校验）
});
```

### expand 模式：等比缩放

将所有元素等比压缩到画布内，PPT 画布尺寸不变。

```
内容尺寸：10.0" × 6.63"
画布尺寸：10.0" × 5.625"

scale_y = 5.625 / 6.63 = 0.8484
scale_x = 10.0 / 10.0  = 1.0
scale   = min(0.8484, 1.0) = 0.8484

水平居中偏移（内容比画布窄时）：
offsetX = (slideW - contentW × scale) / 2
```

**受缩放影响的完整字段列表**：

| 元素类型 | 缩放字段 |
|---|---|
| 所有元素 | `position.x` `position.y` `position.w` `position.h` |
| 文字（p/h1-h6） | `style.fontSize`（最小 6pt）`style.lineSpacing` `style.paraSpaceAfter` `style.margin[]` |
| inline runs | `options.fontSize` |
| list items | `options.fontSize` `options.bullet.indent` |
| table | `colWidths[]` `rowHeights[]` 每格 `fontSize` `margin[]` |
| line | `x1` `y1` `x2` `y2` |
| Canvas/SVG 截图 | `position` 同步应用 scale + offsetX |

**scale 基准**：取 `max(body.scrollHeight, elements 最大 bottom)` / `slideH`。

不直接用 `scrollHeight` 的原因：当所有子元素都是 `position:absolute` 时，`scrollHeight` 可能为 0 或小于实际内容高度，导致 scale 计算不足，元素仍然越界。

### clip 模式：截断

坐标和字号完全不变，只根据位置决定保留、截断或丢弃。

```
规则（按优先级）：
  y ≥ slideH                → 完全超出 → 丢弃
  x ≥ slideW                → 完全超出右边界 → 丢弃
  y + h > slideH            → 跨底部边界 → 截断 h = slideH - y
  截断后 h < 0.01"          → 太小 → 丢弃
  其余                       → 完全保留
```

**表格特殊处理**：跨底部的表格不是简单截断高度，而是计算累积行高，只保留能完整放入的行（及其对应 `rowHeights` 数组）：

```
table y=5.0", rowH=0.4", slideH=5.625"
可容纳行数 = floor((5.625 - 5.0) / 0.4) = 1
只保留第 0 行，删除其余行和对应行高
```

### overflow 检测与容差

两种模式都有一个 **2% 的容差**（`OVERFLOW_TOLERANCE = 0.02`），避免浮点误差触发不必要的变换：

```
overflowH = contentH - slideH
relH      = overflowH / slideH

relH ≤ 0.02  → 不处理（在容差内）
relH > 0.02  → 触发 expand 或 clip
```

### validateDimensions 的适配

expand/clip 模式下，HTML `body` 可以不设固定 `height`（让内容自然撑开）。原版 `validateDimensions` 会把 `body.height = 0` 和 PPT layout 高度做比较，必然报错。

修复：expand/clip 模式下跳过高度维度校验，只校验宽度（宽度不匹配会导致布局完全错位，必须保留）：

```javascript
// error 模式：同时校验宽度和高度
// expand/clip 模式：只校验宽度
const hMismatch = overflowMode === 'error' && Math.abs(lh - hIn) > 0.1;
```

---

## 8. 完整 API 参考

### html2pptx(htmlFile, pres, options)

| 参数 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `htmlFile` | `string` | — | HTML 文件路径（绝对或相对） |
| `pres` | `Presentation` | — | pptxgenjs 实例 |
| `options.slideIndex` | `number` | `1` | 当前 slide 在 PPTX 中的序号（1-based），用于渐变定位 |
| `options.registry` | `GradientRegistry` | `null` | 渐变注册表，多 slide 共享同一实例 |
| `options.slide` | `Slide` | `null` | 复用已有 slide 对象（不传则自动 addSlide） |
| `options.tmpDir` | `string` | `/tmp` | 临时文件目录 |
| `options.overflowMode` | `'expand'\|'clip'\|'error'` | `'error'` | overflow 处理模式 |
| `options.captureCanvas` | `boolean` | `true` | 是否截图 `<canvas>` 元素 |
| `options.captureSvg` | `boolean` | `true` | 是否截图 `<svg>` 元素 |
| `options.screenshotScale` | `number` | `2` | 截图渲染倍数（1 或 2） |
| `options.screenshotTimeout` | `number` | `8000` | 单个图表等待渲染超时（ms） |

**返回值**：

```typescript
{
  slide:        Slide,           // pptxgenjs Slide 对象
  placeholders: Array<{id, x, y, w, h}>,   // placeholder 坐标
  screenshots:  Array<{
    tag:        'CANVAS' | 'SVG',
    objectName: string,
    position:   {x, y, w, h},   // 英寸
    pngBuffer:  Buffer | null,
    dataUrl:    string | null,
    error?:     string,
  }>,
  overflow: {
    mode:          string,
    didTransform:  boolean,
    transformInfo: string,       // 人类可读的变换描述
    scale:         number,       // expand 模式的缩放比例
    offsetX:       number,       // expand 模式的水平居中偏移（英寸）
  }
}
```

### postprocess(pptxPath, registry)

```javascript
const { processed, warnings } = await postprocess('output.pptx', registry);
// processed — 成功替换的渐变数量
// warnings  — 未能处理的渐变警告列表
```

### GradientRegistry

```javascript
const registry = new GradientRegistry();
registry.isEmpty()                          // → boolean
registry.registerBackground(slideIndex, gradient)
registry.registerShape(slideIndex, objectName, gradient)
registry.hasSlide(slideIndex)              // → boolean
registry.getSlide(slideIndex)              // → { background, shapes: Map }
registry.slideIndices()                    // → number[]
```

### handleOverflow(slideData, bodyDimensions, presLayout, mode)

```javascript
const { handleOverflow } = require('./overflow-handler');
const result = handleOverflow(slideData, bodyDimensions, pres.presLayout, 'expand');
// result: { slideData, mode, didTransform, transformInfo, scale, offsetX }
```

---

## 9. 测试体系

### 测试文件及覆盖

| 文件 | 类型 | 用例数 | 覆盖内容 |
|---|---|---|---|
| `test-gradient.js` | 纯逻辑单元测试 | 44 | 颜色解析、角度转换、linear/radial/multi-stop、XML 生成、GradientRegistry、XML 后处理替换逻辑 |
| `test-screenshot-logic.js` | 纯逻辑单元测试 | 17 | 元素过滤、坐标转换、截图区域计算、等待策略解析、dataUrl 格式、坐标边界验证 |
| `test-overflow.js` | 单元 + 集成 | 22（+4 集成） | overflow 阈值检测、expand scale 计算、expand 样式缩放、clip 截断规则、表格截断、三种模式入口 |
| `test-e2e-charts.js` | 端到端集成测试 | 4 个图表 × 多断言 | Canvas 柱状图、SVG 折线图、模拟 ECharts（400ms 延迟）、SVG 饼图，验证 PPTX 媒体文件数量和 blip 计数 |

### 运行方式

```bash
# 纯逻辑单元测试（无需浏览器，快速）
node test-gradient.js
node test-screenshot-logic.js
node test-overflow.js

# 集成测试（需要配置 Playwright 浏览器）
export PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH=/path/to/chrome

node test-overflow.js          # 同时跑单元 + 集成
node test-e2e-charts.js        # 图表截图端到端测试
```

### 测试设计原则

**纯函数测试不依赖浏览器**：渐变解析、overflow 缩放计算、XML 生成、截图参数计算等所有纯逻辑函数，在 Node.js 环境直接运行，不启动 Playwright，秒级完成。

**浏览器集成测试验证端到端**：集成测试生成真实 PPTX，用 JSZip 解压验证 XML 内容（`gradFill` 存在、坐标 EMU 值在画布范围内、媒体文件数量正确），而不是只检查函数返回值。

---

## 10. 环境配置与运行方式

### Node.js 版本

Node.js ≥ 18（需要 `Array.prototype.at`、`structuredClone` 等现代 API）

### 环境变量

| 变量 | 说明 | 示例 |
|---|---|---|
| `PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH` | 指定 Chrome/Chromium 可执行文件路径 | `/opt/pw-browsers/chromium-1194/chrome-linux/chrome` |

macOS 自动使用系统安装的 Chrome，无需设置。

Linux 容器/服务器环境推荐用系统包安装 Chromium，或通过 `npx playwright install chromium` 下载（需要网络访问 playwright.dev）。

### 启动参数

`html2pptx.js` 默认传给 Playwright 的 Chrome 参数：
```
--no-sandbox
--disable-setuid-sandbox
--disable-dev-shm-usage
```

这三个参数是 Linux Docker 环境的标配，macOS 不需要（会走系统 Chrome 通道）。

### HTML 文件要求

**固定宽度**（必须）：
```css
body {
  width: 720pt;   /* 对应 16:9 幻灯片 10" × 72pt/in */
  /* height 在 expand/clip 模式下可省略 */
}
```

**高度设置**：
- `error` 模式（默认）：必须设 `height: 405pt`，严格匹配画布
- `expand` / `clip` 模式：可省略 `height`，让内容自然撑开

**单位换算**：
```
16:9 幻灯片 = 10" × 5.625"
           = 720pt × 405pt
           = 960px × 540px（96 DPI）
```

**CSS 渐变语法**：使用标准 CSS 语法，浏览器会在 `getComputedStyle` 中转换为标准格式（如将 `to right` 转为 `90deg`），解析器对两种格式均支持。

**图表 ready 标志**（可选，提高截图精度）：
```html
<canvas data-h2p-wait="flag" id="chart"></canvas>
<script>
  // 在图表完成渲染后设置：
  window._h2p_ready = true;
</script>
```

---

## 11. 已知限制与降级规则

### 渐变

| 限制 | 原因 | 处理 |
|---|---|---|
| `conic-gradient` 不支持 | PPT OOXML 无对应元素 | 取第一个 stop 纯色 + warning |
| 多层渐变叠加 | PPT shape 只有一个 fill | 取第一层渐变 |
| `radial-gradient` 形状近似 | PPT 只有 circle/rect path，无 ellipse | 以 circle 近似 |
| 渐变在 `<p>/<span>` 上 | text 标签不支持渐变 | 需改用 `<div>` 包裹 |

### 表格

| 限制 | 说明 |
|---|---|
| 复杂合并单元格（colspan × rowspan 嵌套） | 支持，用 occupiedMap 追踪 |
| 表格内嵌图片 | 不支持，图片在 td 内会被忽略 |
| `border-radius` on cell | 忽略（PPT 表格单元格无圆角） |
| 嵌套表格 | 不支持（外层表格提取时会把内层一并标记为已处理） |

### 截图

| 限制 | 说明 |
|---|---|
| WebGL canvas | `getImageData` 无法读取，自动 fallback 到双 rAF 等待 |
| 跨域资源 | `file://` 协议下 canvas 可能因 CORS 无法读取像素，fallback 正常 |
| 动态数据更新的图表 | 只截当前帧；若需截特定状态，配合 `data-h2p-wait="flag"` 精确控制 |
| 超大 SVG（> 20MB） | sharp 处理可能较慢，增加 `screenshotTimeout` |

### overflow

| 限制 | 说明 |
|---|---|
| expand 最小字号 6pt | 防止极端缩放（scale < 0.1）时文字完全不可读 |
| clip 不保证视觉完整 | 跨边界的色块会被截断，可能出现半截内容 |
| 水平 overflow | 当前版本以高度 overflow 为主要场景；宽度 overflow 时两个方向同时缩放 |

---

## 附录：修改文件清单

### 新增文件（5 个）

| 文件 | 行数 | 功能 |
|---|---|---|
| `overflow-handler.js` | 432 | overflow 检测、expand 等比缩放、clip 截断，纯函数，无副作用 |
| `element-screenshot.js` | 356 | canvas/svg 元素收集、三级等待策略、2× 截图 + sharp 降采样 |
| `gradient-parser.js` | 332 | CSS 渐变字符串解析、角度转换、stop 提取 |
| `gradient-xml.js` | 98 | OOXML gradFill XML 生成（linear / radial） |
| `gradient-postprocess.js` | 290 | GradientRegistry 类、JSZip 解压重打包、solidFill→gradFill 替换 |

### 修改文件（1 个）

`html2pptx.js`（原版约 400 行 → 增强版 986 行）

主要变更：
- **require 区**：新增 4 个模块引用
- **`validateDimensions`**：增加 `overflowMode` 参数，expand/clip 模式跳过高度校验
- **`extractSlideData`（浏览器侧）**：
  - 新增 `extractTable()` — 表格完整提取逻辑
  - 新增渐变检测和 `gradientShapes` 收集
  - `backgroundGradientCss` 字段传回原始渐变 CSS 字符串
  - 修复 SVG 元素的 `className.baseVal` 兼容（原版 `className.includes` 对 SVG 报错）
- **`addElements`（Node.js 侧）**：新增 `table` 分支调用 `addTable()`
- **`addTable`**：新增函数，处理 colspan/rowspan 占位格和所有单元格样式映射
- **主函数 `html2pptx`**：
  - 新增 `captureCanvas` `captureSvg` `screenshotScale` `screenshotTimeout` `overflowMode` 参数
  - 浏览器会话内调用 `captureElements()` 截图
  - 校验阶段按模式过滤 overflow 错误
  - 新增 `handleOverflow()` 调用
  - 截图嵌入时同步应用 overflow 的 scale/offsetX
  - 返回值增加 `screenshots` 和 `overflow` 字段

---

*文档基于 2026-03-18 代码版本*
