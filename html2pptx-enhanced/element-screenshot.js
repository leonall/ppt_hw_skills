'use strict';

/**
 * element-screenshot.js
 *
 * 把 HTML 中的 Canvas（图表）和 SVG 元素截图为 PNG Buffer，
 * 用于嵌入 PowerPoint 幻灯片。
 *
 * 核心设计：
 *   - 所有截图都在调用者已打开的 Playwright page 上进行（复用浏览器）
 *   - 用 2× deviceScaleFactor 渲染，再用 sharp 降采样，得到高清图像
 *   - 图表等待：先等 window._h2p_ready 标志，再等 requestAnimationFrame 稳定
 *   - 返回 { dataUrl, widthInch, heightInch, x, y }，调用方直接 addImage
 */

const sharp = require('sharp');

// ─────────────────────────────────────────────────────
// 常量
// ─────────────────────────────────────────────────────

const PX_PER_IN = 96;

// 2x 渲染缩放倍数：截图用 2x，再降采样到 1x，清晰度翻倍
const RENDER_SCALE = 2;

// 默认等待超时（ms）
const DEFAULT_TIMEOUT = 8000;

// ─────────────────────────────────────────────────────
// 等待图表渲染完成
// ─────────────────────────────────────────────────────

/**
 * 等待策略（按优先级依次尝试）：
 *
 * 1. window._h2p_ready === true
 *    HTML 作者在图表初始化完成后设置此标志，最精确
 *    示例：chart.on('finished', () => window._h2p_ready = true)
 *
 * 2. window._h2p_ready 是 Promise
 *    await window._h2p_ready  （适合异步初始化）
 *
 * 3. 自动检测：canvas 元素上有像素内容（非全黑/全白）
 *    轮询 canvas.toDataURL() 采样点变化
 *
 * 4. 固定延迟 fallback（默认 1500ms）
 *    如以上均不适用
 */
async function waitForRender(page, elementInfo, timeout) {
  const { tag, selector } = elementInfo;
  const start = Date.now();

  // 策略1&2：window._h2p_ready
  try {
    const hasReadyFlag = await page.evaluate(() =>
      typeof window._h2p_ready !== 'undefined'
    );
    if (hasReadyFlag) {
      await page.waitForFunction(
        async () => {
          const r = window._h2p_ready;
          if (r === true) return true;
          if (r && typeof r.then === 'function') {
            await r;
            return true;
          }
          return false;
        },
        { timeout }
      );
      return;
    }
  } catch (e) {
    // 超时或出错，继续下一策略
  }

  // 策略3：Canvas 像素稳定检测
  if (tag === 'CANVAS') {
    try {
      await page.waitForFunction(
        (sel) => {
          const canvas = document.querySelector(sel);
          if (!canvas || !canvas.getContext) return false;
          try {
            const ctx = canvas.getContext('2d');
            if (!ctx) return false;
            // 采样 5x5 中心区域的像素，判断是否有非空内容
            const w = canvas.width, h = canvas.height;
            const data = ctx.getImageData(w/2-2, h/2-2, 5, 5).data;
            // 至少有一个非透明像素
            for (let i = 3; i < data.length; i += 4) {
              if (data[i] > 0) return true;
            }
            return false;
          } catch (e) {
            return true; // WebGL canvas 无法读取，假设已就绪
          }
        },
        selector,
        { timeout: Math.min(timeout, 4000) }
      );
      return;
    } catch (e) {
      // fallthrough
    }
  }

  // 策略4：等待 2 个 animation frame 确保绘制完成
  await page.evaluate(() => new Promise(resolve => {
    requestAnimationFrame(() => requestAnimationFrame(resolve));
  }));

  // 再加一点固定延迟（应对 ECharts 的动画初始阶段）
  const elapsed = Date.now() - start;
  const remainWait = Math.min(1500, timeout / 2) - elapsed;
  if (remainWait > 0) {
    await new Promise(r => setTimeout(r, remainWait));
  }
}

// ─────────────────────────────────────────────────────
// 单个元素截图
// ─────────────────────────────────────────────────────

/**
 * 对 page 上的一个元素截图，返回高清 PNG Buffer
 * @param {object} page      - Playwright page（已 goto 目标 HTML）
 * @param {object} info      - 元素信息 { selector, rect: {x,y,width,height}, tag }
 * @param {object} [opts]
 * @param {number} [opts.scale=2]         - 渲染倍数（1 或 2）
 * @param {number} [opts.timeout=8000]    - 等待超时 ms
 * @param {number} [opts.padding=0]       - 截图外扩 px（防止边缘裁切）
 * @returns {Promise<Buffer>} PNG buffer（已降采样到 1x）
 */
async function screenshotElement(page, info, opts = {}) {
  const {
    scale = RENDER_SCALE,
    timeout = DEFAULT_TIMEOUT,
    padding = 2,
  } = opts;

  // 1. 等待渲染
  await waitForRender(page, info, timeout);

  // 2. 计算截图区域（加 padding 防裁切）
  const { x, y, width, height } = info.rect;
  const clip = {
    x: Math.max(0, x - padding),
    y: Math.max(0, y - padding),
    width:  width  + padding * 2,
    height: height + padding * 2,
  };

  // 3. 截图
  // 注意：page 已经在 2x deviceScaleFactor context 下创建（见调用方）
  // 所以 clip 坐标是 CSS 坐标，截图像素是 clip * scale
  const rawBuf = await page.screenshot({
    type: 'png',
    clip,
  });

  // 4. 用 sharp 裁掉 padding 并降采样回 1x
  const targetW = Math.round(width);
  const targetH = Math.round(height);

  const outBuf = await sharp(rawBuf)
    // 从 scaled 图上裁掉 padding（padding * scale 像素）
    .extract({
      left:   Math.round(padding * scale),
      top:    Math.round(padding * scale),
      width:  Math.round(width  * scale),
      height: Math.round(height * scale),
    })
    // 降采样到目标尺寸
    .resize(targetW, targetH, { kernel: 'lanczos3' })
    .png({ compressionLevel: 6 })
    .toBuffer();

  return outBuf;
}

// ─────────────────────────────────────────────────────
// 浏览器侧：提取页面中所有需要截图的元素
// ─────────────────────────────────────────────────────

/**
 * 在 page.evaluate 中运行，收集所有 canvas 和 svg 元素的位置信息
 * 返回给 Node.js 侧处理
 */
const BROWSER_COLLECT_SCRIPT = () => {
  const PX_PER_IN = 96;
  const pxToInch = px => px / PX_PER_IN;

  const results = [];
  let idx = 0;

  // 为每个元素生成唯一选择器
  const getSelector = (el) => {
    if (el.id) return `#${CSS.escape(el.id)}`;
    // 用 nth-of-type 定位
    const parent = el.parentElement;
    if (!parent) return el.tagName.toLowerCase();
    const siblings = Array.from(parent.children).filter(c => c.tagName === el.tagName);
    const nth = siblings.indexOf(el) + 1;
    const parentSel = parent.id ? `#${CSS.escape(parent.id)}` : parent.tagName.toLowerCase();
    return `${parentSel} > ${el.tagName.toLowerCase()}:nth-of-type(${nth})`;
  };

  // 收集 CANVAS 元素
  document.querySelectorAll('canvas').forEach(el => {
    const rect = el.getBoundingClientRect();
    if (rect.width < 10 || rect.height < 10) return; // 跳过空元素

    // 尝试从 data-h2p-* 属性获取覆盖配置
    const waitMode = el.dataset.h2pWait || 'auto';      // 'auto' | 'flag' | 'delay:Nms'
    const customDelay = parseInt(el.dataset.h2pDelay || '0');

    results.push({
      tag: 'CANVAS',
      idx: idx++,
      objectName: el.id || `canvas-${idx}`,
      selector: getSelector(el),
      rect: {
        x: rect.left,
        y: rect.top,
        width: rect.width,
        height: rect.height,
      },
      position: {  // 英寸，用于 pptxgenjs addImage
        x: pxToInch(rect.left),
        y: pxToInch(rect.top),
        w: pxToInch(rect.width),
        h: pxToInch(rect.height),
      },
      waitMode,
      customDelay,
    });
  });

  // 收集独立 SVG 元素（不在 canvas 内，不是 img 内嵌的）
  document.querySelectorAll('svg').forEach(el => {
    // 跳过：在 canvas 内 / 在 button 内 / display:none / 尺寸为 0
    if (el.closest('canvas') || el.closest('button')) return;
    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return;
    const rect = el.getBoundingClientRect();
    if (rect.width < 10 || rect.height < 10) return;

    results.push({
      tag: 'SVG',
      idx: idx++,
      objectName: el.id || `svg-${idx}`,
      selector: getSelector(el),
      rect: {
        x: rect.left,
        y: rect.top,
        width: rect.width,
        height: rect.height,
      },
      position: {
        x: pxToInch(rect.left),
        y: pxToInch(rect.top),
        w: pxToInch(rect.width),
        h: pxToInch(rect.height),
      },
      waitMode: el.dataset.h2pWait || 'auto',
      customDelay: parseInt(el.dataset.h2pDelay || '0'),
    });
  });

  return results;
};

// ─────────────────────────────────────────────────────
// 主函数：对一个已加载的 page 截取所有图表和 SVG
// ─────────────────────────────────────────────────────

/**
 * @param {object} page        - Playwright page（已 goto，已 setViewport）
 * @param {object} [opts]
 * @param {number} [opts.scale=2]      - 截图倍数
 * @param {number} [opts.timeout=8000] - 单元素等待超时
 * @param {Set}    [opts.skipTags]     - 要跳过的标签集合，如 new Set(['SVG'])
 * @returns {Promise<Array>} 每个元素的截图结果
 *   [ { tag, objectName, selector, position, pngBuffer, dataUrl } ]
 */
async function captureElements(page, opts = {}) {
  const {
    scale   = RENDER_SCALE,
    timeout = DEFAULT_TIMEOUT,
    skipTags = new Set(),
  } = opts;

  // 1. 收集元素信息（在原始 1x page 上，坐标准确）
  const elements = await page.evaluate(BROWSER_COLLECT_SCRIPT);

  const toCapture = elements.filter(el => !skipTags.has(el.tag));
  if (toCapture.length === 0) return [];

  // 2. 创建 2x context，重新加载同一 URL
  const url = page.url();
  const viewportW = Math.ceil(await page.evaluate(() => document.body.scrollWidth));
  const viewportH = Math.ceil(await page.evaluate(() => document.body.scrollHeight));

  const browser = page.context().browser();
  const ctx2x = await browser.newContext({ deviceScaleFactor: scale });
  const page2x = await ctx2x.newPage();
  await page2x.setViewportSize({ width: viewportW, height: viewportH });
  await page2x.goto(url);

  // 3. 等待页面整体就绪
  try {
    await page2x.waitForLoadState('networkidle', { timeout: 3000 });
  } catch (_) {}
  // 额外等待两帧，确保 canvas 初始绘制完成
  await page2x.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));

  // 4. 逐元素截图
  const results = [];
  for (const info of toCapture) {
    try {
      if (info.customDelay > 0) {
        await new Promise(r => setTimeout(r, info.customDelay));
      }

      const pngBuffer = await screenshotElement(page2x, info, { scale, timeout });
      const dataUrl = 'image/png;base64,' + pngBuffer.toString('base64');

      results.push({
        tag:        info.tag,
        objectName: info.objectName,
        selector:   info.selector,
        position:   info.position,
        pngBuffer,
        dataUrl,
      });
    } catch (e) {
      console.warn(`[element-screenshot] ${info.tag} "${info.objectName}": ${e.message}`);
      results.push({
        tag:        info.tag,
        objectName: info.objectName,
        selector:   info.selector,
        position:   info.position,
        pngBuffer:  null,
        dataUrl:    null,
        error:      e.message,
      });
    }
  }

  await ctx2x.close();
  return results;
}

module.exports = { captureElements, screenshotElement, BROWSER_COLLECT_SCRIPT };
