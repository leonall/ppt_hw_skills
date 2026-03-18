'use strict';

/**
 * overflow-handler.js
 *
 * 处理 HTML 内容超出 PPT 画布边界的两种模式：
 *
 * ─────────────────────────────────────────────────────
 * EXPAND（展开）模式
 * ─────────────────────────────────────────────────────
 *   当 HTML 实际渲染高度 > PPT 画布高度时，对所有元素坐标和尺寸
 *   等比缩放（scale ≤ 1），使内容整体压缩到画布内。
 *
 *   受影响的字段：
 *     元素 position:  x / y / w / h  × scale
 *     字体:           fontSize        × scale（向下取整，最小 6pt）
 *     行高:           lineSpacing     × scale
 *     表格列宽:       colWidths[]     × scale
 *     表格行高:       rowHeights[]    × scale
 *     文字内边距:     margin[]        × scale
 *     图片坐标:       同 position
 *
 *   水平方向：若缩放后水平方向有空白，元素整体居中（offsetX）。
 *
 * ─────────────────────────────────────────────────────
 * CLIP（截取）模式
 * ─────────────────────────────────────────────────────
 *   只保留完全或部分在画布内的元素，超出底部的元素按以下规则处理：
 *
 *   - 元素顶部 y ≥ slideH            → 丢弃
 *   - 元素底部 y + h > slideH        → 截断 h 为 slideH - y
 *   - 元素完全在画布内               → 保留不变
 *
 *   字体/列宽等不做变换，坐标不变，只改 h。
 *   超出右边界（x ≥ slideW）的元素同样丢弃。
 *
 * ─────────────────────────────────────────────────────
 * 不需要处理时（内容未 overflow）
 * ─────────────────────────────────────────────────────
 *   两种模式均原样返回，零开销。
 */

// ─────────────────────────────────────────────────────
// 常量
// ─────────────────────────────────────────────────────

const SLIDE_W_IN = 10.0;    // 16:9 幻灯片宽度（英寸）
const SLIDE_H_IN = 5.625;   // 16:9 幻灯片高度（英寸）
const MIN_FONT_PT = 6;       // expand 模式最小字号（pt）
const OVERFLOW_TOLERANCE = 0.02; // 2% 以内不做处理（浮点误差容差）

// ─────────────────────────────────────────────────────
// 工具
// ─────────────────────────────────────────────────────

/** 按 scale 缩放 position 对象 { x, y, w, h } */
function scalePos(pos, scale, offsetX = 0) {
  return {
    x: pos.x * scale + offsetX,
    y: pos.y * scale,
    w: pos.w * scale,
    h: pos.h * scale,
  };
}

/** 深拷贝 slideData（避免修改原始数据） */
function cloneSlideData(slideData) {
  return JSON.parse(JSON.stringify(slideData));
}

/**
 * 计算实际内容边界
 * 遍历所有元素，取最大的 y + h（底部） 和 x + w（右边）
 */
function getContentBounds(elements) {
  let maxRight = 0, maxBottom = 0;
  for (const el of elements) {
    if (!el.position) continue;
    const right  = el.position.x + el.position.w;
    const bottom = el.position.y + el.position.h;
    if (right  > maxRight)  maxRight  = right;
    if (bottom > maxBottom) maxBottom = bottom;
  }
  return { maxRight, maxBottom };
}

// ─────────────────────────────────────────────────────
// EXPAND 模式：等比缩放所有元素
// ─────────────────────────────────────────────────────

/**
 * 对单个元素应用 expand 缩放
 * @param {object} el     - slideData.elements 中的单个元素
 * @param {number} scale  - 缩放比例（0 < scale ≤ 1）
 * @param {number} offsetX - 水平居中偏移（英寸）
 * @returns {object} 缩放后的新元素（不修改原对象）
 */
function applyExpandToElement(el, scale, offsetX) {
  const out = { ...el };

  // position：所有元素都有
  if (out.position) {
    out.position = scalePos(out.position, scale, offsetX);
  }

  // 文字相关样式
  if (out.style) {
    out.style = { ...out.style };
    if (out.style.fontSize != null) {
      out.style.fontSize = Math.max(MIN_FONT_PT, out.style.fontSize * scale);
    }
    if (out.style.lineSpacing != null) {
      out.style.lineSpacing = out.style.lineSpacing * scale;
    }
    if (out.style.paraSpaceBefore != null) {
      out.style.paraSpaceBefore = out.style.paraSpaceBefore * scale;
    }
    if (out.style.paraSpaceAfter != null) {
      out.style.paraSpaceAfter = out.style.paraSpaceAfter * scale;
    }
    // margin: [top, right, bottom, left] 单位英寸
    if (Array.isArray(out.style.margin)) {
      out.style.margin = out.style.margin.map(m => m * scale);
    }
  }

  // inline 文字 runs 中的 fontSize
  if (Array.isArray(out.text)) {
    out.text = out.text.map(run => {
      if (!run.options?.fontSize) return run;
      return {
        ...run,
        options: {
          ...run.options,
          fontSize: Math.max(MIN_FONT_PT, run.options.fontSize * scale),
        },
      };
    });
  }

  // list items 中的 fontSize 和 bullet indent
  if (Array.isArray(out.items)) {
    out.items = out.items.map(item => {
      if (!item.options) return item;
      const opts = { ...item.options };
      if (opts.fontSize != null) {
        opts.fontSize = Math.max(MIN_FONT_PT, opts.fontSize * scale);
      }
      if (opts.bullet?.indent != null) {
        opts.bullet = { ...opts.bullet, indent: opts.bullet.indent * scale };
      }
      return { ...item, options: opts };
    });
  }

  // table：colWidths / rowHeights / 单元格 fontSize / margin
  if (out.type === 'table') {
    if (Array.isArray(out.colWidths)) {
      out.colWidths = out.colWidths.map(w => w * scale);
    }
    if (Array.isArray(out.rowHeights)) {
      out.rowHeights = out.rowHeights.map(h => h * scale);
    }
    if (Array.isArray(out.rows)) {
      out.rows = out.rows.map(row =>
        row.map(cell => {
          if (cell._spanPlaceholder) return cell;
          const c = { ...cell };
          if (c.fontSize != null) {
            c.fontSize = Math.max(MIN_FONT_PT, c.fontSize * scale);
          }
          if (Array.isArray(c.margin)) {
            c.margin = c.margin.map(m => m * scale);
          }
          return c;
        })
      );
    }
  }

  // line 元素：x1/y1/x2/y2
  if (out.type === 'line') {
    out.x1 = out.x1 * scale + offsetX;
    out.y1 = out.y1 * scale;
    out.x2 = out.x2 * scale + offsetX;
    out.y2 = out.y2 * scale;
  }

  return out;
}

/**
 * EXPAND 模式主函数
 * @param {object} slideData      - extractSlideData 的输出
 * @param {object} bodyDimensions - { width, height } px
 * @param {object} presLayout     - pptxgenjs presLayout { width, height } EMU
 * @returns {{ slideData, scale, offsetX, didScale, info }}
 */
function applyExpand(slideData, bodyDimensions, presLayout) {
  const EMU_PER_IN = 914400;
  const PX_PER_IN = 96;

  // 幻灯片画布尺寸（英寸）
  const slideW = presLayout.width  / EMU_PER_IN;
  const slideH = presLayout.height / EMU_PER_IN;

  // HTML 内容实际尺寸（英寸）
  // scrollHeight 是浏览器报告值，但 position:absolute 的元素可能超出 scrollHeight
  // 同时取 elements 的实际最大 bottom，确保所有元素都能被压入画布
  const scrollW = bodyDimensions.scrollWidth  / PX_PER_IN;
  const scrollH = bodyDimensions.scrollHeight / PX_PER_IN;
  const { maxRight, maxBottom } = getContentBounds(slideData.elements);
  const contentW = Math.max(scrollW, maxRight);
  const contentH = Math.max(scrollH, maxBottom);

  // 计算 overflow 量
  const overflowH = contentH - slideH;
  const overflowW = contentW - slideW;

  // 没有超出，直接返回
  const relH = overflowH / slideH;
  const relW = overflowW / slideW;
  if (relH <= OVERFLOW_TOLERANCE && relW <= OVERFLOW_TOLERANCE) {
    return { slideData, scale: 1, offsetX: 0, didScale: false, info: 'no overflow' };
  }

  // 等比缩放：取 X 和 Y 两个方向缩放比中较小的
  const scaleX = overflowW > OVERFLOW_TOLERANCE * slideW ? slideW / contentW : 1;
  const scaleY = overflowH > OVERFLOW_TOLERANCE * slideH ? slideH / contentH : 1;
  const scale  = Math.min(scaleX, scaleY);

  // 若只是高度方向溢出，横向有空余 → 居中
  const scaledW  = contentW * scale;
  const offsetX  = Math.max(0, (slideW - scaledW) / 2);

  const newSlideData = cloneSlideData(slideData);
  newSlideData.elements = newSlideData.elements.map(el =>
    applyExpandToElement(el, scale, offsetX)
  );

  return {
    slideData: newSlideData,
    scale,
    offsetX,
    didScale: true,
    info: `expand: scale=${scale.toFixed(4)}, offsetX=${offsetX.toFixed(3)}in` +
          `, content ${contentW.toFixed(2)}x${contentH.toFixed(2)}in` +
          ` → ${(contentW*scale).toFixed(2)}x${(contentH*scale).toFixed(2)}in`,
  };
}

// ─────────────────────────────────────────────────────
// CLIP 模式：截断超出边界的元素
// ─────────────────────────────────────────────────────

/**
 * 对单个元素应用 clip 截断
 * @param {object} el
 * @param {number} slideW - 幻灯片宽度（英寸）
 * @param {number} slideH - 幻灯片高度（英寸）
 * @returns {object|null} null 表示元素被完全裁掉
 */
function applyClipToElement(el, slideW, slideH) {
  // line 元素单独处理
  if (el.type === 'line') {
    // 如果两端都在画布外（y1 和 y2 均 >= slideH），丢弃
    if (el.y1 >= slideH && el.y2 >= slideH) return null;
    // 简单截断：确保不超出（精确裁线比较复杂，先保留）
    return el;
  }

  if (!el.position) return el; // 没有 position 的元素不处理

  const { x, y, w, h } = el.position;

  // 完全超出右边界 or 底部边界 → 丢弃
  if (x >= slideW) return null;
  if (y >= slideH) return null;

  // 完全在画布内 → 保留
  if (x + w <= slideW && y + h <= slideH) return el;

  // 跨边界 → 截断
  const out = { ...el };
  out.position = { ...el.position };

  // 截断右边界
  if (x + w > slideW) {
    out.position.w = Math.max(0, slideW - x);
  }

  // 截断底部边界
  if (y + h > slideH) {
    out.position.h = Math.max(0, slideH - y);
  }

  // 截断后尺寸过小（< 0.01 inch）→ 丢弃
  if (out.position.w < 0.01 || out.position.h < 0.01) return null;

  // 表格行高需要同步截断：只保留能完整放入的行
  if (out.type === 'table' && Array.isArray(out.rowHeights)) {
    out.rowHeights = [...el.rowHeights];
    out.rows = [...el.rows];
    let cumH = y;
    let keepRows = 0;
    for (let i = 0; i < out.rowHeights.length; i++) {
      cumH += out.rowHeights[i];
      if (cumH > slideH) break;
      keepRows++;
    }
    out.rows = out.rows.slice(0, keepRows);
    out.rowHeights = out.rowHeights.slice(0, keepRows);
    if (out.rows.length === 0) return null;
  }

  return out;
}

/**
 * CLIP 模式主函数
 * @param {object} slideData
 * @param {object} presLayout - { width, height } EMU
 * @returns {{ slideData, dropped, clipped, info }}
 */
function applyClip(slideData, presLayout) {
  const EMU_PER_IN = 914400;
  const slideW = presLayout.width  / EMU_PER_IN;
  const slideH = presLayout.height / EMU_PER_IN;

  let dropped = 0, clipped = 0;
  const newSlideData = cloneSlideData(slideData);

  newSlideData.elements = newSlideData.elements.reduce((acc, el) => {
    const result = applyClipToElement(el, slideW, slideH);
    if (result === null) {
      dropped++;
      return acc;
    }
    // 检查是否被截断
    if (el.position && result.position) {
      const hChanged = Math.abs(result.position.h - el.position.h) > 0.001;
      const wChanged = Math.abs(result.position.w - el.position.w) > 0.001;
      if (hChanged || wChanged) clipped++;
    }
    acc.push(result);
    return acc;
  }, []);

  return {
    slideData: newSlideData,
    dropped,
    clipped,
    info: `clip: ${dropped} dropped, ${clipped} clipped`,
  };
}

// ─────────────────────────────────────────────────────
// 主入口
// ─────────────────────────────────────────────────────

/**
 * 根据模式处理 overflow
 *
 * @param {object} slideData      - extractSlideData 输出
 * @param {object} bodyDimensions - { width, height, scrollWidth, scrollHeight } px
 * @param {object} presLayout     - pptxgenjs presLayout { width, height } EMU
 * @param {'expand'|'clip'|'error'} mode
 *   'expand'  展开模式：等比缩放所有元素
 *   'clip'    截取模式：截断超出画布的内容
 *   'error'   原有行为：overflow 时抛出错误（默认）
 *
 * @returns {{ slideData, mode, didTransform, transformInfo, scale, offsetX }}
 */
function handleOverflow(slideData, bodyDimensions, presLayout, mode) {
  const EMU_PER_IN = 914400;
  const PX_PER_IN = 96;
  const slideH = presLayout.height / EMU_PER_IN;
  const slideW = presLayout.width  / EMU_PER_IN;
  const contentH = bodyDimensions.scrollHeight / PX_PER_IN;
  const contentW = bodyDimensions.scrollWidth  / PX_PER_IN;

  const overflowH = (contentH - slideH) / slideH;
  const overflowW = (contentW - slideW) / slideW;
  const hasOverflow = overflowH > OVERFLOW_TOLERANCE || overflowW > OVERFLOW_TOLERANCE;

  if (!hasOverflow) {
    return {
      slideData,
      mode,
      didTransform: false,
      transformInfo: 'no overflow',
      scale: 1,
      offsetX: 0,
    };
  }

  if (mode === 'expand') {
    const result = applyExpand(slideData, bodyDimensions, presLayout);
    return {
      slideData:     result.slideData,
      mode:          'expand',
      didTransform:  result.didScale,
      transformInfo: result.info,
      scale:         result.scale,
      offsetX:       result.offsetX,
    };
  }

  if (mode === 'clip') {
    const result = applyClip(slideData, presLayout);
    return {
      slideData:     result.slideData,
      mode:          'clip',
      didTransform:  result.dropped > 0 || result.clipped > 0,
      transformInfo: result.info,
      scale:         1,
      offsetX:       0,
    };
  }

  // 'error' 模式：调用方会从 bodyDimensions.errors 读取错误
  return {
    slideData,
    mode:          'error',
    didTransform:  false,
    transformInfo: 'overflow not handled (mode=error)',
    scale:         1,
    offsetX:       0,
  };
}

module.exports = { handleOverflow, applyExpand, applyClip, getContentBounds, OVERFLOW_TOLERANCE };
