'use strict';

/**
 * gradient-parser.js
 *
 * 把浏览器 getComputedStyle 返回的渐变字符串解析为结构化对象，
 * 再转换为 PPT OOXML gradFill 所需的格式。
 *
 * 支持：
 *   linear-gradient(...)
 *   radial-gradient(...)
 *   repeating-linear-gradient(...)  → 当作 linear 处理
 *   repeating-radial-gradient(...)  → 当作 radial 处理
 *   多层渐变叠加                    → 取第一个渐变层
 *   conic-gradient(...)             → 降级为纯色
 *
 * 浏览器 computed style 的渐变格式（与 authored CSS 不同）：
 *   linear-gradient(135deg, rgb(255, 59, 79) 0%, rgb(192, 32, 48) 100%)
 *   radial-gradient(circle at 50% 50%, rgb(255, 255, 255) 0%, rgb(0, 0, 100) 100%)
 *   to right / to bottom 关键字已被浏览器转换为角度值
 */

// ─────────────────────────────────────────────────────
// 工具函数
// ─────────────────────────────────────────────────────

/**
 * rgb/rgba 字符串 → { hex: 'RRGGBB', alpha: 0-100000 }
 * alpha 遵循 OOXML 规范：100000 = 完全不透明，0 = 完全透明
 */
function parseColor(colorStr) {
  if (!colorStr) return null;
  colorStr = colorStr.trim();

  const rgbaMatch = colorStr.match(
    /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/
  );
  if (!rgbaMatch) return null;

  const r = parseInt(rgbaMatch[1]).toString(16).padStart(2, '0').toUpperCase();
  const g = parseInt(rgbaMatch[2]).toString(16).padStart(2, '0').toUpperCase();
  const b = parseInt(rgbaMatch[3]).toString(16).padStart(2, '0').toUpperCase();
  const hex = r + g + b;

  // CSS alpha 0-1 → OOXML alpha 0-100000
  const cssAlpha = rgbaMatch[4] !== undefined ? parseFloat(rgbaMatch[4]) : 1;
  const alpha = Math.round(cssAlpha * 100000);

  return { hex, alpha };
}

/**
 * 解析位置值：'0%' → 0, '50%' → 50000, '100%' → 100000
 * OOXML gs pos 范围 0-100000
 */
function parsePosition(posStr) {
  if (!posStr) return null;
  posStr = posStr.trim();
  if (posStr.endsWith('%')) {
    return Math.round(parseFloat(posStr) * 1000);
  }
  return null;
}

/**
 * CSS linear-gradient 角度 → OOXML lin ang（单位：60000ths of a degree）
 *
 * CSS 角度语义：0deg = 向上，顺时针增大
 * OOXML 角度语义：0 = 向右（3点钟方向），顺时针增大，单位 1/60000 度
 *
 * 转换：ooxmlDeg = (90 - cssDeg + 360) % 360
 * 再乘 60000 → OOXML ang 值
 */
function cssAngleToOoxml(cssDeg) {
  const ooxmlDeg = ((90 - cssDeg) % 360 + 360) % 360;
  return Math.round(ooxmlDeg * 60000);
}

// ─────────────────────────────────────────────────────
// 核心解析器
// ─────────────────────────────────────────────────────

/**
 * 从 computed style 字符串中提取第一个渐变层
 * 处理多层叠加：background 可能是 "url(...), linear-gradient(...)"
 */
function extractFirstGradientString(bgValue) {
  if (!bgValue || bgValue === 'none') return null;

  // 找第一个 *-gradient( 的位置
  const gradientTypes = [
    'linear-gradient',
    'radial-gradient',
    'conic-gradient',
    'repeating-linear-gradient',
    'repeating-radial-gradient',
  ];

  let firstIdx = -1;
  let foundType = null;
  for (const t of gradientTypes) {
    const idx = bgValue.indexOf(t);
    if (idx !== -1 && (firstIdx === -1 || idx < firstIdx)) {
      firstIdx = idx;
      foundType = t;
    }
  }

  if (firstIdx === -1) return null;

  // 提取括号内容（需要处理嵌套括号）
  let depth = 0;
  let start = bgValue.indexOf('(', firstIdx);
  let end = start;
  for (let i = start; i < bgValue.length; i++) {
    if (bgValue[i] === '(') depth++;
    else if (bgValue[i] === ')') {
      depth--;
      if (depth === 0) { end = i; break; }
    }
  }

  const innerContent = bgValue.slice(start + 1, end);
  return { type: foundType, content: innerContent };
}

/**
 * 解析 color-stop 列表
 * 浏览器格式：rgb(255, 59, 79) 0%, rgb(192, 32, 48) 100%
 * 或无位置：  rgb(255, 255, 255), rgb(0, 0, 0)
 *
 * 返回 [ { hex, alpha, position } ] position 为 0-100000
 */
function parseColorStops(stopsStr) {
  const stops = [];

  // 把 stops 字符串按逗号分割，但需要跳过 rgb(...) 内部的逗号
  const parts = [];
  let depth = 0;
  let current = '';
  for (const ch of stopsStr) {
    if (ch === '(') depth++;
    else if (ch === ')') depth--;
    if (ch === ',' && depth === 0) {
      parts.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  if (current.trim()) parts.push(current.trim());

  // 解析每个 stop
  let autoPositionCount = 0;
  const rawStops = parts.map(part => {
    part = part.trim();
    // 提取颜色（rgb/rgba）
    const colorMatch = part.match(/rgba?\([^)]+\)/);
    if (!colorMatch) return null;
    const color = parseColor(colorMatch[0]);
    if (!color) return null;

    // 提取位置（%）
    const posMatch = part.match(/([\d.]+)%/);
    const position = posMatch ? Math.round(parseFloat(posMatch[1]) * 1000) : null;

    return { ...color, position };
  }).filter(Boolean);

  // 填充自动位置：首个默认 0，末个默认 100000，中间均匀分布
  if (rawStops.length === 0) return [];
  if (rawStops[0].position === null) rawStops[0].position = 0;
  if (rawStops[rawStops.length - 1].position === null) rawStops[rawStops.length - 1].position = 100000;

  // 填充中间 null 位置（线性插值）
  let i = 0;
  while (i < rawStops.length) {
    if (rawStops[i].position === null) {
      // 找下一个有位置的 stop
      let j = i + 1;
      while (j < rawStops.length && rawStops[j].position === null) j++;
      const startPos = rawStops[i - 1].position;
      const endPos = rawStops[j].position;
      const steps = j - i + 1;
      for (let k = i; k < j; k++) {
        rawStops[k].position = Math.round(startPos + (endPos - startPos) * (k - i + 1) / steps);
      }
      i = j + 1;
    } else {
      i++;
    }
  }

  return rawStops;
}

/**
 * 解析 linear-gradient 内容
 * 浏览器 computed style 格式：第一个参数是角度（如 "135deg"），后跟 stops
 */
function parseLinear(content) {
  // 找角度（浏览器总是转换为 Ndeg 格式）
  const angleMatch = content.match(/^(-?[\d.]+)deg\s*,\s*/);
  let angle = 180; // 默认向下
  let stopsStr = content;

  if (angleMatch) {
    angle = parseFloat(angleMatch[1]);
    stopsStr = content.slice(angleMatch[0].length);
  }

  const stops = parseColorStops(stopsStr);
  if (stops.length < 2) return null;

  return {
    type: 'linear',
    angle,
    angOoxml: cssAngleToOoxml(angle),
    stops,
  };
}

/**
 * 解析 radial-gradient 内容
 * 浏览器格式：ellipse/circle at X% Y%, stop1, stop2
 * PPT 的 radial 支持 circle（path="circle"）和 rect（path="rect"），
 * 中心点通过 fillToRect l/t/r/b 控制
 */
function parseRadial(content) {
  // 解析 shape 和 center
  // 格式示例：
  //   "circle at 50% 50%, rgb(...) 0%, rgb(...) 100%"
  //   "ellipse at 30% 70%, ..."
  //   "closest-side, ..."  (无 at，浏览器可能省略)
  let shape = 'circle';
  let cx = 50, cy = 50; // center 百分比
  let stopsStr = content;

  // 匹配 "shape [size] at X% Y%," 前缀
  const headerMatch = content.match(
    /^(circle|ellipse)?\s*(?:[\w-]+\s+)?(?:at\s+([\d.]+)%\s+([\d.]+)%)?\s*,\s*/
  );
  if (headerMatch) {
    if (headerMatch[1]) shape = headerMatch[1];
    if (headerMatch[2]) cx = parseFloat(headerMatch[2]);
    if (headerMatch[3]) cy = parseFloat(headerMatch[3]);
    stopsStr = content.slice(headerMatch[0].length);
  }

  const stops = parseColorStops(stopsStr);
  if (stops.length < 2) return null;

  // 计算 fillToRect（中心点偏移）
  // fillToRect l/t/r/b 是渐变焦点到各边的距离（单位 1/1000 %，即 100000 = 100%）
  const l = Math.round(cx * 1000);
  const t = Math.round(cy * 1000);
  const r = Math.round((100 - cx) * 1000);
  const b = Math.round((100 - cy) * 1000);

  return {
    type: 'radial',
    shape, // 'circle' | 'ellipse'
    cx, cy,
    fillToRect: { l, t, r, b },
    stops,
  };
}

// ─────────────────────────────────────────────────────
// 公共 API
// ─────────────────────────────────────────────────────

/**
 * 主解析入口
 * @param {string} bgValue - getComputedStyle 返回的 background 或 background-image 值
 * @returns {object|null} 解析结果，null 表示无法识别（降级为纯色）
 *
 * 返回结构：
 *   { type: 'linear', angle, angOoxml, stops: [{hex, alpha, position}] }
 *   { type: 'radial', shape, cx, cy, fillToRect, stops }
 *   { type: 'unsupported', reason }  → 调用方降级为纯色
 */
function parseGradient(bgValue) {
  if (!bgValue) return null;

  const extracted = extractFirstGradientString(bgValue);
  if (!extracted) return null;

  const { type, content } = extracted;

  if (type === 'conic-gradient' || type === 'repeating-conic-gradient') {
    return { type: 'unsupported', reason: 'conic-gradient not supported in PPT' };
  }

  if (type === 'linear-gradient' || type === 'repeating-linear-gradient') {
    const result = parseLinear(content);
    if (!result) return { type: 'unsupported', reason: 'failed to parse linear-gradient' };
    return result;
  }

  if (type === 'radial-gradient' || type === 'repeating-radial-gradient') {
    const result = parseRadial(content);
    if (!result) return { type: 'unsupported', reason: 'failed to parse radial-gradient' };
    return result;
  }

  return null;
}

/**
 * 从渐变中取第一个 stop 的颜色，用作降级纯色
 */
function getFallbackColor(bgValue) {
  const extracted = extractFirstGradientString(bgValue);
  if (!extracted) return 'FFFFFF';

  const colorMatch = extracted.content.match(/rgba?\([^)]+\)/);
  if (!colorMatch) return 'FFFFFF';

  const color = parseColor(colorMatch[0]);
  return color ? color.hex : 'FFFFFF';
}

/**
 * 判断一个背景值是否包含渐变
 */
function hasGradient(bgValue) {
  if (!bgValue || bgValue === 'none') return false;
  return bgValue.includes('gradient');
}

module.exports = { parseGradient, getFallbackColor, hasGradient, parseColor, cssAngleToOoxml };
