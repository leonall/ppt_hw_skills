'use strict';

/**
 * html2pptx - Convert HTML slide to pptxgenjs slide with positioned elements
 *
 * USAGE:
 *   const pptx = new pptxgen();
 *   pptx.layout = 'LAYOUT_16x9';
 *
 *   const { slide, placeholders } = await html2pptx('slide.html', pptx, { slideIndex: 1, registry });
 *   slide.addChart(pptx.charts.LINE, data, placeholders[0]);
 *
 *   await pptx.writeFile({ fileName: 'output.pptx' });
 *   await postprocess('output.pptx', registry);   // ← 渐变后处理
 *
 * GRADIENT SUPPORT:
 *   - body background: linear-gradient / radial-gradient → 真实 gradFill
 *   - div fill:        linear-gradient / radial-gradient → 真实 gradFill
 *   - conic-gradient → 降级为第一个 stop 纯色，并写 warning
 *   - 多层渐变叠加   → 只取第一个渐变层
 *   - 渐变 stop 透明度（rgba）→ 映射为 OOXML alpha
 *
 * TABLE SUPPORT (同上一版本，完整保留):
 *   - thead/tbody、per-cell 样式、colspan/rowspan
 *   - 渐变背景降级为第一个 stop 纯色
 *
 * EXPORTED:
 *   html2pptx(htmlFile, pres, options)  → { slide, placeholders }
 *   GradientRegistry                    → 跨 slide 收集渐变信息
 *   postprocess(pptxPath, registry)     → 写入真实渐变 XML
 */

const { chromium } = require('playwright');
const path = require('path');
const { GradientRegistry, postprocess } = require('./gradient-postprocess');
const { parseGradient, getFallbackColor, hasGradient } = require('./gradient-parser');
const { captureElements } = require('./element-screenshot');
const { handleOverflow } = require('./overflow-handler');

const PT_PER_PX = 0.75;
const PX_PER_IN = 96;
const EMU_PER_IN = 914400;

// ─────────────────────────────────────────────────────
// 校验工具
// ─────────────────────────────────────────────────────

async function getBodyDimensions(page) {
  const bodyDimensions = await page.evaluate(() => {
    const body = document.body;
    const style = window.getComputedStyle(body);
    return {
      width: parseFloat(style.width),
      height: parseFloat(style.height),
      scrollWidth: body.scrollWidth,
      scrollHeight: body.scrollHeight,
    };
  });

  const errors = [];
  const wOverPt = Math.max(0, bodyDimensions.scrollWidth - bodyDimensions.width - 1) * PT_PER_PX;
  const hOverPt = Math.max(0, bodyDimensions.scrollHeight - bodyDimensions.height - 1) * PT_PER_PX;
  if (wOverPt > 0 || hOverPt > 0) {
    const dirs = [];
    if (wOverPt > 0) dirs.push(`${wOverPt.toFixed(1)}pt horizontally`);
    if (hOverPt > 0) dirs.push(`${hOverPt.toFixed(1)}pt vertically`);
    const hint = hOverPt > 0 ? ' (Remember: leave 0.5" margin at bottom of slide)' : '';
    errors.push(`HTML content overflows body by ${dirs.join(' and ')}${hint}`);
  }
  return { ...bodyDimensions, errors };
}

function validateDimensions(bodyDimensions, pres, overflowMode) {
  const errors = [];
  const wIn = bodyDimensions.width / PX_PER_IN;
  // expand/clip 模式：body 可以不设固定高度（用 scrollHeight 代替）
  const effectiveH = (overflowMode !== 'error' && bodyDimensions.height < 1)
    ? bodyDimensions.scrollHeight
    : bodyDimensions.height;
  const hIn = effectiveH / PX_PER_IN;
  if (pres.presLayout) {
    const lw = pres.presLayout.width / EMU_PER_IN;
    const lh = pres.presLayout.height / EMU_PER_IN;
    // 宽度必须匹配；高度在 expand/clip 模式下只检查宽度
    const wMismatch = Math.abs(lw - wIn) > 0.1;
    const hMismatch = overflowMode === 'error' && Math.abs(lh - hIn) > 0.1;
    if (wMismatch || hMismatch) {
      errors.push(
        `HTML dimensions (${wIn.toFixed(1)}" × ${hIn.toFixed(1)}") ` +
        `don't match presentation layout (${lw.toFixed(1)}" × ${lh.toFixed(1)}")`
      );
    }
  }
  return errors;
}

function validateTextBoxPosition(slideData, bodyDimensions) {
  const errors = [];
  const slideH = bodyDimensions.height / PX_PER_IN;
  for (const el of slideData.elements) {
    if (!['p','h1','h2','h3','h4','h5','h6','list'].includes(el.type)) continue;
    const fontSize = el.style?.fontSize || 0;
    const bottom = el.position.y + el.position.h;
    const gap = slideH - bottom;
    if (fontSize > 12 && gap < 0.5) {
      const txt = (typeof el.text === 'string' ? el.text
        : Array.isArray(el.text) ? (el.text[0]?.text || '') : '').substring(0, 50);
      errors.push(
        `Text box "${txt}" ends too close to bottom edge ` +
        `(${gap.toFixed(2)}" from bottom, minimum 0.5" required)`
      );
    }
  }
  return errors;
}

// ─────────────────────────────────────────────────────
// addElements：逐元素调用 pptxgenjs API
// ─────────────────────────────────────────────────────

async function addBackground(slideData, targetSlide) {
  if (slideData.background.type === 'image' && slideData.background.path) {
    let p = slideData.background.path;
    if (p.startsWith('file://')) p = p.replace('file://', '');
    targetSlide.background = { path: p };
  } else if (slideData.background.type === 'color' && slideData.background.value) {
    targetSlide.background = { color: slideData.background.value };
  }
  // gradient background: 先用降级纯色占位，postprocess 再替换
  // (background.value 已在浏览器侧设为 fallback 颜色)
}

function addElements(slideData, targetSlide, pres) {
  for (const el of slideData.elements) {
    if (el.type === 'image') {
      let p = el.src.startsWith('file://') ? el.src.replace('file://', '') : el.src;
      targetSlide.addImage({ path: p, x: el.position.x, y: el.position.y, w: el.position.w, h: el.position.h });

    } else if (el.type === 'line') {
      targetSlide.addShape(pres.ShapeType.line, {
        x: el.x1, y: el.y1, w: el.x2 - el.x1, h: el.y2 - el.y1,
        line: { color: el.color, width: el.width },
      });

    } else if (el.type === 'shape') {
      const opts = {
        x: el.position.x, y: el.position.y, w: el.position.w, h: el.position.h,
        shape: el.shape.rectRadius > 0 ? pres.ShapeType.roundRect : pres.ShapeType.rect,
        objectName: el.objectName,   // ← 渐变定位用
      };
      if (el.shape.fill) {
        opts.fill = { color: el.shape.fill };
        if (el.shape.transparency != null) opts.fill.transparency = el.shape.transparency;
      }
      if (el.shape.line)       opts.line       = el.shape.line;
      if (el.shape.rectRadius > 0) opts.rectRadius = el.shape.rectRadius;
      if (el.shape.shadow)     opts.shadow     = el.shape.shadow;
      targetSlide.addShape(opts.shape, opts);

    } else if (el.type === 'list') {
      const opts = {
        x: el.position.x, y: el.position.y, w: el.position.w, h: el.position.h,
        fontSize: el.style.fontSize, fontFace: el.style.fontFace,
        color: el.style.color, align: el.style.align, valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore, paraSpaceAfter: el.style.paraSpaceAfter,
      };
      if (el.style.margin) opts.margin = el.style.margin;
      targetSlide.addText(el.items, opts);

    } else if (el.type === 'table') {
      addTable(targetSlide, el);

    } else {
      // TEXT (p, h1-h6)
      const lineHeight = el.style.lineSpacing || el.style.fontSize * 1.2;
      const isSingle = el.position.h <= lineHeight * 1.5;
      let ax = el.position.x, aw = el.position.w;
      if (isSingle) {
        const inc = aw * 0.02;
        const align = el.style.align;
        if (align === 'center') { ax -= inc / 2; aw += inc; }
        else if (align === 'right') { ax -= inc; aw += inc; }
        else aw += inc;
      }
      const opts = {
        x: ax, y: el.position.y, w: aw, h: el.position.h,
        fontSize: el.style.fontSize, fontFace: el.style.fontFace,
        color: el.style.color, bold: el.style.bold,
        italic: el.style.italic, underline: el.style.underline,
        valign: 'top', lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore, paraSpaceAfter: el.style.paraSpaceAfter,
        inset: 0,
      };
      if (el.style.align)   opts.align   = el.style.align;
      if (el.style.margin)  opts.margin  = el.style.margin;
      if (el.style.rotate !== undefined) opts.rotate = el.style.rotate;
      if (el.style.transparency != null) opts.transparency = el.style.transparency;
      targetSlide.addText(el.text, opts);
    }
  }
}

// ─────────────────────────────────────────────────────
// addTable（保持上一版本完整逻辑）
// ─────────────────────────────────────────────────────

function addTable(targetSlide, el) {
  const pptRows = el.rows.map(row =>
    row.map(cell => {
      const cellOptions = {};
      if (cell.fontSize)  cellOptions.fontSize  = cell.fontSize;
      if (cell.fontFace)  cellOptions.fontFace  = cell.fontFace;
      if (cell.bold)      cellOptions.bold       = true;
      if (cell.italic)    cellOptions.italic     = true;
      if (cell.color)     cellOptions.color      = cell.color;
      if (cell.align)     cellOptions.align      = cell.align;
      if (cell.valign)    cellOptions.valign     = cell.valign;
      if (cell.margin)    cellOptions.margin     = cell.margin;
      if (cell.fill && cell.fill !== 'transparent') cellOptions.fill = { color: cell.fill };
      if (cell.border)    cellOptions.border     = cell.border;
      if (cell.colspan && cell.colspan > 1) cellOptions.colspan = cell.colspan;
      if (cell.rowspan && cell.rowspan > 1) cellOptions.rowspan = cell.rowspan;
      return { text: cell._spanPlaceholder ? '' : (cell.text || ''), options: cellOptions };
    })
  );
  targetSlide.addTable(pptRows, {
    x: el.position.x, y: el.position.y, w: el.position.w,
    colW: el.colWidths, rowH: el.rowHeights, autoPage: false,
  });
}

// ─────────────────────────────────────────────────────
// BROWSER-SIDE: extractSlideData
// ─────────────────────────────────────────────────────

async function extractSlideData(page) {
  return await page.evaluate(() => {
    const PT_PER_PX = 0.75;
    const PX_PER_IN = 96;
    const SINGLE_WEIGHT_FONTS = ['impact'];

    // ── 工具函数 ────────────────────────────────────

    const shouldSkipBold = f => {
      if (!f) return false;
      return SINGLE_WEIGHT_FONTS.includes(f.toLowerCase().replace(/['"]/g,'').split(',')[0].trim());
    };

    const pxToInch  = px  => px / PX_PER_IN;
    const pxToPoints = px => parseFloat(px) * PT_PER_PX;

    const rgbToHex = s => {
      if (!s) return null;
      if (s === 'rgba(0, 0, 0, 0)' || s === 'transparent') return null;
      const m = s.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      if (!m) return null;
      return m.slice(1).map(n => parseInt(n).toString(16).padStart(2,'0')).join('').toUpperCase();
    };

    const extractAlpha = s => {
      const m = s && s.match(/rgba\([\d,\s]+,\s*([\d.]+)\)/);
      if (!m) return null;
      return Math.round((1 - parseFloat(m[1])) * 100);
    };

    const applyTextTransform = (t, tf) => {
      if (tf === 'uppercase') return t.toUpperCase();
      if (tf === 'lowercase') return t.toLowerCase();
      if (tf === 'capitalize') return t.replace(/\b\w/g, c => c.toUpperCase());
      return t;
    };

    // ── 渐变相关（浏览器侧只做"有无判断"和"降级颜色提取"） ──

    /**
     * 提取第一个 color-stop 作为降级纯色
     * 浏览器格式: linear-gradient(Ndeg, rgb(...) P%, rgb(...) P%, ...)
     */
    const extractGradientFallback = bgVal => {
      if (!bgVal || !bgVal.includes('gradient')) return null;
      const m = bgVal.match(/rgba?\([^)]+\)/);
      return m ? rgbToHex(m[0]) : null;
    };

    /**
     * 把 background / background-image 的渐变值原样传回 Node.js，
     * 由 gradient-parser.js 做完整解析。
     * 只需要检测"有没有渐变"并收集原始字符串。
     */
    const isGradient = val => val && val.includes('gradient');

    // 唯一 ID 生成器（用于 objectName，在 Node.js 侧定位 shape）
    let shapeCounter = 0;
    const nextObjectName = prefix => `${prefix}-grad-${shapeCounter++}`;

    // ── 背景 ────────────────────────────────────────

    const body = document.body;
    const bodyStyle = window.getComputedStyle(body);
    const bgImage = bodyStyle.backgroundImage;
    const bgColor = bodyStyle.backgroundColor;

    const errors = [];

    // 背景提取：区分渐变、图片、纯色三种情况
    let background;
    let backgroundGradientCss = null;  // 传回 Node.js 做完整解析

    if (bgImage && bgImage !== 'none') {
      if (isGradient(bgImage)) {
        // 渐变背景：fallback 纯色给 pptxgenjs，原始 CSS 传给后处理
        const fallback = extractGradientFallback(bgImage) || rgbToHex(bgColor) || 'FFFFFF';
        background = { type: 'color', value: fallback };
        backgroundGradientCss = bgImage;  // 原始 CSS string
      } else {
        // 图片背景
        const urlMatch = bgImage.match(/url\(["']?([^"')]+)["']?\)/);
        background = urlMatch
          ? { type: 'image', path: urlMatch[1] }
          : { type: 'color', value: rgbToHex(bgColor) || 'FFFFFF' };
      }
    } else {
      if (isGradient(bgColor)) {
        // 极少见：background 直接是渐变颜色值
        const fallback = extractGradientFallback(bgColor) || 'FFFFFF';
        background = { type: 'color', value: fallback };
        backgroundGradientCss = bgColor;
      } else {
        background = { type: 'color', value: rgbToHex(bgColor) || 'FFFFFF' };
      }
    }

    // body background gradient 报告给 Node.js（不再报错，改为 info）
    // 注意：原来这里会 errors.push，现在不报错，由 gradient-postprocess 处理

    // ── 表格提取工具 ────────────────────────────────

    const extractSolidColor = bgVal => {
      if (!bgVal || bgVal === 'none') return null;
      if (bgVal === 'rgba(0, 0, 0, 0)' || bgVal === 'transparent') return null;
      if (bgVal.includes('gradient')) {
        const m = bgVal.match(/rgba?\([^)]+\)/);
        return m ? rgbToHex(m[0]) : null;
      }
      return rgbToHex(bgVal);
    };

    const extractCellBorder = computed => {
      const sides = ['Top','Right','Bottom','Left'];
      const parseSide = side => {
        const w = parseFloat(computed[`border${side}Width`]) || 0;
        const s = computed[`border${side}Style`];
        const c = computed[`border${side}Color`];
        if (s === 'none' || w === 0) return { type: 'none' };
        return { type: 'solid', pt: Math.max(0.25, w * PT_PER_PX), color: rgbToHex(c) || '000000' };
      };
      const [top, right, bottom, left] = sides.map(parseSide);
      if ([top,right,bottom,left].every(b => b.type === 'none')) return [{type:'none'},{type:'none'},{type:'none'},{type:'none'}];
      const allSame = [right,bottom,left].every(b => b.type===top.type && b.pt===top.pt && b.color===top.color);
      if (allSame && top.type !== 'none') return { type: top.type, pt: top.pt, color: top.color };
      return [top, right, bottom, left];
    };

    const extractCellMargin = computed => {
      const t = pxToInch(parseFloat(computed.paddingTop) || 0);
      const r = pxToInch(parseFloat(computed.paddingRight) || 0);
      const b = pxToInch(parseFloat(computed.paddingBottom) || 0);
      const l = pxToInch(parseFloat(computed.paddingLeft) || 0);
      if (t===0 && r===0 && b===0 && l===0) return null;
      return [t, r, b, l];
    };

    const extractTable = tableEl => {
      const tableRect = tableEl.getBoundingClientRect();
      if (tableRect.width === 0 || tableRect.height === 0) return null;
      const allRows = Array.from(tableEl.querySelectorAll('tr'));
      const colWidths = [];
      const rowHeights = allRows.map(tr => pxToInch(tr.getBoundingClientRect().height) || 0.25);
      const occupiedMap = allRows.map(() => []);
      const extractedRows = [];

      allRows.forEach((tr, rowIdx) => {
        const cells = Array.from(tr.cells);
        const rowData = [];
        let domCellIdx = 0, visualColIdx = 0;

        while (domCellIdx < cells.length || visualColIdx < 20) {
          if (occupiedMap[rowIdx] && occupiedMap[rowIdx][visualColIdx]) {
            rowData.push({ _spanPlaceholder: true });
            visualColIdx++; continue;
          }
          if (domCellIdx >= cells.length) break;

          const cell = cells[domCellIdx];
          const colspan = parseInt(cell.getAttribute('colspan') || 1);
          const rowspan = parseInt(cell.getAttribute('rowspan') || 1);
          const computed = window.getComputedStyle(cell);
          const isHeader = cell.tagName === 'TH';

          if (rowIdx === 0 && colspan === 1) {
            colWidths[visualColIdx] = pxToInch(cell.getBoundingClientRect().width);
          }
          if (rowIdx === 0 && colspan > 1) {
            const tw = pxToInch(cell.getBoundingClientRect().width);
            for (let c = 0; c < colspan; c++) {
              if (!colWidths[visualColIdx + c]) colWidths[visualColIdx + c] = tw / colspan;
            }
          }

          const tt = computed.textTransform;
          const text = applyTextTransform(cell.innerText.trim(), tt);
          const fontSize = pxToPoints(computed.fontSize);
          const fontFace = computed.fontFamily.split(',')[0].replace(/['"]/g,'').trim();
          const isBold = isHeader || computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;
          let align = computed.textAlign;
          if (align === 'start') align = 'left';
          if (align === 'end') align = 'right';
          const valign = ({ top:'top', middle:'middle', bottom:'bottom', baseline:'bottom' })[computed.verticalAlign] || 'middle';
          const fill = extractSolidColor(computed.backgroundColor);
          const border = extractCellBorder(computed);
          const margin = extractCellMargin(computed);

          for (let r = 0; r < rowspan; r++) for (let c = 0; c < colspan; c++) {
            if (r === 0 && c === 0) continue;
            const tr2 = rowIdx + r, tc = visualColIdx + c;
            if (!occupiedMap[tr2]) occupiedMap[tr2] = [];
            occupiedMap[tr2][tc] = true;
          }

          rowData.push({ text, fontSize, fontFace, bold: isBold && !shouldSkipBold(fontFace),
            italic: computed.fontStyle === 'italic', color: rgbToHex(computed.color) || '000000',
            align, valign, fill, border, margin,
            colspan: colspan > 1 ? colspan : undefined,
            rowspan: rowspan > 1 ? rowspan : undefined });

          for (let c = 1; c < colspan; c++) rowData.push({ _spanPlaceholder: true });
          visualColIdx += colspan; domCellIdx++;
        }
        extractedRows.push(rowData);
      });

      const totalW = pxToInch(tableRect.width);
      const maxCols = Math.max(...extractedRows.map(r => r.length));
      for (let c = 0; c < maxCols; c++) {
        if (!colWidths[c] || colWidths[c] === 0) colWidths[c] = totalW / maxCols;
      }

      return {
        type: 'table', rows: extractedRows,
        position: { x: pxToInch(tableRect.left), y: pxToInch(tableRect.top),
                    w: pxToInch(tableRect.width), h: pxToInch(tableRect.height) },
        colWidths, rowHeights,
      };
    };

    // ── 其他工具函数 ────────────────────────────────

    const getRotation = (transform, writingMode) => {
      let angle = 0;
      if (writingMode === 'vertical-rl') angle = 90;
      else if (writingMode === 'vertical-lr') angle = 270;
      if (transform && transform !== 'none') {
        const rm = transform.match(/rotate\((-?[\d.]+)deg\)/);
        if (rm) angle += parseFloat(rm[1]);
        else {
          const mm = transform.match(/matrix\(([^)]+)\)/);
          if (mm) {
            const v = mm[1].split(',').map(parseFloat);
            angle += Math.round(Math.atan2(v[1], v[0]) * (180 / Math.PI));
          }
        }
      }
      angle = ((angle % 360) + 360) % 360;
      return angle === 0 ? null : angle;
    };

    const getPositionAndSize = (el, rect, rotation) => {
      if (rotation === null) return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
      const isVertical = rotation === 90 || rotation === 270;
      const cx = rect.left + rect.width / 2, cy = rect.top + rect.height / 2;
      if (isVertical) return { x: cx - rect.height/2, y: cy - rect.width/2, w: rect.height, h: rect.width };
      return { x: cx - el.offsetWidth/2, y: cy - el.offsetHeight/2, w: el.offsetWidth, h: el.offsetHeight };
    };

    const parseBoxShadow = bs => {
      if (!bs || bs === 'none') return null;
      if (bs.includes('inset')) return null;
      const cm = bs.match(/rgba?\([^)]+\)/);
      const parts = bs.match(/([-\d.]+)(px|pt)/g);
      if (!parts || parts.length < 2) return null;
      const ox = parseFloat(parts[0]), oy = parseFloat(parts[1]);
      const blur = parts.length > 2 ? parseFloat(parts[2]) : 0;
      let angle = 0;
      if (ox !== 0 || oy !== 0) { angle = Math.atan2(oy, ox) * (180/Math.PI); if (angle<0) angle+=360; }
      const offset = Math.sqrt(ox*ox+oy*oy) * PT_PER_PX;
      let opacity = 0.5;
      if (cm) { const om = cm[0].match(/[\d.]+\)$/); if (om) opacity = parseFloat(om[0].replace(')',''));}
      return { type:'outer', angle: Math.round(angle), blur: blur*0.75, color: cm ? (rgbToHex(cm[0])||'000000') : '000000', offset, opacity };
    };

    const parseInlineFormatting = (element, baseOptions={}, runs=[], baseTextTransform=x=>x) => {
      let prevNodeIsText = false;
      element.childNodes.forEach(node => {
        let textTransform = baseTextTransform;
        const isText = node.nodeType === Node.TEXT_NODE || node.tagName === 'BR';
        if (isText) {
          const text = node.tagName === 'BR' ? '\n' : textTransform(node.textContent.replace(/\s+/g,' '));
          const prev = runs[runs.length-1];
          if (prevNodeIsText && prev) prev.text += text;
          else runs.push({ text, options: {...baseOptions} });
        } else if (node.nodeType === Node.ELEMENT_NODE && node.textContent.trim()) {
          const options = {...baseOptions};
          const computed = window.getComputedStyle(node);
          const tags = ['SPAN','B','STRONG','I','EM','U'];
          if (tags.includes(node.tagName)) {
            const bold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;
            if (bold && !shouldSkipBold(computed.fontFamily)) options.bold = true;
            if (computed.fontStyle === 'italic') options.italic = true;
            if (computed.textDecoration?.includes('underline')) options.underline = true;
            if (computed.color && computed.color !== 'rgb(0, 0, 0)') {
              options.color = rgbToHex(computed.color);
              const tr = extractAlpha(computed.color);
              if (tr !== null) options.transparency = tr;
            }
            if (computed.fontSize) options.fontSize = pxToPoints(computed.fontSize);
            if (computed.textTransform && computed.textTransform !== 'none') {
              const tf = computed.textTransform;
              textTransform = t => applyTextTransform(t, tf);
            }
            parseInlineFormatting(node, options, runs, textTransform);
          }
        }
        prevNodeIsText = isText;
      });
      if (runs.length > 0) {
        runs[0].text = runs[0].text.replace(/^\s+/,'');
        runs[runs.length-1].text = runs[runs.length-1].text.replace(/\s+$/,'');
      }
      return runs.filter(r => r.text.length > 0);
    };

    // ── 主提取循环 ───────────────────────────────────

    const elements = [];
    const placeholders = [];
    const gradientShapes = [];  // { objectName, gradientCss } 传回 Node.js
    const textTags = ['P','H1','H2','H3','H4','H5','H6','UL','OL','LI'];
    const processed = new Set();

    document.querySelectorAll('*').forEach(el => {
      if (processed.has(el)) return;

      // ── TABLE ─────────────────────────────────────
      if (el.tagName === 'TABLE') {
        el.querySelectorAll('*').forEach(c => processed.add(c));
        processed.add(el);
        const tableData = extractTable(el);
        if (tableData) elements.push(tableData);
        return;
      }

      // ── VALIDATION ────────────────────────────────
      if (textTags.includes(el.tagName)) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasBorder = ['Top','Right','Bottom','Left'].some(s => parseFloat(computed[`border${s}Width`]) > 0);
        const hasShadow = computed.boxShadow && computed.boxShadow !== 'none';
        if (hasBg || hasBorder || hasShadow) {
          errors.push(`Text element <${el.tagName.toLowerCase()}> has ${hasBg?'background':hasBorder?'border':'shadow'}. Use <div> for styled containers.`);
          return;
        }
      }

      // ── PLACEHOLDER ───────────────────────────────
      const elClassName = typeof el.className === 'string' ? el.className : (el.className.baseVal || '');
      if (elClassName.includes('placeholder')) {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) {
          errors.push(`Placeholder "${el.id||'unnamed'}" has zero dimension.`);
        } else {
          placeholders.push({ id: el.id||`placeholder-${placeholders.length}`,
            x: pxToInch(rect.left), y: pxToInch(rect.top),
            w: pxToInch(rect.width), h: pxToInch(rect.height) });
        }
        processed.add(el); return;
      }

      // ── IMG ───────────────────────────────────────
      if (el.tagName === 'IMG') {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          elements.push({ type: 'image', src: el.src,
            position: { x: pxToInch(rect.left), y: pxToInch(rect.top), w: pxToInch(rect.width), h: pxToInch(rect.height) } });
          processed.add(el);
        }
        return;
      }

      // ── DIV → SHAPE（含渐变处理）──────────────────
      if (el.tagName === 'DIV') {
        const computed = window.getComputedStyle(el);
        const bgVal = computed.backgroundColor;
        const bgImg = computed.backgroundImage;

        // 检查 background-image 上的渐变（更常见）
        const activeBg = (bgImg && bgImg !== 'none') ? bgImg : bgVal;
        const hasGradientBg = isGradient(activeBg);
        const hasSolidBg = bgVal && bgVal !== 'rgba(0, 0, 0, 0)';

        for (const node of el.childNodes) {
          if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
            errors.push(`DIV contains unwrapped text "${node.textContent.trim().substring(0,50)}". Wrap in <p>.`);
          }
        }

        if (bgImg && bgImg !== 'none' && !isGradient(bgImg)) {
          errors.push('Background images on DIV not supported. Use <img> tag.');
          return;
        }

        const borders = ['Top','Right','Bottom','Left'].map(s => parseFloat(computed[`border${s}Width`]) || 0);
        const hasBorder = borders.some(b => b > 0);
        const hasUniformBorder = hasBorder && borders.every(b => b === borders[0]);
        const borderLines = [];

        if (hasBorder && !hasUniformBorder) {
          const rect = el.getBoundingClientRect();
          const x = pxToInch(rect.left), y = pxToInch(rect.top);
          const w = pxToInch(rect.width), h = pxToInch(rect.height);
          const sideInfo = [
            ['Top',    'borderTopWidth',    'borderTopColor',    x, y,     x+w, y    ],
            ['Right',  'borderRightWidth',  'borderRightColor',  x+w, y,   x+w, y+h  ],
            ['Bottom', 'borderBottomWidth', 'borderBottomColor', x, y+h,   x+w, y+h  ],
            ['Left',   'borderLeftWidth',   'borderLeftColor',   x, y,     x,   y+h  ],
          ];
          for (const [, wp, cp, x1, y1, x2, y2] of sideInfo) {
            if (parseFloat(computed[wp]) > 0) {
              const wpt = pxToPoints(computed[wp]);
              const inset = (wpt / 72) / 2;
              borderLines.push({ type:'line', x1, y1: y1+inset, x2, y2: y2-inset, width: wpt, color: rgbToHex(computed[cp])||'000000' });
            }
          }
        }

        if (hasGradientBg || hasSolidBg || hasBorder) {
          const rect = el.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) {
            // 准备 objectName（渐变时需要唯一名称定位）
            const needsGradient = hasGradientBg;
            const objName = needsGradient ? nextObjectName('shape') : undefined;

            // fill：渐变时用 fallback 纯色
            let fillHex;
            if (hasGradientBg) {
              fillHex = extractGradientFallback(activeBg);
            } else if (hasSolidBg) {
              fillHex = rgbToHex(bgVal);
            }

            const shadow = parseBoxShadow(computed.boxShadow);
            const radius = (() => {
              const r = computed.borderRadius, rv = parseFloat(r);
              if (rv === 0) return 0;
              if (r.includes('%')) return rv >= 50 ? 1 : (rv/100) * pxToInch(Math.min(rect.width, rect.height));
              if (r.includes('pt')) return rv / 72;
              return rv / PX_PER_IN;
            })();

            if (hasSolidBg || hasGradientBg || hasUniformBorder) {
              elements.push({
                type: 'shape',
                objectName: objName,  // undefined 表示不需要渐变后处理
                position: { x: pxToInch(rect.left), y: pxToInch(rect.top), w: pxToInch(rect.width), h: pxToInch(rect.height) },
                shape: {
                  fill: fillHex || null,
                  transparency: hasSolidBg && !hasGradientBg ? extractAlpha(bgVal) : null,
                  line: hasUniformBorder ? { color: rgbToHex(computed.borderColor)||'000000', width: pxToPoints(computed.borderWidth) } : null,
                  rectRadius: radius,
                  shadow,
                },
              });

              // 注册到渐变列表（传回 Node.js 侧做完整解析）
              if (needsGradient) {
                gradientShapes.push({ objectName: objName, gradientCss: activeBg });
              }
            }

            elements.push(...borderLines.map(l => ({
              type: 'line', x1: l.x1, y1: l.y1, x2: l.x2, y2: l.y2, color: l.color, width: l.width,
            })));
            processed.add(el);
            return;
          }
        }
      }

      // ── UL / OL → LIST ───────────────────────────
      if (el.tagName === 'UL' || el.tagName === 'OL') {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return;
        const lis = Array.from(el.querySelectorAll('li'));
        const items = [];
        const ulComputed = window.getComputedStyle(el);
        const ulPL = pxToPoints(ulComputed.paddingLeft);
        const marginLeft = ulPL * 0.5, textIndent = ulPL * 0.5;
        lis.forEach((li, idx) => {
          const isLast = idx === lis.length - 1;
          const runs = parseInlineFormatting(li, { breakLine: false });
          if (runs.length > 0) { runs[0].text = runs[0].text.replace(/^[•\-\*▪▸]\s*/,''); runs[0].options.bullet = { indent: textIndent }; }
          if (runs.length > 0 && !isLast) runs[runs.length-1].options.breakLine = true;
          items.push(...runs);
        });
        const computed = window.getComputedStyle(lis[0] || el);
        elements.push({ type:'list', items,
          position: { x: pxToInch(rect.left), y: pxToInch(rect.top), w: pxToInch(rect.width), h: pxToInch(rect.height) },
          style: { fontSize: pxToPoints(computed.fontSize), fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g,'').trim(),
            color: rgbToHex(computed.color)||'000000', transparency: extractAlpha(computed.color),
            align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
            lineSpacing: computed.lineHeight && computed.lineHeight!=='normal' ? pxToPoints(computed.lineHeight) : null,
            paraSpaceBefore: 0, paraSpaceAfter: pxToPoints(computed.marginBottom), margin: [marginLeft,0,0,0] } });
        lis.forEach(li => processed.add(li));
        processed.add(el); return;
      }

      // ── TEXT (P, H1-H6) ──────────────────────────
      if (!textTags.includes(el.tagName)) return;
      const rect = el.getBoundingClientRect();
      const text = el.textContent.trim();
      if (rect.width === 0 || rect.height === 0 || !text) return;
      if (el.tagName !== 'LI' && /^[•\-\*▪▸○●◆◇■□]\s/.test(text.trimStart())) {
        errors.push(`<${el.tagName.toLowerCase()}> starts with bullet symbol. Use <ul> or <ol>.`);
        return;
      }
      const computed = window.getComputedStyle(el);
      const rotation = getRotation(computed.transform, computed.writingMode);
      const { x, y, w, h } = getPositionAndSize(el, rect, rotation);
      const baseStyle = {
        fontSize: pxToPoints(computed.fontSize),
        fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g,'').trim(),
        color: rgbToHex(computed.color) || '000000',
        align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
        lineSpacing: pxToPoints(computed.lineHeight),
        paraSpaceBefore: pxToPoints(computed.marginTop),
        paraSpaceAfter: pxToPoints(computed.marginBottom),
        margin: [pxToPoints(computed.paddingLeft), pxToPoints(computed.paddingRight), pxToPoints(computed.paddingBottom), pxToPoints(computed.paddingTop)],
      };
      const tr = extractAlpha(computed.color);
      if (tr !== null) baseStyle.transparency = tr;
      if (rotation !== null) baseStyle.rotate = rotation;
      const hasFmt = el.querySelector('b,i,u,strong,em,span,br');
      if (hasFmt) {
        const runs = parseInlineFormatting(el, {}, [], s => applyTextTransform(s, computed.textTransform));
        const adj = {...baseStyle};
        if (adj.lineSpacing) {
          const maxFs = Math.max(adj.fontSize, ...runs.map(r => r.options?.fontSize||0));
          if (maxFs > adj.fontSize) adj.lineSpacing = maxFs * (adj.lineSpacing / adj.fontSize);
        }
        elements.push({ type: el.tagName.toLowerCase(), text: runs,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) }, style: adj });
      } else {
        const tt = applyTextTransform(text, computed.textTransform);
        const bold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;
        elements.push({ type: el.tagName.toLowerCase(), text: tt,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: { ...baseStyle, bold: bold && !shouldSkipBold(computed.fontFamily), italic: computed.fontStyle==='italic', underline: computed.textDecoration?.includes('underline') } });
      }
      processed.add(el);
    });

    return { background, backgroundGradientCss, elements, gradientShapes, placeholders, errors };
  });
}

// ─────────────────────────────────────────────────────
// 主入口
// ─────────────────────────────────────────────────────

/**
 * @param {string} htmlFile     - HTML 文件路径
 * @param {object} pres         - pptxgenjs Presentation 实例
 * @param {object} [options]
 * @param {number} [options.slideIndex=1]       - 当前幻灯片在 PPTX 中的序号（1-based）
 * @param {GradientRegistry} [options.registry] - 渐变注册表（多 slide 共享）
 * @param {string} [options.tmpDir]
 * @param {object} [options.slide]              - 复用已有 slide 对象
 * @param {boolean} [options.captureCanvas=true]  - 是否截图 <canvas> 元素（图表）
 * @param {boolean} [options.captureSvg=true]     - 是否截图 <svg> 元素
 * @param {number}  [options.screenshotScale=2]   - 截图渲染倍数（1 或 2）
 * @param {number}  [options.screenshotTimeout=8000] - 图表等待超时 ms
 * @param {'expand'|'clip'|'error'} [options.overflowMode='error']
 *   'expand' 展开：等比缩放所有元素到画布内
 *   'clip'   截取：截断/丢弃超出底部的元素
 *   'error'  原有行为：overflow 时抛出错误
 */
async function html2pptx(htmlFile, pres, options = {}) {
  const {
    slideIndex      = 1,
    registry        = null,
    tmpDir          = process.env.TMPDIR || '/tmp',
    slide           = null,
    captureCanvas   = true,
    captureSvg      = true,
    screenshotScale   = 2,
    screenshotTimeout = 8000,
    overflowMode    = 'error',   // 'expand' | 'clip' | 'error'
  } = options;

  try {
    const execPath = process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH || undefined;
    const launchOptions = {
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
      env: { TMPDIR: tmpDir },
    };
    if (execPath) launchOptions.executablePath = execPath;
    if (process.platform === 'darwin') launchOptions.channel = 'chrome';

    const browser = await chromium.launch(launchOptions);
    let bodyDimensions, slideData, capturedElements = [];
    const filePath = path.isAbsolute(htmlFile) ? htmlFile : path.join(process.cwd(), htmlFile);
    const fileUrl  = `file://${filePath}`;
    const validationErrors = [];

    try {
      // ── 第一阶段：结构提取（1x page） ─────────────
      const page = await browser.newPage();
      page.on('console', msg => console.log(`[browser] ${msg.text()}`));
      await page.goto(fileUrl);

      bodyDimensions = await getBodyDimensions(page);
      await page.setViewportSize({
        width:  Math.round(bodyDimensions.width),
        height: Math.round(bodyDimensions.height),
      });

      slideData = await extractSlideData(page);

      // ── 第二阶段：Canvas / SVG 截图（2x page） ────
      // 只要有 canvas 或 svg 需要截图就启动
      const skipTags = new Set();
      if (!captureCanvas) skipTags.add('CANVAS');
      if (!captureSvg)    skipTags.add('SVG');

      if (skipTags.size < 2) {
        // captureElements 内部会创建新的 2x context 并重新加载同一 URL
        capturedElements = await captureElements(page, {
          scale:   screenshotScale,
          timeout: screenshotTimeout,
          skipTags,
        });
      }

    } finally {
      await browser.close();
    }

    // ── 校验 ──────────────────────────────────────

    // overflow 错误只在 'error' 模式下才阻断流程；
    // expand/clip 模式把 overflow 错误过滤掉，由后续逻辑处理
    const overflowErrors = (bodyDimensions.errors || [])
      .filter(e => e.includes('overflows body'));
    const nonOverflowErrors = (bodyDimensions.errors || [])
      .filter(e => !e.includes('overflows body'));

    validationErrors.push(...nonOverflowErrors);
    validationErrors.push(...validateDimensions(bodyDimensions, pres, overflowMode));
    validationErrors.push(...(slideData.errors || []));

    // overflow 错误的处理取决于模式
    if (overflowMode === 'error') {
      validationErrors.push(...overflowErrors);
      // 原有的 textBoxPosition 校验也只在 error 模式下执行
      validationErrors.push(...validateTextBoxPosition(slideData, bodyDimensions));
    }
    // expand / clip 模式：忽略 overflow 错误和 textBoxPosition 警告

    if (validationErrors.length > 0) {
      const msg = validationErrors.length === 1
        ? validationErrors[0]
        : `Multiple validation errors:\n${validationErrors.map((e,i) => `  ${i+1}. ${e}`).join('\n')}`;
      throw new Error(`${htmlFile}: ${msg}`);
    }

    // ── Overflow 处理 ────────────────────────────

    let overflowResult = { didTransform: false, transformInfo: 'no overflow', scale: 1, offsetX: 0 };
    if (pres.presLayout) {
      overflowResult = handleOverflow(
        slideData, bodyDimensions, pres.presLayout, overflowMode
      );
      slideData = overflowResult.slideData;

      if (overflowResult.didTransform) {
        console.log(`[html2pptx] slide${slideIndex} overflow: ${overflowResult.transformInfo}`);
      }
    }

    // ── 注册渐变到 Registry ──────────────────────────

    if (registry) {
      if (slideData.backgroundGradientCss) {
        const gradient = parseGradient(slideData.backgroundGradientCss);
        if (gradient) registry.registerBackground(slideIndex, gradient);
      }
      for (const { objectName, gradientCss } of (slideData.gradientShapes || [])) {
        const gradient = parseGradient(gradientCss);
        if (gradient) {
          registry.registerShape(slideIndex, objectName, gradient);
        } else {
          console.warn(`[html2pptx] slide${slideIndex} "${objectName}": could not parse gradient`);
        }
      }
    }

    // ── 构建幻灯片 ────────────────────────────────

    const targetSlide = slide || pres.addSlide();
    await addBackground(slideData, targetSlide);
    addElements(slideData, targetSlide, pres);

    // ── 嵌入截图（Canvas / SVG → PNG image） ───────
    // expand 模式需要对截图坐标同步应用 scale 和 offsetX

    const { scale: ovScale, offsetX: ovOffsetX } = overflowResult;
    const screenshotWarnings = [];

    for (const captured of capturedElements) {
      if (!captured.dataUrl) {
        screenshotWarnings.push(
          `${captured.tag} "${captured.objectName}": screenshot failed — ${captured.error || 'unknown'}`
        );
        continue;
      }

      // 在 expand 模式下对截图坐标应用相同的缩放
      let pos = { ...captured.position };
      if (ovScale !== 1 || ovOffsetX !== 0) {
        if (overflowMode === 'expand') {
          pos = {
            x: pos.x * ovScale + ovOffsetX,
            y: pos.y * ovScale,
            w: pos.w * ovScale,
            h: pos.h * ovScale,
          };
        } else if (overflowMode === 'clip') {
          // clip 模式：超出底部的截图也丢弃
          const EMU_PER_IN = 914400;
          const slideH = pres.presLayout.height / EMU_PER_IN;
          if (pos.y >= slideH) continue; // 完全超出，跳过
          if (pos.y + pos.h > slideH) pos.h = slideH - pos.y; // 截断
        }
      }

      targetSlide.addImage({
        data: captured.dataUrl,
        x:    pos.x,
        y:    pos.y,
        w:    pos.w,
        h:    pos.h,
      });
    }

    if (screenshotWarnings.length > 0) {
      console.warn(`[html2pptx] slide${slideIndex} screenshot warnings:`);
      screenshotWarnings.forEach(w => console.warn(`  ⚠ ${w}`));
    }

    return {
      slide:         targetSlide,
      placeholders:  slideData.placeholders,
      screenshots:   capturedElements,
      overflow:      overflowResult,   // { didTransform, transformInfo, scale, offsetX, mode }
    };

  } catch (error) {
    if (!error.message.startsWith(htmlFile)) {
      throw new Error(`${htmlFile}: ${error.message}`);
    }
    throw error;
  }
}

// 导出
module.exports = { html2pptx, GradientRegistry, postprocess };
