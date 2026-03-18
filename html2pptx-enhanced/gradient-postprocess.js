'use strict';

/**
 * gradient-postprocess.js
 *
 * html2pptx 的渐变后处理模块。
 *
 * 工作流程：
 *   1. pptxgenjs 用降级纯色生成 PPTX（solidFill 作为占位）
 *   2. 本模块读入 PPTX，解压为内存 ZIP
 *   3. 对每张幻灯片的 XML：
 *      a. 背景渐变：替换 <p:bgPr> 内的 <a:solidFill> → <a:gradFill>
 *      b. shape/text 渐变：定位到对应 objectName 的 <p:sp>，
 *         替换其 <p:spPr> 内的 <a:solidFill> → <a:gradFill>
 *   4. 重新打包为 PPTX Buffer 写回磁盘
 *
 * 渐变注册表（GradientRegistry）：
 *   在 html2pptx 主流程中，每次提取到渐变元素时：
 *     registry.registerShape(slideIndex, objectName, gradient)
 *     registry.registerBackground(slideIndex, gradient)
 *   后处理时按 slideIndex 分组处理。
 */

const JSZip = require('jszip');
const fs = require('fs');
const { buildGradFillXml, buildBgPrXml } = require('./gradient-xml');

// ─────────────────────────────────────────────────────
// GradientRegistry：收集所有需要后处理的渐变信息
// ─────────────────────────────────────────────────────

class GradientRegistry {
  constructor() {
    // Map<slideIndex(1-based), { background?, shapes: Map<objectName, gradient> }>
    this._slides = new Map();
  }

  _getSlide(slideIndex) {
    if (!this._slides.has(slideIndex)) {
      this._slides.set(slideIndex, { background: null, shapes: new Map() });
    }
    return this._slides.get(slideIndex);
  }

  /**
   * 注册一个 shape/text 的渐变
   * @param {number} slideIndex - 1-based
   * @param {string} objectName - pptxgenjs 的 objectName，用于在 XML 中定位
   * @param {object} gradient   - gradient-parser.js 的解析结果
   */
  registerShape(slideIndex, objectName, gradient) {
    this._getSlide(slideIndex).shapes.set(objectName, gradient);
  }

  /**
   * 注册幻灯片背景渐变
   */
  registerBackground(slideIndex, gradient) {
    this._getSlide(slideIndex).background = gradient;
  }

  isEmpty() {
    return this._slides.size === 0;
  }

  hasSlide(slideIndex) {
    return this._slides.has(slideIndex);
  }

  getSlide(slideIndex) {
    return this._slides.get(slideIndex);
  }

  slideIndices() {
    return Array.from(this._slides.keys());
  }
}

// ─────────────────────────────────────────────────────
// XML 操作工具
// ─────────────────────────────────────────────────────

/**
 * 在 XML 字符串中，用 name 属性定位 <p:sp>（cNvPr name="..."），
 * 找到它的 <p:spPr> 内第一个 <a:solidFill>，替换为 gradFill XML。
 *
 * pptxgenjs 生成的 shape XML 结构（已验证）：
 *   <p:sp>
 *     <p:nvSpPr>
 *       <p:cNvPr id="N" name="OBJECT_NAME">...</p:cNvPr>
 *       ...
 *     </p:nvSpPr>
 *     <p:spPr>
 *       ...
 *       <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>  ← 替换这里
 *       ...
 *     </p:spPr>
 *     [<p:txBody>
 *       ...
 *       <a:solidFill>...</a:solidFill>  ← 文字颜色，不替换
 *     </p:txBody>]
 *   </p:sp>
 *
 * 策略：精确定位到目标 <p:sp> 块，只替换其 <p:spPr> 段内的 solidFill，
 * 不影响 <p:txBody> 里的文字颜色。
 */
function replaceShapeFill(xml, objectName, gradFillXml) {
  // 1. 找到包含此 objectName 的 <p:cNvPr> 的位置
  //    pptxgenjs 生成格式：<p:cNvPr id="N" name="OBJECT_NAME">
  const namePattern = new RegExp(
    `<p:cNvPr[^>]+name="${escapeRegex(objectName)}"[^>]*>`
  );
  const nameMatch = xml.match(namePattern);
  if (!nameMatch) {
    console.warn(`[gradient-postprocess] objectName not found in XML: ${objectName}`);
    return xml;
  }

  const nameIdx = xml.indexOf(nameMatch[0]);

  // 2. 从 nameIdx 往前找到包含它的 <p:sp> 起点
  const spStart = xml.lastIndexOf('<p:sp>', nameIdx);
  if (spStart === -1) return xml;

  // 3. 找到这个 <p:sp> 的结束位置（找匹配的 </p:sp>）
  let depth = 0;
  let spEnd = -1;
  let i = spStart;
  while (i < xml.length) {
    if (xml.startsWith('<p:sp>', i) || xml.startsWith('<p:sp ', i)) {
      depth++;
    } else if (xml.startsWith('</p:sp>', i)) {
      depth--;
      if (depth === 0) { spEnd = i + '</p:sp>'.length; break; }
    }
    i++;
  }
  if (spEnd === -1) return xml;

  const spBlock = xml.slice(spStart, spEnd);

  // 4. 在 spBlock 内找 <p:spPr>...</p:spPr>，替换其中的 solidFill
  const spPrStart = spBlock.indexOf('<p:spPr>');
  const spPrEnd = spBlock.indexOf('</p:spPr>');
  if (spPrStart === -1 || spPrEnd === -1) return xml;

  const spPrBlock = spBlock.slice(spPrStart, spPrEnd + '</p:spPr>'.length);

  // 5. 在 spPr 块内替换 <a:solidFill>...</a:solidFill>
  const newSpPrBlock = replaceSolidFillInBlock(spPrBlock, gradFillXml);

  // 6. 重组
  const newSpBlock = spBlock.slice(0, spPrStart) + newSpPrBlock + spBlock.slice(spPrEnd + '</p:spPr>'.length);
  return xml.slice(0, spStart) + newSpBlock + xml.slice(spEnd);
}

/**
 * 在 XML 块内替换第一个 <a:solidFill>...</a:solidFill> 为 gradFillXml
 */
function replaceSolidFillInBlock(block, gradFillXml) {
  // solidFill 可能是自闭合的单行，也可能跨多行
  // 格式：<a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
  const solidFillStart = block.indexOf('<a:solidFill>');
  if (solidFillStart === -1) return block;

  const solidFillEnd = block.indexOf('</a:solidFill>', solidFillStart);
  if (solidFillEnd === -1) return block;

  const endTagLen = '</a:solidFill>'.length;
  return (
    block.slice(0, solidFillStart) +
    gradFillXml +
    block.slice(solidFillEnd + endTagLen)
  );
}

/**
 * 替换幻灯片背景渐变
 * 目标：<p:bg><p:bgPr><a:solidFill>...</a:solidFill></p:bgPr></p:bg>
 *   → <p:bg><p:bgPr><a:gradFill>...</a:gradFill>...</p:bgPr></p:bg>
 */
function replaceBackgroundFill(xml, gradFillXml) {
  // 找 <p:bgPr>...</p:bgPr> 块
  const bgPrStart = xml.indexOf('<p:bgPr>');
  const bgPrEnd = xml.indexOf('</p:bgPr>');
  if (bgPrStart === -1 || bgPrEnd === -1) return xml;

  const bgPrBlock = xml.slice(bgPrStart, bgPrEnd + '</p:bgPr>'.length);
  const newBgPrBlock = replaceSolidFillInBlock(bgPrBlock, gradFillXml);

  return xml.slice(0, bgPrStart) + newBgPrBlock + xml.slice(bgPrEnd + '</p:bgPr>'.length);
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ─────────────────────────────────────────────────────
// 主后处理函数
// ─────────────────────────────────────────────────────

/**
 * 对 pptxgenjs 写出的 PPTX 文件做渐变后处理
 *
 * @param {string} pptxPath     - PPTX 文件路径（会被原地修改）
 * @param {GradientRegistry} registry
 * @returns {Promise<{processed: number, warnings: string[]}>}
 */
async function postprocess(pptxPath, registry) {
  if (registry.isEmpty()) return { processed: 0, warnings: [] };

  const warnings = [];
  let processedCount = 0;

  // 1. 读入 PPTX
  const buffer = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(buffer);

  // 2. 处理每张需要渐变的幻灯片
  for (const slideIndex of registry.slideIndices()) {
    const slideData = registry.getSlide(slideIndex);
    const slideFile = `ppt/slides/slide${slideIndex}.xml`;

    const zipEntry = zip.file(slideFile);
    if (!zipEntry) {
      warnings.push(`slide${slideIndex}.xml not found in PPTX`);
      continue;
    }

    let xml = await zipEntry.async('string');
    let modified = false;

    // 2a. 背景渐变
    if (slideData.background) {
      const gradFillXml = buildGradFillXml(slideData.background);
      if (gradFillXml) {
        const newXml = replaceBackgroundFill(xml, gradFillXml);
        if (newXml !== xml) {
          xml = newXml;
          modified = true;
          processedCount++;
        } else {
          warnings.push(`slide${slideIndex}: background solidFill not found, skipped`);
        }
      } else if (slideData.background.type === 'unsupported') {
        warnings.push(`slide${slideIndex} bg: ${slideData.background.reason}`);
      }
    }

    // 2b. Shape 渐变
    for (const [objectName, gradient] of slideData.shapes) {
      if (gradient.type === 'unsupported') {
        warnings.push(`slide${slideIndex} "${objectName}": ${gradient.reason}`);
        continue;
      }

      const gradFillXml = buildGradFillXml(gradient);
      if (!gradFillXml) {
        warnings.push(`slide${slideIndex} "${objectName}": buildGradFillXml returned null`);
        continue;
      }

      const newXml = replaceShapeFill(xml, objectName, gradFillXml);
      if (newXml !== xml) {
        xml = newXml;
        modified = true;
        processedCount++;
      } else {
        warnings.push(`slide${slideIndex} "${objectName}": solidFill not found or already replaced`);
      }
    }

    // 3. 写回 ZIP
    if (modified) {
      zip.file(slideFile, xml);
    }
  }

  // 4. 重新打包写回磁盘
  const outBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  fs.writeFileSync(pptxPath, outBuffer);

  return { processed: processedCount, warnings };
}

module.exports = { GradientRegistry, postprocess };
