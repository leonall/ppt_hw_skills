'use strict';

/**
 * gradient-xml.js
 *
 * 把 gradient-parser.js 输出的结构化渐变对象转换为 OOXML gradFill XML。
 *
 * OOXML 规范参考（ECMA-376 §20.1.8.33 gradFill）：
 *
 * 线性渐变：
 *   <a:gradFill rotWithShape="1">
 *     <a:gsLst>
 *       <a:gs pos="0">
 *         <a:srgbClr val="FF3B4F">
 *           <a:alpha val="100000"/>   <!-- 可选，完全不透明时省略 -->
 *         </a:srgbClr>
 *       </a:gs>
 *       <a:gs pos="100000">
 *         <a:srgbClr val="C02030"/>
 *       </a:gs>
 *     </a:gsLst>
 *     <a:lin ang="8100000" scaled="0"/>  <!-- ang = 135deg * 60000 -->
 *   </a:gradFill>
 *
 * 径向渐变：
 *   <a:gradFill rotWithShape="1">
 *     <a:gsLst>...</a:gsLst>
 *     <a:path path="circle">
 *       <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
 *     </a:path>
 *   </a:gradFill>
 */

/**
 * 生成单个 color stop 的 XML
 * position: 0-100000
 * hex: 'RRGGBB'
 * alpha: 0-100000（100000=完全不透明）
 */
function buildGsXml(position, hex, alpha) {
  const alphaXml = (alpha !== undefined && alpha < 100000)
    ? `<a:alpha val="${alpha}"/>`
    : '';
  return `<a:gs pos="${position}"><a:srgbClr val="${hex}">${alphaXml}</a:srgbClr></a:gs>`;
}

/**
 * 生成完整的 gradFill XML
 * @param {object} gradient - gradient-parser.js 解析结果
 * @returns {string} OOXML XML 字符串
 */
function buildGradFillXml(gradient) {
  if (!gradient || gradient.type === 'unsupported') return null;

  // 生成 gsLst
  const gsXmls = gradient.stops.map(s =>
    buildGsXml(s.position, s.hex, s.alpha)
  ).join('');
  const gsLst = `<a:gsLst>${gsXmls}</a:gsLst>`;

  if (gradient.type === 'linear') {
    // ang: OOXML 单位是 60000ths of a degree
    return (
      `<a:gradFill rotWithShape="1">` +
        gsLst +
        `<a:lin ang="${gradient.angOoxml}" scaled="0"/>` +
      `</a:gradFill>`
    );
  }

  if (gradient.type === 'radial') {
    const { l, t, r, b } = gradient.fillToRect;
    return (
      `<a:gradFill rotWithShape="1">` +
        gsLst +
        `<a:path path="circle">` +
          `<a:fillToRect l="${l}" t="${t}" r="${r}" b="${b}"/>` +
        `</a:path>` +
      `</a:gradFill>`
    );
  }

  return null;
}

/**
 * 生成幻灯片背景的 bgPr XML（替换整个 <p:bgPr> 内容）
 * @param {object} gradient
 * @returns {string}
 */
function buildBgPrXml(gradient) {
  const gradFill = buildGradFillXml(gradient);
  if (!gradFill) return null;
  // <p:bgPr> 标准结构：fill + effectLst（可选）
  return `<p:bgPr>${gradFill}<a:effectLst/></p:bgPr>`;
}

module.exports = { buildGradFillXml, buildBgPrXml };
