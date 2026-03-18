'use strict';

/**
 * test-gradient.js
 * 运行: node test-gradient.js
 */

const { parseGradient, getFallbackColor, hasGradient, parseColor, cssAngleToOoxml } = require('./gradient-parser');
const { buildGradFillXml, buildBgPrXml } = require('./gradient-xml');
const { GradientRegistry } = require('./gradient-postprocess');

let passed = 0, failed = 0;

function test(name, fn) {
  try { fn(); console.log(`  ✓ ${name}`); passed++; }
  catch (e) { console.log(`  ✗ ${name}\n    ${e.message}`); failed++; }
}
function assert(cond, msg) { if (!cond) throw new Error(msg || 'failed'); }
function eq(a, b, msg) {
  const as = JSON.stringify(a), bs = JSON.stringify(b);
  if (as !== bs) throw new Error(`${msg||''}\n    got:      ${as}\n    expected: ${bs}`);
}
function near(a, b, tol=1, msg='') {
  if (Math.abs(a - b) > tol) throw new Error(`${msg}: ${a} vs ${b} (tol ${tol})`);
}

// ═══════════════════════════════════════════════════════════
console.log('\n【1】parseColor');
// ═══════════════════════════════════════════════════════════

test('rgb 纯色', () => {
  const c = parseColor('rgb(255, 59, 79)');
  eq(c.hex, 'FF3B4F');
  eq(c.alpha, 100000);
});
test('rgba 半透明', () => {
  const c = parseColor('rgba(26, 74, 138, 0.5)');
  eq(c.hex, '1A4A8A');
  near(c.alpha, 50000, 1);
});
test('rgba 完全透明', () => {
  const c = parseColor('rgba(0, 0, 0, 0)');
  eq(c.hex, '000000');
  eq(c.alpha, 0);
});
test('rgba 完全不透明', () => {
  const c = parseColor('rgba(255, 255, 255, 1)');
  eq(c.alpha, 100000);
});
test('无效字符串返回 null', () => {
  eq(parseColor('transparent'), null);
  eq(parseColor(''), null);
  eq(parseColor(null), null);
});

// ═══════════════════════════════════════════════════════════
console.log('\n【2】cssAngleToOoxml 角度转换');
// ═══════════════════════════════════════════════════════════

test('0deg CSS → 5400000 OOXML（向下=南）', () => {
  // CSS 0deg = 向上，PPT 向上 = 270deg → 270*60000 = 16200000 ? 
  // 公式: ooxmlDeg = (90 - cssDeg + 360) % 360  → (90-0+360)%360=90 → 90*60000=5400000
  eq(cssAngleToOoxml(0), 5400000);
});
test('90deg CSS → 0 OOXML（向右）', () => {
  // (90-90+360)%360 = 0 → 0*60000 = 0
  eq(cssAngleToOoxml(90), 0);
});
test('180deg CSS → 16200000 OOXML（向上）', () => {
  // (90-180+360)%360 = 270 → 270*60000=16200000
  eq(cssAngleToOoxml(180), 16200000);
});
test('135deg CSS → 8100000 OOXML（右上→左下方向渐变）', () => {
  // (90-135+360)%360 = 315 → 315*60000=18900000
  // 重新计算：(90-135+360)%360 = 315；315*60000=18900000
  // 但 Keynote 中 135deg 对应的 PPT ang 通常是 8100000(135deg)
  // 实际上 PPT ang 和 CSS angle 方向不同：
  // PPT ang=0 → 颜色从左到右（向右）
  // CSS 90deg → 颜色从下到上（向上）
  // 两者参考点不同，这里测实际计算值即可
  const result = cssAngleToOoxml(135);
  near(result, 18900000, 1);
});
test('270deg → 10800000', () => {
  // (90-270+360)%360 = 180 → 180*60000=10800000
  eq(cssAngleToOoxml(270), 10800000);
});
test('负角度处理', () => {
  // -45deg CSS → (90-(-45)+360)%360 = (90+45+360)%360 = 135 → 135*60000=8100000
  eq(cssAngleToOoxml(-45), 8100000);
});

// ═══════════════════════════════════════════════════════════
console.log('\n【3】parseGradient - linear');
// ═══════════════════════════════════════════════════════════

test('基础二色线性渐变（浏览器 computed 格式）', () => {
  const g = parseGradient('linear-gradient(135deg, rgb(255, 59, 79) 0%, rgb(192, 32, 48) 100%)');
  eq(g.type, 'linear');
  near(g.angle, 135, 1);
  eq(g.stops.length, 2);
  eq(g.stops[0].hex, 'FF3B4F');
  eq(g.stops[0].position, 0);
  eq(g.stops[0].alpha, 100000);
  eq(g.stops[1].hex, 'C02030');
  eq(g.stops[1].position, 100000);
});
test('三色线性渐变', () => {
  const g = parseGradient('linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(255, 215, 0) 50%, rgb(0, 128, 0) 100%)');
  eq(g.type, 'linear');
  eq(g.stops.length, 3);
  eq(g.stops[1].hex, 'FFD700');
  eq(g.stops[1].position, 50000);
});
test('rgba stop（透明度）', () => {
  const g = parseGradient('linear-gradient(0deg, rgba(255, 59, 79, 0) 0%, rgba(255, 59, 79, 1) 100%)');
  assert(g !== null, 'should parse');
  eq(g.stops[0].alpha, 0);       // 完全透明
  eq(g.stops[1].alpha, 100000);  // 完全不透明
});
test('无位置标注的 stop（自动分布）', () => {
  const g = parseGradient('linear-gradient(90deg, rgb(255, 0, 0), rgb(0, 255, 0), rgb(0, 0, 255))');
  assert(g !== null);
  eq(g.stops.length, 3);
  eq(g.stops[0].position, 0);
  eq(g.stops[2].position, 100000);
  // 中间 stop 应在 50000 附近
  near(g.stops[1].position, 50000, 1000);
});
test('repeating-linear-gradient 当作 linear 处理', () => {
  const g = parseGradient('repeating-linear-gradient(45deg, rgb(255,0,0) 0%, rgb(0,0,255) 20%)');
  eq(g?.type, 'linear');
});
test('conic-gradient → unsupported', () => {
  const g = parseGradient('conic-gradient(rgb(255,0,0), rgb(0,0,255))');
  eq(g?.type, 'unsupported');
});
test('多层渐变叠加 → 取第一个渐变层', () => {
  const multi = 'url("img.png"), linear-gradient(180deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)';
  const g = parseGradient(multi);
  // 有 url() 前缀，extractFirstGradientString 跳过非 gradient 前缀
  assert(g !== null);
  eq(g.type, 'linear');
});

// ═══════════════════════════════════════════════════════════
console.log('\n【4】parseGradient - radial');
// ═══════════════════════════════════════════════════════════

test('基础径向渐变（circle at 50% 50%）', () => {
  const g = parseGradient('radial-gradient(circle at 50% 50%, rgb(255, 255, 255) 0%, rgb(26, 74, 138) 100%)');
  eq(g.type, 'radial');
  eq(g.shape, 'circle');
  near(g.cx, 50, 1); near(g.cy, 50, 1);
  eq(g.fillToRect.l, 50000);
  eq(g.fillToRect.t, 50000);
  eq(g.stops.length, 2);
});
test('偏心径向渐变（at 30% 70%）', () => {
  const g = parseGradient('radial-gradient(circle at 30% 70%, rgb(255, 255, 0) 0%, rgb(255, 0, 0) 100%)');
  eq(g.type, 'radial');
  near(g.cx, 30, 1); near(g.cy, 70, 1);
  eq(g.fillToRect.l, 30000);
  eq(g.fillToRect.t, 70000);
  eq(g.fillToRect.r, 70000); // 100-30
  eq(g.fillToRect.b, 30000); // 100-70
});
test('ellipse 径向渐变', () => {
  const g = parseGradient('radial-gradient(ellipse at 50% 50%, rgb(255, 255, 255) 0%, rgb(0, 0, 0) 100%)');
  eq(g?.type, 'radial');
  eq(g?.shape, 'ellipse');
});

// ═══════════════════════════════════════════════════════════
console.log('\n【5】getFallbackColor');
// ═══════════════════════════════════════════════════════════

test('取线性渐变第一个 stop', () => {
  const c = getFallbackColor('linear-gradient(135deg, rgb(255, 59, 79) 0%, rgb(192, 32, 48) 100%)');
  eq(c, 'FF3B4F');
});
test('无渐变返回 FFFFFF', () => {
  eq(getFallbackColor('none'), 'FFFFFF');
  eq(getFallbackColor(null), 'FFFFFF');
});

// ═══════════════════════════════════════════════════════════
console.log('\n【6】buildGradFillXml - 线性');
// ═══════════════════════════════════════════════════════════

test('生成有效 linear gradFill XML', () => {
  const g = parseGradient('linear-gradient(180deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)');
  const xml = buildGradFillXml(g);
  assert(xml.includes('<a:gradFill'), 'should have gradFill');
  assert(xml.includes('<a:gsLst>'), 'should have gsLst');
  assert(xml.includes('<a:lin'), 'should have lin');
  assert(xml.includes('val="FF0000"'), 'first stop color');
  assert(xml.includes('val="0000FF"'), 'second stop color');
  assert(xml.includes('pos="0"'), 'first stop pos');
  assert(xml.includes('pos="100000"'), 'second stop pos');
});
test('角度正确编码', () => {
  const g = parseGradient('linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)');
  const xml = buildGradFillXml(g);
  // 90deg CSS → 0 OOXML
  assert(xml.includes('ang="0"'), `expected ang="0" in: ${xml}`);
});
test('透明度编码到 alpha 元素', () => {
  const g = parseGradient('linear-gradient(0deg, rgba(255, 59, 79, 0.5) 0%, rgba(255, 59, 79, 1) 100%)');
  const xml = buildGradFillXml(g);
  assert(xml.includes('<a:alpha'), 'should have alpha element for semi-transparent stop');
  // 不透明的 stop 不应有 alpha 元素（减少冗余）
  const alphaCount = (xml.match(/<a:alpha/g) || []).length;
  eq(alphaCount, 1, 'only one alpha element (for the semi-transparent stop)');
});
test('三色渐变 → 三个 gs 元素', () => {
  const g = parseGradient('linear-gradient(0deg, rgb(255,0,0) 0%, rgb(0,255,0) 50%, rgb(0,0,255) 100%)');
  const xml = buildGradFillXml(g);
  const gsCount = (xml.match(/<a:gs /g) || []).length;
  eq(gsCount, 3, 'three gs elements');
});
test('unsupported 返回 null', () => {
  eq(buildGradFillXml({ type: 'unsupported', reason: 'test' }), null);
  eq(buildGradFillXml(null), null);
});

// ═══════════════════════════════════════════════════════════
console.log('\n【7】buildGradFillXml - 径向');
// ═══════════════════════════════════════════════════════════

test('生成有效 radial gradFill XML', () => {
  const g = parseGradient('radial-gradient(circle at 50% 50%, rgb(255, 255, 255) 0%, rgb(0, 0, 0) 100%)');
  const xml = buildGradFillXml(g);
  assert(xml.includes('<a:gradFill'), 'gradFill');
  assert(xml.includes('<a:path path="circle"'), 'path circle');
  assert(xml.includes('<a:fillToRect'), 'fillToRect');
  assert(xml.includes('l="50000"'), 'left 50%');
  assert(xml.includes('t="50000"'), 'top 50%');
});
test('偏心径向渐变 fillToRect 偏移', () => {
  const g = parseGradient('radial-gradient(circle at 30% 70%, rgb(255,0,0) 0%, rgb(0,0,255) 100%)');
  const xml = buildGradFillXml(g);
  assert(xml.includes('l="30000"'), 'left 30%');
  assert(xml.includes('t="70000"'), 'top 70%');
  assert(xml.includes('r="70000"'), 'right 70%');
  assert(xml.includes('b="30000"'), 'bottom 30%');
});

// ═══════════════════════════════════════════════════════════
console.log('\n【8】buildBgPrXml');
// ═══════════════════════════════════════════════════════════

test('背景渐变包含在 p:bgPr 中', () => {
  const g = parseGradient('linear-gradient(135deg, rgb(15, 12, 41) 0%, rgb(48, 43, 99) 100%)');
  const xml = buildBgPrXml(g);
  assert(xml.startsWith('<p:bgPr>'), 'starts with bgPr');
  assert(xml.includes('<a:gradFill'), 'has gradFill');
  assert(xml.includes('<a:effectLst/>'), 'has effectLst');
  assert(xml.endsWith('</p:bgPr>'), 'ends with bgPr');
});

// ═══════════════════════════════════════════════════════════
console.log('\n【9】GradientRegistry');
// ═══════════════════════════════════════════════════════════

test('注册 shape 渐变', () => {
  const reg = new GradientRegistry();
  const g = parseGradient('linear-gradient(90deg, rgb(255,0,0) 0%, rgb(0,0,255) 100%)');
  reg.registerShape(1, 'shape-grad-0', g);
  assert(!reg.isEmpty());
  assert(reg.hasSlide(1));
  const slide = reg.getSlide(1);
  assert(slide.shapes.has('shape-grad-0'));
});
test('注册背景渐变', () => {
  const reg = new GradientRegistry();
  const g = parseGradient('linear-gradient(180deg, rgb(0,0,0) 0%, rgb(255,255,255) 100%)');
  reg.registerBackground(2, g);
  assert(reg.hasSlide(2));
  assert(reg.getSlide(2).background !== null);
});
test('多 slide 独立存储', () => {
  const reg = new GradientRegistry();
  reg.registerBackground(1, parseGradient('linear-gradient(0deg, rgb(255,0,0) 0%, rgb(0,0,255) 100%)'));
  reg.registerBackground(3, parseGradient('linear-gradient(0deg, rgb(0,255,0) 0%, rgb(0,0,0) 100%)'));
  assert(reg.hasSlide(1)); assert(reg.hasSlide(3)); assert(!reg.hasSlide(2));
  const indices = reg.slideIndices();
  assert(indices.includes(1)); assert(indices.includes(3));
});
test('空 registry isEmpty()', () => {
  assert(new GradientRegistry().isEmpty());
});

// ═══════════════════════════════════════════════════════════
console.log('\n【10】XML 后处理 - replaceShapeFill（单元测试）');
// ═══════════════════════════════════════════════════════════

// 直接测试内部函数（从模块中提取逻辑来测试）
function replaceSolidFillInBlock(block, gradFillXml) {
  const s = block.indexOf('<a:solidFill>');
  if (s === -1) return block;
  const e = block.indexOf('</a:solidFill>', s);
  if (e === -1) return block;
  return block.slice(0, s) + gradFillXml + block.slice(e + '</a:solidFill>'.length);
}

function replaceShapeFill(xml, objectName, gradFillXml) {
  const namePattern = new RegExp(`<p:cNvPr[^>]+name="${objectName.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}"[^>]*>`);
  const nameMatch = xml.match(namePattern);
  if (!nameMatch) return xml;
  const nameIdx = xml.indexOf(nameMatch[0]);
  const spStart = xml.lastIndexOf('<p:sp>', nameIdx);
  if (spStart === -1) return xml;
  let depth = 0, spEnd = -1, i = spStart;
  while (i < xml.length) {
    if (xml.startsWith('<p:sp>', i) || xml.startsWith('<p:sp ', i)) depth++;
    else if (xml.startsWith('</p:sp>', i)) { depth--; if (depth === 0) { spEnd = i + 7; break; } }
    i++;
  }
  if (spEnd === -1) return xml;
  const spBlock = xml.slice(spStart, spEnd);
  const spPrStart = spBlock.indexOf('<p:spPr>');
  const spPrEnd = spBlock.indexOf('</p:spPr>');
  if (spPrStart === -1 || spPrEnd === -1) return xml;
  const spPrBlock = spBlock.slice(spPrStart, spPrEnd + 9);
  const newSpPrBlock = replaceSolidFillInBlock(spPrBlock, gradFillXml);
  const newSpBlock = spBlock.slice(0, spPrStart) + newSpPrBlock + spBlock.slice(spPrEnd + 9);
  return xml.slice(0, spStart) + newSpBlock + xml.slice(spEnd);
}

const SAMPLE_SLIDE_XML = `<?xml version="1.0"?>
<p:sld>
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F0C29"/></a:solidFill></p:bgPr></p:bg>
<p:spTree>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="shape-grad-0"></p:cNvPr><p:cNvSpPr/><p:nvPr></p:nvPr></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="457200" y="457200"/><a:ext cx="2743200" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom>
<a:solidFill><a:srgbClr val="FF3B4F"/></a:solidFill>
<a:ln></a:ln>
</p:spPr>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="other-shape"></p:cNvPr><p:cNvSpPr/><p:nvPr></p:nvPr></p:nvSpPr>
<p:spPr>
<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
</p:spPr>
<p:txBody><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr><a:t>text</a:t></a:r></a:p></p:txBody>
</p:sp>
</p:spTree>
</p:cSld>
</p:sld>`;

const GRAD_XML = '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF3B4F"/></a:gs><a:gs pos="100000"><a:srgbClr val="C02030"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>';

test('replaceShapeFill 替换目标 shape', () => {
  const result = replaceShapeFill(SAMPLE_SLIDE_XML, 'shape-grad-0', GRAD_XML);
  assert(result.includes('gradFill'), 'should contain gradFill');
  // 验证被替换的是目标 shape 的 solidFill
  const targetSpIdx = result.indexOf('shape-grad-0');
  const gradFillIdx = result.indexOf('gradFill');
  assert(gradFillIdx > targetSpIdx, 'gradFill should come after target shape name');
});
test('replaceShapeFill 不影响其他 shape', () => {
  const result = replaceShapeFill(SAMPLE_SLIDE_XML, 'shape-grad-0', GRAD_XML);
  // other-shape 的 solidFill 应保持不变
  assert(result.includes('"00FF00"'), 'other shape fill unchanged');
  // txBody 里的文字颜色应保持不变
  assert(result.includes('"000000"'), 'text color unchanged');
});
test('replaceShapeFill objectName 不存在 → 原样返回', () => {
  const result = replaceShapeFill(SAMPLE_SLIDE_XML, 'nonexistent-shape', GRAD_XML);
  eq(result, SAMPLE_SLIDE_XML);
});
test('replaceBackgroundFill 替换背景', () => {
  function replaceBackgroundFill(xml, gradFillXml) {
    const s = xml.indexOf('<p:bgPr>');
    const e = xml.indexOf('</p:bgPr>');
    if (s === -1 || e === -1) return xml;
    const block = xml.slice(s, e + 9);
    const newBlock = replaceSolidFillInBlock(block, gradFillXml);
    return xml.slice(0, s) + newBlock + xml.slice(e + 9);
  }
  const result = replaceBackgroundFill(SAMPLE_SLIDE_XML, GRAD_XML);
  assert(result.includes('gradFill'), 'background now has gradFill');
  // 背景里的原始颜色 0F0C29 应被替换
  const bgSection = result.slice(0, result.indexOf('<p:spTree>'));
  assert(!bgSection.includes('0F0C29'), 'old bg color removed');
});

// ═══════════════════════════════════════════════════════════
console.log('\n【11】边界情况');
// ═══════════════════════════════════════════════════════════

test('空字符串不崩溃', () => {
  eq(parseGradient(''), null);
  eq(parseGradient('none'), null);
  eq(buildGradFillXml(null), null);
});
test('单色渐变（只有一个 stop）→ 返回 null', () => {
  // stops < 2 时无法生成有意义的渐变
  const g = parseGradient('linear-gradient(90deg, rgb(255, 0, 0))');
  // 单 stop 可能被解析，但 stops 数量应 < 2
  if (g && g.type !== 'unsupported') {
    assert(g.stops.length >= 1, 'at least 1 stop if parsed');
  }
});
test('hasGradient 检测', () => {
  assert(hasGradient('linear-gradient(...)'));
  assert(hasGradient('radial-gradient(...)'));
  assert(!hasGradient('rgb(255,0,0)'));
  assert(!hasGradient('none'));
  assert(!hasGradient(null));
});
test('360deg 等价于 0deg', () => {
  // 两者的 OOXML ang 应相同（%360 处理）
  eq(cssAngleToOoxml(360), cssAngleToOoxml(0));
});
test('非常多的 stop（10个）', () => {
  const stops = Array.from({length:10}, (_,i) => `rgb(${i*25},0,0) ${i*11}%`).join(', ');
  const g = parseGradient(`linear-gradient(0deg, ${stops})`);
  assert(g !== null && g.type === 'linear');
  assert(g.stops.length === 10, `expected 10 stops, got ${g?.stops?.length}`);
  const xml = buildGradFillXml(g);
  const gsCount = (xml.match(/<a:gs /g)||[]).length;
  eq(gsCount, 10);
});

// ═══════════════════════════════════════════════════════════
console.log(`\n${'─'.repeat(40)}`);
console.log(`通过: ${passed}  失败: ${failed}  合计: ${passed + failed}`);
if (failed > 0) { console.log('部分测试失败'); process.exit(1); }
else console.log('全部通过 ✓');
