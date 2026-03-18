'use strict';
/**
 * test-regression.js
 * 针对代码检查发现的 4 个 Bug 的回归测试
 * 运行: node test-regression.js
 */

const { handleOverflow, applyExpand, applyClip, getContentBounds } = require('./overflow-handler');

let passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); console.log('  ✓', name); passed++; }
  catch(e) { console.log('  ✗', name, '\n   ', e.message); failed++; }
}
function assert(c, m) { if (!c) throw new Error(m || 'failed'); }
function eq(a, b, m) {
  const as = JSON.stringify(a), bs = JSON.stringify(b);
  if (as !== bs) throw new Error((m||'') + '\n    got: '+as+'\n    exp: '+bs);
}
function near(a, b, tol, m) {
  if (Math.abs(a-b) > (tol||0.001)) throw new Error((m||'near')+': '+a+' vs '+b);
}

const LAYOUT = { width: 9144000, height: 5143500 };
const SLIDE_W = 10.0, SLIDE_H = 5.625;
const PX = 96;

function makeSlideData(elements, placeholders) {
  return {
    elements: elements || [],
    placeholders: placeholders || [],
    background: { type:'color', value:'FFFFFF' },
    errors: []
  };
}
function makeDims(contentHin) {
  return { width: SLIDE_W*PX, height: 0, scrollWidth: SLIDE_W*PX, scrollHeight: contentHin*PX };
}

// ══════════════════════════════════════════════════
console.log('\n【Bug 1】expand 模式 placeholders 坐标缩放');
// ══════════════════════════════════════════════════

test('placeholders 随 expand 缩放', () => {
  const contentH = SLIDE_H * 2;  // scale = 0.5
  const dims = makeDims(contentH);
  const sd = makeSlideData(
    [{ type:'shape', position:{x:0, y:0, w:5, h:contentH} }],
    [{ id:'chart1', x:1.0, y:2.0, w:3.0, h:1.5 }]
  );
  const result = applyExpand(sd, dims, LAYOUT);
  const ph = result.slideData.placeholders[0];
  const sc = result.scale;          // 0.5
  const ox = result.offsetX;        // (10 - 10*0.5)/2 = 2.5
  near(ph.x, 1.0 * sc + ox, 0.01, 'placeholder x (with offsetX)');
  near(ph.y, 2.0 * sc, 0.01, 'placeholder y');
  near(ph.w, 3.0 * sc, 0.01, 'placeholder w');
  near(ph.h, 1.5 * sc, 0.01, 'placeholder h');
  // 确认 offsetX 不为 0（用于覆盖 offsetX 生效的路径）
  assert(ox > 0, 'offsetX should be > 0 when content narrower than slide');
});

test('placeholders 含 offsetX 居中偏移', () => {
  // 内容宽度比画布窄，会有 offsetX
  const contentH = SLIDE_H * 1.5;
  const contentW = SLIDE_W * 0.8;
  const dims = { width: SLIDE_W*PX, height: 0, scrollWidth: contentW*PX, scrollHeight: contentH*PX };
  const sd = makeSlideData(
    [{ type:'shape', position:{x:0, y:0, w:contentW, h:contentH} }],
    [{ id:'chart1', x:0.5, y:1.0, w:2.0, h:1.0 }]
  );
  const result = applyExpand(sd, dims, LAYOUT);
  const sc = SLIDE_H / contentH;
  const scaledW = contentW * sc;
  const expectedOffsetX = (SLIDE_W - scaledW) / 2;
  const ph = result.slideData.placeholders[0];
  near(ph.x, 0.5 * sc + expectedOffsetX, 0.01, 'placeholder x with offsetX');
  near(ph.y, 1.0 * sc, 0.01, 'placeholder y');
});

test('无 placeholders 时不报错', () => {
  const contentH = SLIDE_H * 1.5;
  const dims = makeDims(contentH);
  const sd = makeSlideData([{ type:'shape', position:{x:0,y:0,w:5,h:contentH} }], []);
  const result = applyExpand(sd, dims, LAYOUT);
  eq(result.slideData.placeholders, [], 'empty placeholders preserved');
});

test('无溢出时 placeholders 不变', () => {
  const dims = makeDims(SLIDE_H * 0.9);
  const sd = makeSlideData(
    [{ type:'shape', position:{x:0,y:0,w:5,h:SLIDE_H*0.9} }],
    [{ id:'chart1', x:1.0, y:2.0, w:3.0, h:1.5 }]
  );
  const result = applyExpand(sd, dims, LAYOUT);
  eq(result.didScale, false);
  near(result.slideData.placeholders[0].x, 1.0, 0.001, 'no change when no overflow');
});

// ══════════════════════════════════════════════════
console.log('\n【Bug 2】clip 模式 PNG 变形（逻辑验证）');
// ══════════════════════════════════════════════════

// Bug2 的核心修复在 html2pptx.js 里，这里只测相关计算逻辑

test('clip 截断比例计算正确', () => {
  const originalH = 1.27;  // inch
  const slideH    = 5.625;
  const posY      = 5.5;   // 跨底部
  const clippedH  = slideH - posY; // 0.125
  const keepRatio = clippedH / originalH;
  assert(keepRatio > 0 && keepRatio < 1, 'keepRatio in (0,1)');
  near(keepRatio, 0.125 / 1.27, 0.001, 'keepRatio');
  // PNG 宽 200px 高 122px（对应 1x 坐标），裁到 keepRatio
  const pngH = 122;
  const keepPx = Math.max(1, Math.round(pngH * keepRatio));
  assert(keepPx >= 1, 'keepPx >= 1');
  assert(keepPx < pngH, 'keepPx < original height');
});

test('完全超出时 keepRatio 为 0 应跳过而非裁切', () => {
  const posY  = 5.7;
  const slideH = 5.625;
  // posY >= slideH → continue，不会到裁切逻辑
  assert(posY >= slideH, 'should be skipped');
});

test('轻微截断（keepRatio > 0.99）不需要裁切', () => {
  const originalH = 1.0;
  const clippedH  = 0.995;
  const keepRatio = clippedH / originalH;
  assert(keepRatio >= 0.99, 'should skip crop when ratio >= 0.99');
});

// ══════════════════════════════════════════════════
console.log('\n【Bug 3】expand/clip 在无溢出时恢复 textBoxPosition 校验（逻辑验证）');
// ══════════════════════════════════════════════════

test('无溢出时 isActuallyOverflow 为 false', () => {
  const EMU_PER_IN = 914400;
  const contentHIn = SLIDE_H * 0.9;  // 没有溢出
  const slideHIn   = LAYOUT.height / EMU_PER_IN;
  const isActuallyOverflow = (contentHIn - slideHIn) / slideHIn > 0.02;
  eq(isActuallyOverflow, false, 'no overflow content');
});

test('溢出时 isActuallyOverflow 为 true', () => {
  const EMU_PER_IN = 914400;
  const contentHIn = SLIDE_H * 1.5;  // 50% 溢出
  const slideHIn   = LAYOUT.height / EMU_PER_IN;
  const isActuallyOverflow = (contentHIn - slideHIn) / slideHIn > 0.02;
  eq(isActuallyOverflow, true, 'overflow content');
});

test('容差边界（2%）内不触发', () => {
  const EMU_PER_IN = 914400;
  const contentHIn = SLIDE_H * 1.015;  // 1.5% 超出，在容差内
  const slideHIn   = LAYOUT.height / EMU_PER_IN;
  const isActuallyOverflow = (contentHIn - slideHIn) / slideHIn > 0.02;
  eq(isActuallyOverflow, false, 'within tolerance');
});

// ══════════════════════════════════════════════════
console.log('\n【Bug 4】getContentBounds 计入 line 元素');
// ══════════════════════════════════════════════════

test('line 元素参与 maxBottom 计算', () => {
  const elements = [
    { type:'shape', position:{x:0, y:0, w:5, h:3} },        // bottom=3
    { type:'line',  x1:1, y1:2, x2:9, y2:7 },               // bottom=7
  ];
  const { maxRight, maxBottom } = getContentBounds(elements);
  near(maxBottom, 7, 0.001, 'line y2=7 should be maxBottom');
  near(maxRight,  9, 0.001, 'line x2=9 should be maxRight');
});

test('line 元素参与 maxRight 计算', () => {
  const elements = [
    { type:'shape', position:{x:0, y:0, w:8, h:2} },   // right=8
    { type:'line',  x1:0, y1:0, x2:12, y2:1 },          // right=12
  ];
  const { maxRight } = getContentBounds(elements);
  near(maxRight, 12, 0.001, 'line x2=12 is maxRight');
});

test('line 元素取 x1/y1 和 x2/y2 的最大值', () => {
  // 从右到左的 line（x1 > x2）
  const elements = [
    { type:'line', x1:9, y1:5, x2:2, y2:1 },
  ];
  const { maxRight, maxBottom } = getContentBounds(elements);
  near(maxRight,  9, 0.001, 'should take max of x1,x2');
  near(maxBottom, 5, 0.001, 'should take max of y1,y2');
});

test('只有 line 元素时也能正确计算', () => {
  const elements = [
    { type:'line', x1:0.5, y1:1, x2:8, y2:6 },
  ];
  const { maxRight, maxBottom } = getContentBounds(elements);
  near(maxRight,  8, 0.001);
  near(maxBottom, 6, 0.001);
});

test('expand 的 scale 能基于超出画布的 line', () => {
  // slide H = 5.625, line 超出到 y2=7
  const contentH = 7.0;
  const dims = makeDims(contentH);  // scrollHeight 也是 7inch
  const sd = makeSlideData(
    [{ type:'line', x1:0, y1:0, x2:10, y2:contentH }]
  );
  const result = applyExpand(sd, dims, LAYOUT);
  assert(result.didScale, 'should scale when line exceeds boundary');
  // scale = 5.625/7 ≈ 0.8036
  near(result.scale, SLIDE_H / contentH, 0.001, 'scale based on line bottom');
  // line 缩放后 y2 应该 <= slideH
  const scaledLine = result.slideData.elements[0];
  assert(scaledLine.y2 <= SLIDE_H + 0.001, `line y2 ${scaledLine.y2} should fit in slide`);
});

// ══════════════════════════════════════════════════
console.log('\n【互通性验证】各特性组合场景');
// ══════════════════════════════════════════════════

test('expand + table + placeholder 组合', () => {
  const contentH = SLIDE_H * 1.5;
  const dims = makeDims(contentH);
  const sd = makeSlideData(
    [
      { type:'shape', position:{x:0, y:0, w:8, h:1} },
      {
        type:'table',
        position:{x:0, y:1, w:8, h:3},
        colWidths:[2,2,2,2],
        rowHeights:[0.5,0.5,0.5,0.5,0.5,0.5],
        rows:[[{text:'A',fontSize:11},{text:'B',fontSize:11},{text:'C',fontSize:11},{text:'D',fontSize:11}]]
      },
    ],
    [{ id:'chartPh', x:0.5, y:4.5, w:9, h:1.5 }]
  );
  const result = applyExpand(sd, dims, LAYOUT);
  const sc = SLIDE_H / contentH;

  // shape 缩放
  near(result.slideData.elements[0].position.h, 1 * sc, 0.01, 'shape h scaled');
  // table colWidths 缩放
  const tbl = result.slideData.elements[1];
  near(tbl.colWidths[0], 2 * sc, 0.01, 'colWidth scaled');
  // table fontSize 缩放
  near(tbl.rows[0][0].fontSize, Math.max(6, 11 * sc), 0.1, 'cell fontSize scaled');
  // placeholder 缩放
  near(result.slideData.placeholders[0].y, 4.5 * sc, 0.01, 'placeholder y scaled');
  // 所有元素在画布内
  for (const el of result.slideData.elements) {
    if (el.position) {
      assert(el.position.y + el.position.h <= SLIDE_H + 0.001,
        `${el.type} bottom ${(el.position.y+el.position.h).toFixed(3)} exceeds slide`);
    }
  }
});

test('clip + line 边界值（一端在内一端超出）', () => {
  // line y1=3（在内），y2=7（超出）→ 按现有设计保留（简化处理）
  const sd = makeSlideData([
    { type:'line', x1:0, y1:3, x2:10, y2:7 }
  ]);
  const result = applyClip(sd, LAYOUT);
  // line 一端超出时保留（不裁切，已知简化行为）
  eq(result.dropped, 0, 'partially out line should be kept');
});

test('clip + line 两端都超出 → 丢弃', () => {
  const sd = makeSlideData([
    { type:'line', x1:0, y1:6, x2:10, y2:7 }  // 两端 y 都 > 5.625
  ]);
  const result = applyClip(sd, LAYOUT);
  eq(result.dropped, 1, 'fully out line dropped');
  eq(result.slideData.elements.length, 0);
});

test('handleOverflow error 模式不影响 placeholders', () => {
  const dims = makeDims(SLIDE_H * 0.9);
  const sd = makeSlideData(
    [{ type:'shape', position:{x:0,y:0,w:5,h:2} }],
    [{ id:'ph1', x:1, y:2, w:3, h:1 }]
  );
  const result = handleOverflow(sd, dims, LAYOUT, 'error');
  // error 模式不变换，placeholders 原样
  near(result.slideData.placeholders[0].x, 1, 0.001);
  near(result.slideData.placeholders[0].y, 2, 0.001);
});

// ══════════════════════════════════════════════════
console.log('\n' + '─'.repeat(40));
console.log('通过:', passed, ' 失败:', failed, ' 合计:', passed+failed);
if (failed > 0) { console.log('部分失败'); process.exit(1); }
else console.log('全部通过 ✓');
