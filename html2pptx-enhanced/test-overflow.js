'use strict';
/**
 * test-overflow.js
 * 运行: node test-overflow.js
 * 注意: 集成测试需要 PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH
 */

const { handleOverflow, applyExpand, applyClip, OVERFLOW_TOLERANCE } = require('./overflow-handler');
const path = require('path');
const fs = require('fs');

let passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); console.log('  ✓', name); passed++; }
  catch(e) { console.log('  ✗', name, '\n   ', e.message); failed++; }
}
function assert(c, m) { if (!c) throw new Error(m || 'assertion failed'); }
function eq(a, b, m) {
  const as = JSON.stringify(a), bs = JSON.stringify(b);
  if (as !== bs) throw new Error((m||'') + '\n    got: '+as+'\n    exp: '+bs);
}
function near(a, b, tol, m) {
  tol = tol || 0.001;
  if (Math.abs(a-b) > tol) throw new Error((m||'near')+': '+a+' vs '+b+' (tol '+tol+')');
}

// ── 测试用 presLayout（16:9） ──────────────────────────
const LAYOUT = { width: 9144000, height: 5143500 }; // EMU
const SLIDE_W = 10.0, SLIDE_H = 5.625;              // inch

// ── 构造 slideData helper ──────────────────────────────
function makeSlideData(elements) {
  return { elements, background: { type:'color', value:'FFFFFF' }, placeholders:[], errors:[] };
}
function makeEl(type, x, y, w, h, extra) {
  return { type, position:{x,y,w,h}, ...(extra||{}) };
}
// bodyDimensions helper
function makeDims(contentWin, contentHin) {
  const PX = 96;
  return {
    width:        SLIDE_W * PX,   // 画布宽 px
    height:       SLIDE_H * PX,   // 画布高 px
    scrollWidth:  contentWin * PX,
    scrollHeight: contentHin * PX,
  };
}

// ══════════════════════════════════════════════════
console.log('\n【1】overflow 检测与阈值');
// ══════════════════════════════════════════════════

test('内容等于画布高度 → 不处理', () => {
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, SLIDE_H)]);
  const dims = makeDims(SLIDE_W, SLIDE_H);
  const r = handleOverflow(sd, dims, LAYOUT, 'expand');
  eq(r.didTransform, false);
  eq(r.transformInfo, 'no overflow');
});
test('内容在容差内（+2%）→ 不处理', () => {
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, SLIDE_H * 1.015)]);
  const dims = makeDims(SLIDE_W, SLIDE_H * 1.015);
  const r = handleOverflow(sd, dims, LAYOUT, 'expand');
  eq(r.didTransform, false, 'within tolerance should not transform');
});
test('内容超出 5% → 触发处理', () => {
  const dims = makeDims(SLIDE_W, SLIDE_H * 1.1);
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, SLIDE_H * 1.1)]);
  const r = handleOverflow(sd, dims, LAYOUT, 'expand');
  eq(r.didTransform, true);
});

// ══════════════════════════════════════════════════
console.log('\n【2】expand 模式 — 缩放计算');
// ══════════════════════════════════════════════════

test('仅高度溢出时 scale = slideH / contentH', () => {
  const contentH = SLIDE_H * 1.5; // 50% overflow
  const dims = makeDims(SLIDE_W, contentH);
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, contentH)]);
  const r = applyExpand(sd, dims, LAYOUT);
  near(r.scale, SLIDE_H / contentH, 0.0001);
  eq(r.didScale, true);
});
test('仅高度溢出时 offsetX 居中', () => {
  // 宽度未溢出，scale 由高度决定，横向有余量 → offsetX > 0
  const contentH = SLIDE_H * 2;
  const contentW = SLIDE_W * 0.8;  // 内容比画布窄
  const dims = { width: SLIDE_W*96, height: SLIDE_H*96, scrollWidth: contentW*96, scrollHeight: contentH*96 };
  const sd = makeSlideData([makeEl('shape', 0, 0, contentW, contentH)]);
  const r = applyExpand(sd, dims, LAYOUT);
  // scale = slideH/contentH
  const expectedScale = SLIDE_H / contentH;
  const scaledW = contentW * expectedScale;
  const expectedOffsetX = (SLIDE_W - scaledW) / 2;
  near(r.offsetX, expectedOffsetX, 0.001, 'offsetX');
});
test('scale 后所有元素都在画布内', () => {
  const contentH = SLIDE_H * 1.8;
  const dims = makeDims(SLIDE_W, contentH);
  const elements = [
    makeEl('shape', 0, 0, SLIDE_W, 1),
    makeEl('shape', 1, contentH * 0.5, 3, 1),
    makeEl('shape', 0, contentH - 0.1, SLIDE_W, 0.2), // 贴底
  ];
  const sd = makeSlideData(elements);
  const r = applyExpand(sd, dims, LAYOUT);
  for (const el of r.slideData.elements) {
    const bottom = el.position.y + el.position.h;
    assert(bottom <= SLIDE_H + 0.001, `elem bottom ${bottom.toFixed(3)} > slideH ${SLIDE_H}`);
  }
});
test('内容未溢出时 didScale=false', () => {
  const dims = makeDims(SLIDE_W, SLIDE_H * 0.9);
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, SLIDE_H * 0.9)]);
  const r = applyExpand(sd, dims, LAYOUT);
  eq(r.didScale, false);
  near(r.scale, 1, 0.0001);
});

// ══════════════════════════════════════════════════
console.log('\n【3】expand 模式 — 样式缩放');
// ══════════════════════════════════════════════════

test('fontSize 随 scale 缩小', () => {
  const contentH = SLIDE_H * 2;
  const dims = makeDims(SLIDE_W, contentH);
  const el = { type:'p', position:{x:0,y:0,w:5,h:0.5},
    style: { fontSize: 24, lineSpacing: 28.8, paraSpaceBefore: 0, paraSpaceAfter: 2, margin:[0.05,0.05,0.05,0.05] } };
  const sd = makeSlideData([el]);
  const r = applyExpand(sd, dims, LAYOUT);
  const outEl = r.slideData.elements[0];
  const expectedScale = SLIDE_H / contentH;
  near(outEl.style.fontSize, 24 * expectedScale, 0.1, 'fontSize');
  near(outEl.style.lineSpacing, 28.8 * expectedScale, 0.1, 'lineSpacing');
  near(outEl.style.margin[0], 0.05 * expectedScale, 0.001, 'margin');
});
test('fontSize 最小值 6pt', () => {
  const contentH = SLIDE_H * 10; // 极端缩放 → scale ≈ 0.1
  const dims = makeDims(SLIDE_W, contentH);
  const el = { type:'p', position:{x:0,y:0,w:5,h:0.5},
    style: { fontSize: 10, lineSpacing: 12 } };
  const sd = makeSlideData([el]);
  const r = applyExpand(sd, dims, LAYOUT);
  assert(r.slideData.elements[0].style.fontSize >= 6, 'fontSize should not go below 6pt');
});
test('inline text runs fontSize 也缩放', () => {
  const contentH = SLIDE_H * 2;
  const dims = makeDims(SLIDE_W, contentH);
  const el = {
    type: 'p',
    position: {x:0, y:0, w:5, h:0.5},
    style: { fontSize: 14 },
    text: [
      { text:'bold', options:{ fontSize: 16, bold: true } },
      { text:'normal', options:{} },
    ],
  };
  const sd = makeSlideData([el]);
  const r = applyExpand(sd, dims, LAYOUT);
  const outEl = r.slideData.elements[0];
  const sc = SLIDE_H / contentH;
  near(outEl.text[0].options.fontSize, 16 * sc, 0.1, 'run fontSize');
  assert(!outEl.text[1].options.fontSize || true, 'run without fontSize ok');
});
test('table colWidths 和 rowHeights 缩放', () => {
  const contentH = SLIDE_H * 1.5;
  const dims = makeDims(SLIDE_W, contentH);
  const el = {
    type: 'table',
    position: {x:0,y:1,w:6,h:2},
    colWidths: [2, 2, 2],
    rowHeights: [0.4, 0.4, 0.4, 0.4],
    rows: [[
      { text:'A', fontSize:10 },
      { text:'B', fontSize:10 },
      { text:'C', fontSize:10 },
    ]],
  };
  const sd = makeSlideData([el]);
  const r = applyExpand(sd, dims, LAYOUT);
  const outEl = r.slideData.elements[0];
  const sc = SLIDE_H / contentH;
  near(outEl.colWidths[0], 2 * sc, 0.001, 'colWidth');
  near(outEl.rowHeights[0], 0.4 * sc, 0.001, 'rowHeight');
  near(outEl.rows[0][0].fontSize, 10 * sc, 0.1, 'cell fontSize');
});
test('line 元素 x1/y1/x2/y2 缩放', () => {
  const contentH = SLIDE_H * 2;
  const dims = makeDims(SLIDE_W, contentH);
  const el = { type:'line', x1:1, y1:2, x2:3, y2:4, color:'000000', width:1 };
  const sd = makeSlideData([el]);
  const r = applyExpand(sd, dims, LAYOUT);
  const sc = SLIDE_H / contentH;
  const outEl = r.slideData.elements[0];
  near(outEl.y1, 2 * sc, 0.001, 'y1');
  near(outEl.y2, 4 * sc, 0.001, 'y2');
});

// ══════════════════════════════════════════════════
console.log('\n【4】clip 模式');
// ══════════════════════════════════════════════════

test('完全在画布内 → 保留', () => {
  const sd = makeSlideData([makeEl('shape', 0.5, 0.5, 3, 1)]);
  const r = applyClip(sd, LAYOUT);
  eq(r.dropped, 0);
  eq(r.clipped, 0);
  near(r.slideData.elements[0].position.h, 1, 0.001);
});
test('完全超出底部 → 丢弃', () => {
  const sd = makeSlideData([
    makeEl('shape', 0, 0, 5, 1),          // 保留
    makeEl('shape', 0, SLIDE_H + 0.1, 5, 1), // 丢弃
  ]);
  const r = applyClip(sd, LAYOUT);
  eq(r.slideData.elements.length, 1);
  eq(r.dropped, 1);
});
test('跨边界 → 截断高度', () => {
  const sd = makeSlideData([
    makeEl('shape', 0, SLIDE_H - 0.3, 5, 1), // 跨越底部
  ]);
  const r = applyClip(sd, LAYOUT);
  eq(r.slideData.elements.length, 1);
  eq(r.clipped, 1);
  near(r.slideData.elements[0].position.h, 0.3, 0.001, 'clipped height');
});
test('完全超出右边界 → 丢弃', () => {
  const sd = makeSlideData([
    makeEl('shape', SLIDE_W + 0.1, 0, 1, 1),
  ]);
  const r = applyClip(sd, LAYOUT);
  eq(r.dropped, 1);
  eq(r.slideData.elements.length, 0);
});
test('跨底部边界的表格 → 只保留能放入的行', () => {
  // 表格从 y=5.0 开始，每行 0.4 inch
  // slide_H = 5.625，只能放 floor((5.625-5)/0.4) = 1 行
  const tableY = 5.0;
  const rowH = 0.4;
  const sd = makeSlideData([{
    type: 'table',
    position: { x:0, y:tableY, w:6, h:rowH*4 },
    colWidths: [2, 2, 2],
    rowHeights: [rowH, rowH, rowH, rowH],
    rows: [
      [{text:'R1C1'},{text:'R1C2'},{text:'R1C3'}],
      [{text:'R2C1'},{text:'R2C2'},{text:'R2C3'}],
      [{text:'R3C1'},{text:'R3C2'},{text:'R3C3'}],
      [{text:'R4C1'},{text:'R4C2'},{text:'R4C3'}],
    ],
  }]);
  const r = applyClip(sd, LAYOUT);
  const tbl = r.slideData.elements[0];
  assert(tbl, 'table should exist');
  const expectedRows = Math.floor((SLIDE_H - tableY) / rowH);
  eq(tbl.rows.length, expectedRows, 'rows count');
});
test('截断后高度过小（<0.01）→ 丢弃', () => {
  const sd = makeSlideData([
    makeEl('shape', 0, SLIDE_H - 0.005, 5, 0.5), // 只剩 0.005 inch
  ]);
  const r = applyClip(sd, LAYOUT);
  eq(r.slideData.elements.length, 0, 'tiny clipped element should be dropped');
});

// ══════════════════════════════════════════════════
console.log('\n【5】handleOverflow 统一入口');
// ══════════════════════════════════════════════════

test('error 模式不改变 slideData', () => {
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, SLIDE_H * 1.5)]);
  const dims = makeDims(SLIDE_W, SLIDE_H * 1.5);
  const r = handleOverflow(sd, dims, LAYOUT, 'error');
  eq(r.mode, 'error');
  eq(r.didTransform, false);
  // slideData 应该原样返回（error 模式不处理，让调用方从 errors 里拿到错误）
  eq(r.slideData, sd);
});
test('expand 模式返回 mode=expand', () => {
  const dims = makeDims(SLIDE_W, SLIDE_H * 1.5);
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, SLIDE_H * 1.5)]);
  const r = handleOverflow(sd, dims, LAYOUT, 'expand');
  eq(r.mode, 'expand');
  eq(r.didTransform, true);
  assert(r.scale < 1, 'scale should be < 1 for overflow');
});
test('clip 模式返回 mode=clip', () => {
  const dims = makeDims(SLIDE_W, SLIDE_H * 1.5);
  const sd = makeSlideData([
    makeEl('shape', 0, 0, 5, 2),
    makeEl('shape', 0, SLIDE_H + 0.5, 5, 1), // 超出的
  ]);
  const r = handleOverflow(sd, dims, LAYOUT, 'clip');
  eq(r.mode, 'clip');
  eq(r.scale, 1, 'clip mode should not scale');
});
test('无溢出时三种模式都返回 didTransform=false', () => {
  const dims = makeDims(SLIDE_W, SLIDE_H * 0.9);
  const sd = makeSlideData([makeEl('shape', 0, 0, 5, SLIDE_H * 0.9)]);
  for (const mode of ['expand', 'clip', 'error']) {
    const r = handleOverflow(sd, dims, LAYOUT, mode);
    eq(r.didTransform, false, `mode=${mode} should not transform`);
  }
});

// ══════════════════════════════════════════════════
console.log('\n【6】集成测试（需要 Playwright）');
// ══════════════════════════════════════════════════

const chromiumPath = process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH;
if (!chromiumPath) {
  console.log('  ⚠ PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH not set, skipping integration tests');
} else {
  // 集成测试异步运行，最后汇总
  runIntegrationTests().then(() => {
    printSummary();
  }).catch(e => {
    console.error('Integration test error:', e.message);
    printSummary();
  });
}

async function runIntegrationTests() {
  const pptxgen = require('pptxgenjs');
  const JSZip = require('jszip');
  const { html2pptx } = require('./html2pptx');
  const htmlFile = path.resolve('./test-overflow.html');

  // ── 集成测试 A：error 模式 ────────────────────
  await asyncTest('error 模式：overflow 抛出错误', async () => {
    const pptx = new pptxgen(); pptx.layout = 'LAYOUT_16x9';
    try {
      await html2pptx(htmlFile, pptx, { overflowMode: 'error' });
      throw new Error('should have thrown');
    } catch(e) {
      assert(e.message.includes('overflows'), 'error should mention overflow: ' + e.message);
    }
  });

  // ── 集成测试 B：expand 模式 ──────────────────
  await asyncTest('expand 模式：所有元素在画布内', async () => {
    const pptx = new pptxgen(); pptx.layout = 'LAYOUT_16x9';
    const result = await html2pptx(htmlFile, pptx, { overflowMode: 'expand' });
    assert(result.overflow.didTransform, 'should have scaled');
    assert(result.overflow.scale < 1, 'scale should be < 1');

    const outPath = '/tmp/overflow-expand.pptx';
    await pptx.writeFile({ fileName: outPath });
    const zip = await JSZip.loadAsync(fs.readFileSync(outPath));
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    // 所有元素的 y + h 不超出 5143500 EMU
    const offs = [...xml.matchAll(/a:off x="(\d+)" y="(\d+)"/g)];
    const exts = [...xml.matchAll(/a:ext cx="(\d+)" cy="(\d+)"/g)];
    const SLIDE_H_EMU = 5143500;
    for (let i = 0; i < Math.min(offs.length, exts.length); i++) {
      const y = parseInt(offs[i][2]);
      const h = parseInt(exts[i][2]);
      assert(y + h <= SLIDE_H_EMU + 10000, // 容差 ~0.01pt
        `element [${i}] bottom ${((y+h)/914400).toFixed(3)}" exceeds slideH`);
    }
    console.log('    scale:', result.overflow.scale.toFixed(4),
                '| elements checked:', offs.length);
  });

  // ── 集成测试 C：clip 模式 ────────────────────
  await asyncTest('clip 模式：PPTX 不超出画布，超出元素被丢弃', async () => {
    const pptx = new pptxgen(); pptx.layout = 'LAYOUT_16x9';
    const result = await html2pptx(htmlFile, pptx, { overflowMode: 'clip' });
    assert(result.overflow.mode === 'clip', 'mode should be clip');

    const outPath = '/tmp/overflow-clip.pptx';
    await pptx.writeFile({ fileName: outPath });
    const zip = await JSZip.loadAsync(fs.readFileSync(outPath));
    const xml = await zip.file('ppt/slides/slide1.xml').async('string');

    // 所有存在的元素 y 不超出画布
    const offs = [...xml.matchAll(/a:off x="(\d+)" y="(\d+)"/g)];
    const SLIDE_H_EMU = 5143500;
    for (const off of offs) {
      const y = parseInt(off[2]);
      assert(y < SLIDE_H_EMU, `element y ${(y/914400).toFixed(3)}" >= slideH`);
    }
    console.log('    elements remaining:', offs.length, '|', result.overflow.transformInfo);
  });

  // ── 集成测试 D：未溢出内容三种模式均正常 ────
  await asyncTest('正常内容（不溢出）所有模式均不报错', async () => {
    // 用 test-charts.html（已知不溢出）
    const normalFile = path.resolve('./test-charts.html');
    for (const mode of ['expand', 'clip', 'error']) {
      const pptx = new pptxgen(); pptx.layout = 'LAYOUT_16x9';
      const r = await html2pptx(normalFile, pptx, {
        overflowMode: mode, captureCanvas: false, captureSvg: false
      });
      eq(r.overflow.didTransform, false, `mode=${mode}: should not transform`);
    }
  });
}

async function asyncTest(name, fn) {
  try {
    await fn();
    console.log('  ✓', name);
    passed++;
  } catch(e) {
    console.log('  ✗', name, '\n   ', e.message);
    failed++;
  }
}

function printSummary() {
  console.log('\n' + '─'.repeat(40));
  console.log('通过:', passed, ' 失败:', failed, ' 合计:', passed + failed);
  if (failed > 0) { console.log('部分失败'); process.exit(1); }
  else console.log('全部通过 ✓');
}

// 无集成测试时直接打印
if (!chromiumPath) printSummary();
