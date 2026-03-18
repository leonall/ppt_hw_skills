'use strict';
/**
 * test-screenshot-logic.js
 * 纯逻辑测试，不依赖浏览器
 * 运行: node test-screenshot-logic.js
 */

let passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); console.log('  ✓', name); passed++; }
  catch (e) { console.log('  ✗', name, '\n   ', e.message); failed++; }
}
function assert(c, m) { if (!c) throw new Error(m || 'failed'); }
function eq(a, b, m) {
  const as = JSON.stringify(a), bs = JSON.stringify(b);
  if (as !== bs) throw new Error((m || '') + '\n    got: ' + as + '\n    exp: ' + bs);
}
function near(a, b, tol, m) {
  if (Math.abs(a - b) > (tol || 1)) throw new Error((m || '') + ': ' + a + ' vs ' + b);
}

// ── 从 element-screenshot.js 提取纯函数逻辑 ────────────────

const PX_PER_IN = 96;
const pxToInch = px => px / PX_PER_IN;

// BROWSER_COLLECT_SCRIPT 的逻辑（Node.js 版，用于测试数据构造）
function mockCollect(elements) {
  // elements: [{ tag, id, rect:{x,y,width,height}, dataset:{} }]
  const PX = 96;
  return elements
    .filter(el => el.rect.width >= 10 && el.rect.height >= 10)
    .map((el, idx) => ({
      tag:        el.tag,
      idx,
      objectName: el.id || `${el.tag.toLowerCase()}-${idx}`,
      selector:   el.id ? `#${el.id}` : `${el.tag.toLowerCase()}:nth-of-type(${idx + 1})`,
      rect:       el.rect,
      position: {
        x: el.rect.x / PX,
        y: el.rect.y / PX,
        w: el.rect.width / PX,
        h: el.rect.height / PX,
      },
      waitMode:    el.dataset?.h2pWait || 'auto',
      customDelay: parseInt(el.dataset?.h2pDelay || '0'),
    }));
}

// clip 计算逻辑
function computeClip(rect, padding, scale) {
  const clip = {
    x: Math.max(0, rect.x - padding),
    y: Math.max(0, rect.y - padding),
    width:  rect.width  + padding * 2,
    height: rect.height + padding * 2,
  };
  // 降采样后的目标尺寸
  const targetW = Math.round(rect.width);
  const targetH = Math.round(rect.height);
  // sharp.extract 参数（从 scale 倍图上裁掉 padding）
  const extract = {
    left:   Math.round(padding * scale),
    top:    Math.round(padding * scale),
    width:  Math.round(rect.width  * scale),
    height: Math.round(rect.height * scale),
  };
  return { clip, targetW, targetH, extract };
}

// 等待模式解析
function resolveWaitMode(info) {
  if (info.waitMode === 'flag') return 'window._h2p_ready';
  if (info.waitMode && info.waitMode.startsWith('delay:')) {
    return 'fixed:' + info.waitMode.split(':')[1];
  }
  if (info.tag === 'CANVAS') return 'pixel-detect';
  return 'raf'; // requestAnimationFrame
}

// ══════════════════════════════════════════════════════
console.log('\n【1】元素收集过滤');
// ══════════════════════════════════════════════════════

test('跳过尺寸 < 10px 的元素', () => {
  const els = mockCollect([
    { tag: 'CANVAS', id: 'big', rect: { x: 0, y: 0, width: 300, height: 200 }, dataset: {} },
    { tag: 'SVG',    id: 'tiny', rect: { x: 0, y: 0, width: 5, height: 5 }, dataset: {} },
    { tag: 'CANVAS', id: 'zero', rect: { x: 0, y: 0, width: 0, height: 0 }, dataset: {} },
  ]);
  eq(els.length, 1);
  eq(els[0].objectName, 'big');
});

test('id 生成 selector', () => {
  const els = mockCollect([
    { tag: 'CANVAS', id: 'my-chart', rect: { x:0,y:0,width:200,height:100 }, dataset:{} },
    { tag: 'SVG',    id: '',         rect: { x:0,y:0,width:200,height:100 }, dataset:{} },
  ]);
  eq(els[0].selector, '#my-chart');
  assert(els[1].selector.includes('nth-of-type'), 'no-id should use nth-of-type');
});

test('position 转换为英寸', () => {
  const els = mockCollect([
    { tag: 'CANVAS', id: 'c1', rect: { x: 96, y: 192, width: 288, height: 192 }, dataset: {} }
  ]);
  eq(els.length, 1);
  near(els[0].position.x, 1.0, 0.001, 'x');
  near(els[0].position.y, 2.0, 0.001, 'y');
  near(els[0].position.w, 3.0, 0.001, 'w');
  near(els[0].position.h, 2.0, 0.001, 'h');
});

test('data-h2p-wait 属性解析', () => {
  const els = mockCollect([
    { tag: 'CANVAS', id: 'a', rect:{x:0,y:0,width:100,height:100}, dataset:{ h2pWait:'flag' } },
    { tag: 'CANVAS', id: 'b', rect:{x:0,y:0,width:100,height:100}, dataset:{ h2pDelay:'500' } },
  ]);
  eq(els[0].waitMode, 'flag');
  eq(els[1].customDelay, 500);
});

// ══════════════════════════════════════════════════════
console.log('\n【2】截图区域计算');
// ══════════════════════════════════════════════════════

test('基础 clip 计算（padding=2, scale=2）', () => {
  const rect = { x: 100, y: 50, width: 300, height: 200 };
  const { clip, targetW, targetH, extract } = computeClip(rect, 2, 2);
  eq(clip.x, 98);   // 100 - 2
  eq(clip.y, 48);   // 50 - 2
  eq(clip.width,  304); // 300 + 4
  eq(clip.height, 204); // 200 + 4
  eq(targetW, 300);
  eq(targetH, 200);
  eq(extract.left,  4);  // 2 * 2
  eq(extract.top,   4);  // 2 * 2
  eq(extract.width,  600); // 300 * 2
  eq(extract.height, 400); // 200 * 2
});

test('x/y=0 时 clip 不越界', () => {
  const rect = { x: 0, y: 1, width: 200, height: 100 };
  const { clip } = computeClip(rect, 5, 2);
  assert(clip.x >= 0, 'x should not be negative');
  assert(clip.y >= 0, 'y should not be negative');
});

test('scale=1 时 extract 等于原始尺寸', () => {
  const rect = { x: 50, y: 50, width: 200, height: 150 };
  const { extract, targetW, targetH } = computeClip(rect, 0, 1);
  eq(extract.width,  targetW);
  eq(extract.height, targetH);
});

test('2x scale 产生 2x 像素尺寸', () => {
  const rect = { x: 0, y: 0, width: 400, height: 300 };
  const { extract } = computeClip(rect, 0, 2);
  eq(extract.width,  800);
  eq(extract.height, 600);
});

// ══════════════════════════════════════════════════════
console.log('\n【3】等待模式解析');
// ══════════════════════════════════════════════════════

test('data-h2p-wait=flag → window._h2p_ready', () => {
  eq(resolveWaitMode({ tag: 'CANVAS', waitMode: 'flag' }), 'window._h2p_ready');
});

test('CANVAS 默认 → pixel-detect', () => {
  eq(resolveWaitMode({ tag: 'CANVAS', waitMode: 'auto' }), 'pixel-detect');
});

test('SVG 默认 → raf', () => {
  eq(resolveWaitMode({ tag: 'SVG', waitMode: 'auto' }), 'raf');
});

test('delay: 前缀解析', () => {
  eq(resolveWaitMode({ tag: 'CANVAS', waitMode: 'delay:800' }), 'fixed:800');
});

// ══════════════════════════════════════════════════════
console.log('\n【4】skipTags 过滤逻辑');
// ══════════════════════════════════════════════════════

test('skipTags=SVG 只处理 CANVAS', () => {
  const elements = [
    { tag: 'CANVAS', objectName: 'c1' },
    { tag: 'SVG',    objectName: 's1' },
    { tag: 'CANVAS', objectName: 'c2' },
  ];
  const skipTags = new Set(['SVG']);
  const toCapture = elements.filter(e => !skipTags.has(e.tag));
  eq(toCapture.length, 2);
  assert(toCapture.every(e => e.tag === 'CANVAS'));
});

test('skipTags=CANVAS,SVG → 全部跳过', () => {
  const elements = [
    { tag: 'CANVAS', objectName: 'c1' },
    { tag: 'SVG',    objectName: 's1' },
  ];
  const skipTags = new Set(['CANVAS', 'SVG']);
  eq(elements.filter(e => !skipTags.has(e.tag)).length, 0);
});

test('空 skipTags → 全部处理', () => {
  const elements = [
    { tag: 'CANVAS', objectName: 'c1' },
    { tag: 'SVG',    objectName: 's1' },
  ];
  const skipTags = new Set();
  eq(elements.filter(e => !skipTags.has(e.tag)).length, 2);
});

// ══════════════════════════════════════════════════════
console.log('\n【5】dataUrl 格式');
// ══════════════════════════════════════════════════════

test('PNG buffer 转 dataUrl 格式正确', () => {
  // 模拟一个 PNG buffer（最小合法 PNG）
  const fakePng = Buffer.from(
    '89504e470d0a1a0a0000000d49484452000000010000000108020000009001' +
    '2e00000000c4944415478016360f8cfc00000000200015c4d1700000000049454e44ae426082',
    'hex'
  );
  const dataUrl = 'image/png;base64,' + fakePng.toString('base64');
  assert(dataUrl.startsWith('image/png;base64,'), 'should start with mime');
  assert(dataUrl.length > 30, 'should have content');
  // pptxgenjs addImage data 参数格式验证
  assert(!dataUrl.startsWith('data:'), 'pptxgenjs does NOT want data: prefix');
});

// ══════════════════════════════════════════════════════
console.log('\n【6】pptxgenjs 坐标验证');
// ══════════════════════════════════════════════════════

test('position 英寸值在合理范围内（16:9 幻灯片 = 13.33" x 7.5"）', () => {
  // 模拟一个 720pt x 405pt 页面上的图表位置
  // 720pt = 720/72 inch = 10inch in CSS = 960px
  // 位置: x=20pt=26.7px, y=80pt=106.7px, w=280pt=373px, h=160pt=213px
  const PT_TO_PX = 96 / 72;
  const rect = {
    x: 20 * PT_TO_PX,
    y: 80 * PT_TO_PX,
    width:  280 * PT_TO_PX,
    height: 160 * PT_TO_PX,
  };
  const pos = {
    x: rect.x / PX_PER_IN,
    y: rect.y / PX_PER_IN,
    w: rect.width / PX_PER_IN,
    h: rect.height / PX_PER_IN,
  };
  // 16:9 幻灯片: 13.33" x 7.5"
  assert(pos.x >= 0 && pos.x < 13.33, 'x in slide bounds');
  assert(pos.y >= 0 && pos.y < 7.5,   'y in slide bounds');
  assert(pos.w > 0 && pos.x + pos.w <= 13.33, 'width ok');
  assert(pos.h > 0 && pos.y + pos.h <= 7.5,   'height ok');
  // 验证实际数值
  near(pos.x, 20 / 72, 0.01, 'x in inches');
  near(pos.w, 280 / 72, 0.01, 'w in inches');
});

// ══════════════════════════════════════════════════════
console.log('\n' + '─'.repeat(40));
console.log('通过:', passed, ' 失败:', failed, ' 合计:', passed + failed);
if (failed > 0) { console.log('部分失败'); process.exit(1); }
else console.log('全部通过 ✓');
