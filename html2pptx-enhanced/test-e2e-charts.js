'use strict';
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');
const { html2pptx, GradientRegistry, postprocess } = require('./html2pptx');

async function run() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  const registry = new GradientRegistry();

  console.log('Processing test-charts.html...');
  const result = await html2pptx(
    path.resolve('./test-charts.html'),
    pptx,
    { slideIndex: 1, registry, screenshotScale: 2, screenshotTimeout: 8000 }
  );

  console.log('screenshots captured:', result.screenshots.length);
  result.screenshots.forEach(s => {
    const size = s.pngBuffer ? s.pngBuffer.length + 'B' : 'FAIL:' + s.error;
    console.log(
      ' ', s.tag, s.objectName, size,
      '@ x=' + s.position.x.toFixed(2) + '" y=' + s.position.y.toFixed(2) + '"',
      'w=' + s.position.w.toFixed(2) + '" h=' + s.position.h.toFixed(2) + '"'
    );
  });

  const outPath = '/tmp/charts-test.pptx';
  await pptx.writeFile({ fileName: outPath });
  console.log('\nPPTX written:', outPath);

  // 验证：媒体文件 + blip 数量 + 坐标
  const zip = await JSZip.loadAsync(fs.readFileSync(outPath));
  const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('ppt/media/') && !f.endsWith('/'));
  console.log('embedded media:', mediaFiles.map(f => f.split('/').pop()).join(', '));

  const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
  const blipCount = (slideXml.match(/a:blip /g) || []).length;
  console.log('image blips in slide:', blipCount, '(expect 4)');

  // 解析 p:pic 元素的坐标（EMU → inch）
  const picPattern = /<p:pic>[\s\S]*?<\/p:pic>/g;
  const pics = slideXml.match(picPattern) || [];
  console.log('pic elements:', pics.length);
  const EMU = 914400;
  pics.forEach((pic, i) => {
    const off = pic.match(/a:off x="(\d+)" y="(\d+)"/);
    const ext = pic.match(/a:ext cx="(\d+)" cy="(\d+)"/);
    if (off && ext) {
      console.log(
        '  pic[' + i + ']',
        'x=' + (parseInt(off[1]) / EMU).toFixed(2) + '"',
        'y=' + (parseInt(off[2]) / EMU).toFixed(2) + '"',
        'w=' + (parseInt(ext[1]) / EMU).toFixed(2) + '"',
        'h=' + (parseInt(ext[2]) / EMU).toFixed(2) + '"'
      );
    }
  });

  // 断言
  let ok = true;
  if (result.screenshots.length !== 4) { console.log('FAIL: expected 4 screenshots'); ok = false; }
  if (result.screenshots.some(s => !s.pngBuffer)) { console.log('FAIL: some screenshots failed'); ok = false; }
  if (blipCount !== 4) { console.log('FAIL: expected 4 blips, got', blipCount); ok = false; }
  if (mediaFiles.length !== 4) { console.log('FAIL: expected 4 media files'); ok = false; }

  console.log(ok ? '\n✓ All checks passed' : '\n✗ Some checks failed');
  process.exit(ok ? 0 : 1);
}

run().catch(e => { console.error('ERR:', e.message); process.exit(1); });
