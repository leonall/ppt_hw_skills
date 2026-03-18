#!/usr/bin/env node
'use strict';

/**
 * CLI: node html2pptx-cli.js slide1.html slide2.html -o output.pptx
 */

const pptxgen = require('pptxgenjs');
const { html2pptx, GradientRegistry, postprocess } = require('./html2pptx');
const path = require('path');

async function main() {
  const args = process.argv.slice(2);
  const outIdx = args.indexOf('-o');
  const outputFile = outIdx !== -1 ? args[outIdx + 1] : 'output.pptx';
  const htmlFiles = args.filter((a, i) => !a.startsWith('-') && i !== outIdx + 1);

  if (htmlFiles.length === 0) {
    console.error('Usage: node html2pptx-cli.js slide1.html [slide2.html ...] -o output.pptx');
    process.exit(1);
  }

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  const registry = new GradientRegistry();

  for (let i = 0; i < htmlFiles.length; i++) {
    const f = htmlFiles[i];
    console.log(`[${i+1}/${htmlFiles.length}] Processing: ${f}`);
    await html2pptx(f, pptx, { slideIndex: i + 1, registry });
  }

  const absOut = path.resolve(outputFile);
  await pptx.writeFile({ fileName: absOut });
  console.log(`✓ Written: ${absOut}`);

  // 渐变后处理
  if (!registry.isEmpty()) {
    console.log('Applying gradient post-processing...');
    const { processed, warnings } = await postprocess(absOut, registry);
    console.log(`✓ Gradients applied: ${processed}`);
    if (warnings.length > 0) {
      console.warn('Warnings:');
      warnings.forEach(w => console.warn(`  ⚠ ${w}`));
    }
  }

  console.log('Done.');
}

main().catch(err => { console.error('Error:', err.message); process.exit(1); });
