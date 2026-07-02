#!/usr/bin/env node
const fs = require('node:fs');
const process = require('node:process');
const JSZip = require('jszip');

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

async function main() {
  const file = process.argv[2];
  if (!file) {
    console.error('Usage: node scripts/validate-pptx-layout.js <deck.pptx>');
    process.exit(2);
  }
  
  if (!fs.existsSync(file)) {
    console.error(`File not found: ${file}`);
    process.exit(2);
  }
  
  const EMU_PER_IN = 914400;
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b));
  
  const errors = [];
  const warnings = [];
  
  for (const slideFile of slideFiles) {
    const xml = await zip.files[slideFile].async('string');
    const slideNo = slideNumber(slideFile);
    const nodes = extractDrawableNodes(xml);
  
    const texts = nodes.filter((node) => node.kind === 'text' && node.text.trim());
    const blockers = nodes.filter((node) => node.kind === 'shape' || node.kind === 'picture');
  
    for (const text of texts) {
      for (const blocker of blockers) {
        if (blocker.order <= text.order) continue;
        if (blocker.transparency >= 85000) continue;
        const ratio = overlapRatio(text.box, blocker.box);
        if (ratio > 0.12) {
          errors.push(
            `Slide ${slideNo}: text "${shorten(text.text)}" is likely covered by later ${blocker.kind} "${blocker.name}" (${Math.round(ratio * 100)}% overlap).`
          );
        }
      }
    }
  
    for (let i = 0; i < texts.length; i += 1) {
      for (let j = i + 1; j < texts.length; j += 1) {
        const a = texts[i];
        const b = texts[j];
        if (a.isTiny || b.isTiny) continue;
        const ratio = Math.min(overlapRatio(a.box, b.box), overlapRatio(b.box, a.box));
        if (ratio > 0.35) {
          warnings.push(
            `Slide ${slideNo}: text boxes may overlap: "${shorten(a.text)}" and "${shorten(b.text)}" (${Math.round(ratio * 100)}%).`
          );
        }
      }
    }
  
    for (const text of texts) {
      const isFooterText = text.box.y > 6.85 && text.box.h <= 0.35;
      if (!isFooterText && text.box.y + text.box.h > 6.9) {
        warnings.push(`Slide ${slideNo}: text "${shorten(text.text)}" enters bottom navigation/footer safe zone.`);
      }
      if (text.box.x < 0.25 || text.box.x + text.box.w > 13.1) {
        warnings.push(`Slide ${slideNo}: text "${shorten(text.text)}" is close to horizontal slide edge.`);
      }
    }
  }
  
  if (warnings.length) {
    console.warn('Warnings:');
    for (const warning of warnings) console.warn(`- ${warning}`);
  }
  
  if (errors.length) {
    console.error('PPTX layout validation failed:');
    for (const error of errors) console.error(`- ${error}`);
    process.exit(1);
  }
  
  console.log(`PPTX layout validation passed: ${slideFiles.length} slide(s).`);
  
  function extractDrawableNodes(xml) {
    const nodes = [];
    const re = /<p:(sp|pic)\b[\s\S]*?<\/p:\1>/g;
    let match;
    let order = 0;
    while ((match = re.exec(xml))) {
      const raw = match[0];
      const tag = match[1];
      const box = parseBox(raw);
      if (!box || box.w <= 0 || box.h <= 0) continue;
  
      const hasText = /<[a-z]+:txBody\b/.test(raw);
      const text = hasText ? stripXmlText(raw) : '';
      const name = raw.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/)?.[1] || `${tag}-${order}`;
      const transparency = parseTransparency(raw);
      const isFilled = /<a:solidFill\b/.test(raw) && !/<a:noFill\s*\/>/.test(raw);
      const hasLineOnly = /<a:prstGeom\b[^>]*prst="line"/.test(raw);
  
      if (hasText) {
        nodes.push({
          kind: 'text',
          order,
          name,
          box,
          text,
          isTiny: box.h < 0.24 || text.length <= 2,
        });
      } else if (tag === 'pic') {
        nodes.push({ kind: 'picture', order, name, box, transparency: 0 });
      } else if (isFilled && !hasLineOnly) {
        nodes.push({ kind: 'shape', order, name, box, transparency });
      }
  
      order += 1;
    }
    return nodes;
  }
  
  function parseBox(raw) {
    const off = raw.match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"/);
    const ext = raw.match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    if (!off || !ext) return null;
    return {
      x: Number(off[1]) / EMU_PER_IN,
      y: Number(off[2]) / EMU_PER_IN,
      w: Number(ext[1]) / EMU_PER_IN,
      h: Number(ext[2]) / EMU_PER_IN,
    };
  }
  
  function parseTransparency(raw) {
    const alpha = raw.match(/<a:alpha\b[^>]*\bval="(\d+)"/);
    if (alpha) return 100000 - Number(alpha[1]);
    const trans = raw.match(/<a:transparency\b[^>]*\bval="(\d+)"/);
    if (trans) return Number(trans[1]);
    return 0;
  }
  
  function stripXmlText(raw) {
    return [...raw.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
      .map((m) => decodeXml(m[1]))
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();
  }
  
  function decodeXml(text) {
    return text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }
  
  function overlapRatio(a, b) {
    const x = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
    const y = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
    const area = x * y;
    const base = Math.max(0.0001, a.w * a.h);
    return area / base;
  }
  
  function shorten(text) {
    return text.length > 24 ? `${text.slice(0, 24)}...` : text;
  }
  
  function slideNumber(name) {
    return Number(name.match(/slide(\d+)\.xml$/)?.[1] || 0);
  }
}
