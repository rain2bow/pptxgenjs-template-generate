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
    console.error('Usage: node scripts/validate-pptx-native.js <deck.pptx>');
    process.exit(2);
  }

  if (!fs.existsSync(file)) {
    console.error(`File not found: ${file}`);
    process.exit(2);
  }

  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b));

  const errors = [];
  const warnings = [];

  if (!slideFiles.length) {
    errors.push('No slides found.');
  }

  for (const slideFile of slideFiles) {
    const xml = await zip.files[slideFile].async('string');
    const n = slideNumber(slideFile);
    const textBodies = count(xml, /<[a-z]+:txBody\b/g);
    const shapeNodes = count(xml, /<p:sp\b/g);
    const pictureNodes = count(xml, /<p:pic\b/g);
    const lineNodes = count(xml, /<a:ln\b/g);

    if (textBodies === 0) {
      errors.push(`Slide ${n}: no native text bodies found. This likely is an image-only slide.`);
    }

    if (pictureNodes === 1 && textBodies === 0 && shapeNodes <= 1) {
      errors.push(`Slide ${n}: appears to be a single-picture conversion slide.`);
    }

    if (pictureNodes > 0 && textBodies < 2 && shapeNodes < 3) {
      warnings.push(`Slide ${n}: has weak native structure (${textBodies} text bodies, ${shapeNodes} shapes, ${pictureNodes} pictures). Confirm this is not a screenshot-based page.`);
    }

    if (textBodies > 0 && shapeNodes + lineNodes < 2) {
      warnings.push(`Slide ${n}: native text exists but geometric structure is sparse. Check visual parity with the original Presentation Team layout.`);
    }
  }

  if (warnings.length) {
    console.warn('Warnings:');
    for (const warning of warnings) console.warn(`- ${warning}`);
  }

  if (errors.length) {
    console.error('Native PPTX validation failed:');
    for (const error of errors) console.error(`- ${error}`);
    process.exit(1);
  }

  console.log(`Native PPTX validation passed: ${slideFiles.length} slide(s).`);

  function count(text, re) {
    return (text.match(re) || []).length;
  }

  function slideNumber(name) {
    return Number(name.match(/slide(\d+)\.xml$/)?.[1] || 0);
  }
}

