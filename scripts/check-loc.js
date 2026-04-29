#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

const repoRoot = path.resolve(__dirname, '..');
const locDir = path.join(repoRoot, 'src', 'webparts', 'guestSponsorInfo', 'loc');
const referenceFile = 'en-us.js';

function getLocaleFiles() {
  return fs.readdirSync(locDir)
    .filter((fileName) => fileName.endsWith('.js') && fileName !== 'mystrings.d.ts')
    .sort();
}

function readKeys(fileName) {
  const filePath = path.join(locDir, fileName);
  const source = fs.readFileSync(filePath, 'utf8');
  const propertyPattern = /^\s*(?:"([^"\\]+)"|'([^'\\]+)'|([A-Za-z_$][\w$]*)):\s*/gm;
  return Array.from(
    source.matchAll(propertyPattern),
    (match) => match[1] ?? match[2] ?? match[3]
  );
}

function findDuplicates(keys) {
  const counts = new Map();
  const duplicates = [];

  for (const key of keys) {
    const nextCount = (counts.get(key) ?? 0) + 1;
    counts.set(key, nextCount);
    if (nextCount === 2) {
      duplicates.push(key);
    }
  }

  return duplicates;
}

function main() {
  const files = getLocaleFiles();
  const referenceKeys = readKeys(referenceFile);

  if (!referenceKeys.length) {
    console.error(`Could not parse any keys from ${referenceFile}.`);
    process.exitCode = 1;
    return;
  }

  let hasProblems = false;

  for (const fileName of files) {
    const keys = readKeys(fileName);
    const duplicates = findDuplicates(keys);
    const missing = referenceKeys.filter((key) => !keys.includes(key));
    const extra = keys.filter((key) => !referenceKeys.includes(key));
    const firstOrderDiff = referenceKeys.findIndex((key, index) => keys[index] !== key);

    if (!duplicates.length && !missing.length && !extra.length && firstOrderDiff === -1) {
      continue;
    }

    hasProblems = true;
    console.error(`Locale file ${fileName} is inconsistent with ${referenceFile}:`);

    if (duplicates.length) {
      console.error(`  Duplicate keys: ${duplicates.join(', ')}`);
    }

    if (missing.length) {
      console.error(`  Missing keys: ${missing.join(', ')}`);
    }

    if (extra.length) {
      console.error(`  Extra keys: ${extra.join(', ')}`);
    }

    if (firstOrderDiff !== -1) {
      const expected = referenceKeys[firstOrderDiff] ?? '<missing>';
      const actual = keys[firstOrderDiff] ?? '<missing>';
      console.error(`  First key-order mismatch at position ${firstOrderDiff + 1}: expected ${expected}, found ${actual}`);
    }
  }

  if (hasProblems) {
    process.exitCode = 1;
    return;
  }

  console.log('Locale files match the en-us.js key set and order.');
}

main();
