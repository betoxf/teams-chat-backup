#!/usr/bin/env node
// Checks manifest.json and _locales against Chrome Web Store limits.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'), 'utf8'));
const errors = [];

if (manifest.manifest_version !== 3) errors.push('manifest_version must be 3');
if (manifest.host_permissions?.length) errors.push('host_permissions must stay empty (activeTab only)');

const localesDir = path.join(root, '_locales');
const locales = fs.readdirSync(localesDir).filter(name => !name.startsWith('.'));
if (!locales.includes(manifest.default_locale)) errors.push(`missing default locale ${manifest.default_locale}`);

for (const locale of locales) {
  const messages = JSON.parse(fs.readFileSync(path.join(localesDir, locale, 'messages.json'), 'utf8'));
  const name = messages.extName?.message || '';
  const description = messages.extDescription?.message || '';
  if (!name || name.length > 75) errors.push(`${locale}: extName is ${name.length} chars (1-75)`);
  if (!description || description.length > 132) errors.push(`${locale}: extDescription is ${description.length} chars (1-132)`);
  if (!messages.extShortName?.message) errors.push(`${locale}: missing extShortName`);
}

if (errors.length) {
  console.error(errors.join('\n'));
  process.exit(1);
}
console.log(`manifest ok, ${locales.length} locales ok`);
