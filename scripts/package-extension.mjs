#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(fs.readFileSync(path.join(repoRoot, 'manifest.json'), 'utf8'));
const outputPath = path.resolve(
  process.argv[2] || path.join(os.homedir(), 'Downloads', `teams-chat-backup-extension-v${manifest.version}.zip`)
);

const entries = [
  'manifest.json',
  'background.js',
  'content.js',
  'db.js',
  'license.js',
  'popup.html',
  'popup.js',
  'styles.css',
  'popup-brand.css',
  'viewer.html',
  'viewer.css',
  'viewer.js',
  'icons',
  'fonts',
  '_locales',
  'libraries'
];

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
if (fs.existsSync(outputPath)) fs.rmSync(outputPath);

const result = spawnSync('zip', ['-qr', outputPath, ...entries], {
  cwd: repoRoot,
  stdio: 'inherit'
});

if (result.status !== 0) {
  process.exit(result.status || 1);
}

console.log(outputPath);
