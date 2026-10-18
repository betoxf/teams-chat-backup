import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import test from 'node:test';

const src = fs.readFileSync(new URL('../viewer.js', import.meta.url), 'utf8');
const table = src.match(/const PDF_EMOJI_NAMES = \{[\s\S]*?\n {4}\};/)[0];
const start = src.indexOf('function toPdfText(');
let depth = 0, end = src.indexOf('{', start);
for (; end < src.length; end++) {
    if (src[end] === '{') depth++;
    else if (src[end] === '}' && --depth === 0) break;
}
const context = vm.createContext({});
vm.runInContext(table + '\n' + src.slice(start, end + 1) + '\nthis.toPdfText = toPdfText;', context);
const { toPdfText } = context;
const winAnsiOnly = /^[\t\n\r\x20-\x7E -ÿŒœŠšŸŽžƒˆ˜–—‘’‚“”„†‡•…‰‹›€™]*$/;

test('PDF text names emoji instead of printing garbage', () => {
    assert.equal(toPdfText('below 👇 ok ✅'), 'below [point down] ok [check]');
    assert.equal(toPdfText('❤️ 👍🏽 👨‍👩‍👧 🇺🇸 1️⃣ 🦩'), '[heart] [thumbs up] [family] [flag US] [1] [emoji]');
    assert.equal(toPdfText('Team 🚀', true), 'Team ');
});

test('PDF text keeps Latin-1 and typographic marks, replaces the rest', () => {
    assert.equal(toPdfText('café — “hi” 100€ ™ ©'), 'café — “hi” 100€ ™ ©');
    assert.equal(toPdfText('★ → ←'), '* -> <-');
    assert.equal(toPdfText('안녕'), '??');
    assert.equal(toPdfText(null), '');
});

test('PDF text output always fits the built-in font', () => {
    const sample = '👋🏿 ✌️ ☝🏻 ⚠️ 🧑‍💻 ★☆ 中文 Ω ∑ ​﻿';
    assert.match(toPdfText(sample), winAnsiOnly);
});
