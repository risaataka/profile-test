// 開発用: dev/icon.html から icons/*.png を書き出す。
//   node dev/make-icons.mjs   (playwright が必要)
import { chromium } from 'playwright';
import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const out = resolve(here, '../icons');
mkdirSync(out, { recursive: true });

const exe = process.env.CHROMIUM_PATH;   // 必要なら実行ファイルを指定
const b = await chromium.launch(exe ? { executablePath: exe } : {});
const p = await b.newPage();
await p.goto('file://' + resolve(here, 'icon.html'));
for (const size of [16, 32, 48, 128]) {
  await p.evaluate((s) => {
    const el = document.getElementById('box');
    el.style.width = el.style.height = s + 'px';
  }, size);
  await p.locator('#box').screenshot({ path: `${out}/icon${size}.png`, omitBackground: true });
  console.log('wrote icon' + size + '.png');
}
await b.close();
