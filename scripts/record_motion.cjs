#!/usr/bin/env node
/**
 * HTML 모션(CSS 애니메이션) → MP4  (PPTX 에 넣을 영상 만들기)
 *
 *   node record_motion.cjs deck.html --slide 1 --selector ".ring" --out assets/ring.mp4 [--seconds 10] [--fps 30] [--poster assets/ring.png]
 *
 *   - 애니메이션을 프레임마다 멈춰 놓고 한 장씩 캡처 → ffmpeg 로 H.264 MP4 (PowerPoint 호환: yuv420p · 짝수 크기 · faststart)
 *   - 영상 크기 = 요소가 장표 안에서 보이는 영역 (장표 밖으로 나간 부분은 잘림)
 *   - 영상 자리는 HTML 에 <div data-video="assets/ring.mp4" data-poster="assets/ring.png"> 로 표시 → html2pptx 가 PowerPoint 동영상으로 넣음
 *   - --slide 는 1부터. 장표가 숨겨져 있는 뷰어형 덱이면 해당 장표만 보이게 처리함
 */
const fs = require('fs');
const path = require('path');
const { execSync, execFileSync } = require('child_process');
if (!process['env'].NODE_PATH) { process['env'].NODE_PATH = execSync('npm root -g', { encoding: 'utf8' }).trim(); require('module').Module._initPaths(); }
const { chromium } = require('playwright');

(async () => {
  const a = process.argv.slice(2);
  const opt = (k, d) => { const i = a.indexOf(k); return i >= 0 ? a[i + 1] : d; };
  const html = a.find((x) => !x.startsWith('--') && /\.html?$/.test(x));
  const sel = opt('--selector'), out = opt('--out');
  const slideNo = parseInt(opt('--slide', '1'), 10), secs = parseFloat(opt('--seconds', '8')), fps = parseInt(opt('--fps', '30'), 10);
  if (!html || !sel || !out) { console.error('사용법: node record_motion.cjs deck.html --slide N --selector ".css" --out x.mp4 [--seconds 8] [--fps 30] [--poster x.png]'); process.exit(1); }

  const browser = await chromium.launch();
  const page = await (await browser.newContext({ viewport: { width: 1400, height: 800 }, deviceScaleFactor: 1 })).newPage();
  await page.goto('file://' + path.resolve(html), { waitUntil: 'load' });
  await page.evaluate(() => document.fonts.ready);

  // 대상 장표만 보이게 (뷰어형 덱: is-active)
  await page.evaluate((n) => {
    [...document.querySelectorAll('.slide')].forEach((s, i) => s.classList.toggle('is-active', i === n - 1));
  }, slideNo);
  const clip = await page.evaluate(({ n, sel }) => {
    const root = document.querySelectorAll('.slide')[n - 1], R = root.getBoundingClientRect();
    const r = root.querySelector(sel).getBoundingClientRect();
    const x0 = Math.max(r.left, R.left), y0 = Math.max(r.top, R.top), x1 = Math.min(r.right, R.right), y1 = Math.min(r.bottom, R.bottom);
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }, { n: slideNo, sel });
  const W = Math.floor(clip.w / 2) * 2, H = Math.floor(clip.h / 2) * 2;
  const box = { x: clip.x, y: clip.y, width: W, height: H };

  const tmp = fs.mkdtempSync(path.join(process['env'].TMPDIR || '/tmp', 'motion-'));
  const frames = Math.round(secs * fps);
  for (let f = 0; f < frames; f++) {
    // 모든 CSS 애니메이션을 같은 시각으로 멈춤 → 프레임이 흔들리지 않고 루프도 정확함
    await page.evaluate((t) => document.getAnimations().forEach((an) => { an.pause(); an.currentTime = t; }), (f / fps) * 1000);
    await page.screenshot({ path: path.join(tmp, `f${String(f).padStart(5, '0')}.png`), clip: box, fullPage: true });
  }
  if (opt('--poster')) {
    await page.evaluate(() => document.getAnimations().forEach((an) => { an.pause(); an.currentTime = 0; }));
    fs.mkdirSync(path.dirname(path.resolve(opt('--poster'))), { recursive: true });
    await page.screenshot({ path: path.resolve(opt('--poster')), clip: box, fullPage: true });
  }
  await browser.close();

  fs.mkdirSync(path.dirname(path.resolve(out)), { recursive: true });
  execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-framerate', String(fps), '-i', path.join(tmp, 'f%05d.png'),
    '-c:v', 'libx264', '-pix_fmt', 'yuv420p', '-crf', '18', '-movflags', '+faststart', path.resolve(out)]);
  fs.rmSync(tmp, { recursive: true, force: true });
  console.log(`${out} — ${W}×${H}px · ${secs}s · ${fps}fps · ${(fs.statSync(out).size / 1048576).toFixed(2)}MB`);
})().catch((e) => { console.error(e.message); process.exit(1); });
