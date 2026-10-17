#!/usr/bin/env node
/**
 * HTML 덱 → 편집 가능한 PPTX (한 번에)
 *
 *   node build.cjs deck.html [out.pptx] [--fontmap fonts/fontmap.json] [--no-embed] [--loose] [--keep] [--ppt] [--no-verify]
 *
 *   1) html2pptx.cjs  : 장표(section.slide)마다 글자 · 도형 · 이미지 → PptxGenJS (LAYOUT_WIDE 13.33×7.5in)
 *   2) native.py      : 표(.gtable) · 차트([data-chart]) → PowerPoint 네이티브, 한글 글꼴(ea) 지정, 글꼴 임베딩
 *   3) verify.py      : 파일 · 무결성(복구 창 패턴) · 장표 수 · 네이티브 수 · 글꼴 임베딩 검사 → 합격/불합격
 *                       --ppt 를 주면 실제 PowerPoint 로 열어 보고 HTML ↔ PowerPoint 비교 이미지까지 (out_compare/)
 *
 *   --fontmap  글꼴 대응표. 생략하면 HTML 의 <meta name="pptx-fontmap" content="…"> 를 쓴다.
 *   --no-embed 글꼴을 파일에 넣지 않음 (용량 작게, 받는 PC 에 글꼴 설치 필요)
 *   --loose    변환 규칙 위반을 오류 대신 경고로 출력
 *   --keep     중간 파일(.base.pptx · .natives.json) 남기기
 *   --ppt      검증에 PowerPoint 렌더 · 비교 포함 (다른 문서가 열려 있으면 PowerPoint 단계는 건너뜀)
 *   --no-verify 검증 생략
 */
const fs = require('fs');
const path = require('path');
const { execSync, execFileSync } = require('child_process');

// 전역 설치 모듈(pptxgenjs · playwright) 찾기
if (!process.env.NODE_PATH) {
  process.env.NODE_PATH = execSync('npm root -g', { encoding: 'utf8' }).trim();
  require('module').Module._initPaths();
}
const pptxgen = require('pptxgenjs');
const { convertDeck } = require('./html2pptx.cjs');

(async () => {
  const args = process.argv.slice(2);
  const flag = (k) => args.includes(k);
  const opt = (k) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : null; };
  const pos = args.filter((a, i) => !a.startsWith('--') && !(i > 0 && args[i - 1] === '--fontmap'));
  if (!pos[0]) { console.error('사용법: node build.cjs deck.html [out.pptx] [--fontmap f.json] [--no-embed] [--loose] [--keep]'); process.exit(1); }

  const html = path.resolve(pos[0]);
  const out = path.resolve(pos[1] || html.replace(/\.html?$/, '.pptx'));
  const src = fs.readFileSync(html, 'utf8');
  let fmPath = opt('--fontmap');
  if (!fmPath) { const m = src.match(/<meta\s+name="pptx-fontmap"\s+content="([^"]+)"/); if (m) fmPath = path.resolve(path.dirname(html), m[1]); }
  const fontmap = fmPath ? JSON.parse(fs.readFileSync(fmPath, 'utf8')) : null;

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';   // 1280×720px = 13.333×7.5in → 1px = 0.75pt 그대로
  const t = src.match(/<title>([^<]*)<\/title>/); if (t) pres.title = t[1];

  console.log(`① 장표 측정 · 글자/도형 변환 — ${path.basename(html)}`);
  const natives = await convertDeck(html, pres, { fontmap, strict: !flag('--loose') });
  const nNat = natives.reduce((a, s) => a + s.natives.length, 0);
  console.log(`   장표 ${natives.length}장 · 표/차트 ${nNat}개`);

  const base = out.replace(/\.pptx$/, '.base.pptx');
  const json = out.replace(/\.pptx$/, '.natives.json');
  await pres.writeFile({ fileName: base });
  fs.writeFileSync(json, JSON.stringify(natives));

  console.log('② 표 · 차트 네이티브 생성 · 한글 글꼴 · 글꼴 임베딩');
  const py = ['-I', path.join(__dirname, 'native.py'), base, json, out];
  if (fmPath) py.push('--fontmap', fmPath);
  if (flag('--no-embed')) py.push('--no-embed');
  execFileSync('python3', py, { stdio: 'inherit' });

  if (!flag('--keep')) { fs.unlinkSync(base); fs.unlinkSync(json); }
  console.log(`완료 → ${out} (${(fs.statSync(out).size / 1048576).toFixed(1)}MB)`);

  if (!flag('--no-verify')) {
    console.log('③ 검증');
    const v = ['-I', path.join(__dirname, 'verify.py'), html, out];
    if (flag('--ppt')) v.push('--compare', out.replace(/\.pptx$/, '_compare'));
    try { execFileSync('python3', v, { stdio: 'inherit' }); }
    catch (_) { console.error('검증 불합격 — 위 ✗ 항목을 확인하세요.'); process.exit(2); }
  }
})().catch((e) => { console.error(e.message); process.exit(1); });
