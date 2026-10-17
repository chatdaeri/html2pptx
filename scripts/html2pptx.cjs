/**
 * html2pptx (native) — HTML 장표를 편집 가능한 PowerPoint 요소로 옮긴다.
 * pptx-skill/scripts/html2pptx.cjs 에서 출발해 다음을 바꿨다.
 *   - 덱 모드: 한 HTML 안의 section.slide 여러 장을 장표 순서대로 변환 (convertDeck)
 *   - 배경·테두리가 있는 글자(p.badge · p.num · p.btn …) → 글자를 품은 도형
 *   - 구조형 목록(li 안에 블록 요소) → 컨테이너로 처리 / 단순 목록 → 네이티브 불릿(li::before 글자)
 *   - I · SPAN(글자 없음) · LI · HEADER 등도 배경·테두리가 있으면 도형
 *   - border-radius 50% 정사각형 → 원(ellipse), 균일 테두리는 선 중심선 보정(inset)
 *   - 글꼴: fontmap(CSS family + weight → PowerPoint 글꼴 이름 + bold) · letter-spacing → 자간
 *   - 표(.gtable / [data-pptx-table]) · 차트([data-chart]) → 자리·데이터만 추출(natives) → native.py 가 네이티브로 생성
 *   - 그림: <img> 는 object-fit/position 대로 PowerPoint '자르기'(원본 보존) · 둥근 모서리 · 원
 *           배경 이미지(background-image: url) 도 같은 방식, 장표 배경은 슬라이드 배경 그림
 *           SVG(인라인 · <img src=.svg>) → 대체 PNG(sharp) + SVG 원본(native.py) = PowerPoint 벡터 그림
 * 단위: 1280×720px = 13.333×7.5in (LAYOUT_WIDE), 1px = 0.75pt — 크기 보정 없음.
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const PT_PER_PX = 0.75;
const PX_PER_IN = 96;
const EMU_PER_IN = 914400;

// ───────────────────────── PptxGenJS 로 그리기 ─────────────────────────

function addElements(slideData, targetSlide, pres) {
  for (const el of slideData.elements) {
    if (el.type === 'native' || el.type === 'skip') continue;   // 표·차트 자리 — native.py 가 같은 z 순서로 끼워 넣음
    if (el.type === 'image' || el.type === 'svg') {
      const o = { x: el.position.x, y: el.position.y, w: el.position.w, h: el.position.h, objectName: el.name };
      if (el.type === 'svg') {
        o.data = 'data:image/png;base64,' + el.png.toString('base64');   // 대체 PNG — native.py 가 SVG 원본을 붙인다
      } else {
        if (el.src.startsWith('data:')) o.data = el.src; else o.path = fileOf(el.src);
        if (el.full) {   // 원본 그림 전체 크기 + 자르기 창 = CSS cover/contain/position 과 같은 화면 (원본 보존 → PowerPoint 에서 다시 자를 수 있음)
          Object.assign(o, { w: el.full.w, h: el.full.h, sizing: { type: 'crop', x: el.crop.x, y: el.crop.y, w: el.position.w, h: el.position.h } });
        }
      }
      if (el.ellipse) o.rounding = true;
      targetSlide.addImage(o);
    } else if (el.type === 'video') {
      targetSlide.addMedia({ type: 'video', path: fileOf(el.src), cover: el.posterData || (el.poster ? dataUri(fileOf(el.poster)) : undefined),
                             x: el.position.x, y: el.position.y, w: el.position.w, h: el.position.h, objectName: el.name });
    } else if (el.type === 'line') {
      targetSlide.addShape(pres.ShapeType.line, {
        x: el.x1, y: el.y1, w: el.x2 - el.x1, h: el.y2 - el.y1,
        line: { color: el.color, width: el.width, dashType: el.dash || 'solid' }
      });
    } else if (el.type === 'shape') {
      const s = el.shape;
      const o = {
        x: el.position.x, y: el.position.y, w: el.position.w, h: el.position.h,
        shape: s.ellipse ? pres.ShapeType.ellipse : (s.rectRadius > 0 ? pres.ShapeType.roundRect : pres.ShapeType.rect)
      };
      if (s.fill) { o.fill = { color: s.fill }; if (s.transparency != null) o.fill.transparency = s.transparency; }
      if (s.line) o.line = s.line;
      if (!s.ellipse && s.rectRadius > 0) o.rectRadius = s.rectRadius;
      if (s.shadow) o.shadow = s.shadow;
      if (el.style) Object.assign(o, el.style);   // 글자를 품은 도형
      else o.margin = 0;
      targetSlide.addText(lineRuns(el.text || ''), o);
    } else if (el.type === 'list') {
      targetSlide.addText(el.items, {
        x: el.position.x, y: el.position.y, w: el.position.w, h: el.position.h,
        fontSize: el.style.fontSize, fontFace: el.style.fontFace, color: el.style.color, bold: el.style.bold,
        align: el.style.align, valign: 'top', lineSpacing: el.style.lineSpacing, charSpacing: el.style.charSpacing,
        margin: 0
      });
    } else {
      const lineHeight = el.style.lineSpacing || el.style.fontSize * 1.2;
      const isSingleLine = el.position.h * 72 <= lineHeight * 1.5;   // position 은 inch, lineHeight 는 pt → ×72 로 단위 맞춤
      let x = el.position.x, w = el.position.w;
      if (isSingleLine) {   // 한 줄 글자는 글꼴 렌더 차이로 줄바꿈되지 않게 6% 여유 (원본 규칙)
        const inc = w * 0.06;
        if (el.style.align === 'center') { x -= inc / 2; w += inc; }
        else if (el.style.align === 'right') { x -= inc; w += inc; }
        else w += inc;
      }
      const o = {
        x, y: el.position.y, w, h: el.position.h,
        fontSize: el.style.fontSize, fontFace: el.style.fontFace, color: el.style.color,
        bold: el.style.bold, italic: el.style.italic, underline: el.style.underline,
        valign: el.style.valign || 'top', lineSpacing: el.style.lineSpacing, charSpacing: el.style.charSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore, paraSpaceAfter: el.style.paraSpaceAfter, inset: 0
      };
      // 한 줄 글자는 줄바꿈 금지 — 6% 여유로도 모자란 짧은 문구(버튼 · 캡션)가 다음 줄로 넘어가는 것을 막음
      if (el.style.wrap === false || isSingleLine) o.wrap = false;
      if (el.style.align) o.align = el.style.align;
      if (el.style.margin) o.margin = el.style.margin;
      if (el.style.rotate !== undefined) o.rotate = el.style.rotate;
      if (el.style.transparency != null) o.transparency = el.style.transparency;
      targetSlide.addText(lineRuns(el.text), o);
    }
  }
}

// ───────────────────────── 브라우저 안에서 측정 ─────────────────────────

async function extractSlideData(page, { rootSelector = null, rootIndex = null, fontmap = null } = {}) {
  return await page.evaluate(async ({ rootSelector, rootIndex, fontmap }) => {
    const errors = [];
    const root = rootSelector ? document.querySelectorAll(rootSelector)[rootIndex || 0] : document.body;
    const R = root.getBoundingClientRect();
    const inX = (px) => (px - R.left) / 96;
    const inY = (px) => (px - R.top) / 96;
    const inch = (px) => px / 96;
    const pt = (pxStr) => (parseFloat(pxStr) || 0) * 0.75;
    const rgbToHex = (s) => {
      if (!s || s === 'rgba(0, 0, 0, 0)' || s === 'transparent') return 'FFFFFF';
      const m = s.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      return m ? m.slice(1).map(n => parseInt(n).toString(16).padStart(2, '0')).join('').toUpperCase() : 'FFFFFF';
    };
    const alphaOf = (s) => { const m = s && s.match(/rgba\([^)]*,\s*([\d.]+)\)/); return m ? Math.round((1 - parseFloat(m[1])) * 100) : null; };
    const hasBgOf = (cs) => cs.backgroundColor && cs.backgroundColor !== 'rgba(0, 0, 0, 0)';
    const textTransform = (t, tt) => tt === 'uppercase' ? t.toUpperCase() : tt === 'lowercase' ? t.toLowerCase() : t;

    // 글꼴: CSS family + weight → PowerPoint 글꼴 이름 + bold (fontmap 의 가장 가까운 굵기)
    const mapFont = (familyStr, weightStr) => {
      const fam = (familyStr || '').split(',')[0].replace(/['"]/g, '').trim();
      const w = parseInt(weightStr) || (weightStr === 'bold' ? 700 : 400);
      const table = fontmap && fontmap.faces && fontmap.faces[fam];
      if (!table) return { fontFace: fam, bold: w >= 600 };
      const best = Object.keys(table).map(Number).reduce((a, b) => (Math.abs(b - w) < Math.abs(a - w) ? b : a));
      const [face, bold] = table[String(best)];
      return { fontFace: face, bold: !!bold };
    };
    const charSpacingOf = (cs) => (cs.letterSpacing && cs.letterSpacing !== 'normal' ? pt(cs.letterSpacing) || undefined : undefined);
    const runStyle = (cs) => {
      const f = mapFont(cs.fontFamily, cs.fontWeight);
      const o = { fontSize: pt(cs.fontSize), fontFace: f.fontFace, bold: f.bold, color: rgbToHex(cs.color) };
      const cspc = charSpacingOf(cs); if (cspc) o.charSpacing = cspc;
      if (cs.fontStyle === 'italic') o.italic = true;
      if (cs.textDecorationLine && cs.textDecorationLine.includes('underline')) o.underline = true;
      const tr = alphaOf(cs.color); if (tr !== null) o.transparency = tr;
      return o;
    };

    // 인라인 서식(<span> <b> <strong> <i> <em> <u> <mark> <a> <code> <br>) → 글자 조각(run)
    const INLINE = ['SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'SMALL', 'SUP', 'SUB', 'MARK', 'A', 'CODE', 'KBD', 'ABBR', 'S', 'DEL', 'INS', 'Q', 'CITE', 'TIME', 'LABEL'];
    // 형광펜(<mark> · 배경색 · inset box-shadow 밑줄 면) → PowerPoint 글자 강조색
    const highlightOf = (cs) => {
      if (hasBgOf(cs)) return rgbToHex(cs.backgroundColor);
      const m = cs.boxShadow && cs.boxShadow.includes('inset') && cs.boxShadow.match(/rgba?\([^)]*\)/);
      return m ? rgbToHex(m[0]) : null;
    };
    // 형광펜 띠 → 글자 뒤 사각형 (줄마다 한 개). 배경색이면 글자 상자 전체, inset box-shadow 면 그림자 두께만큼 아래(또는 위) 띠.
    const markBands = (textEl) => {
      const out = [];
      textEl.querySelectorAll('*').forEach((m) => {
        if (!INLINE.includes(m.tagName) || !m.textContent.trim()) return;
        const mcs = getComputedStyle(m);
        const color = highlightOf(mcs);
        if (!color) return;
        let off = null;   // inset 그림자 세로 오프셋(px): 음수 = 아래 띠
        if (!hasBgOf(mcs)) {
          const nums = mcs.boxShadow.replace(/rgba?\([^)]*\)/, '').match(/-?[\d.]+px/g) || [];
          off = parseFloat(nums[1] || '0');
          if (!off) return;
        }
        for (const r of m.getClientRects()) {
          if (r.width < 1 || r.height < 1) continue;
          const h = off === null ? r.height : Math.min(r.height, Math.abs(off));
          const y = off === null || off < 0 ? r.bottom - h : r.top;
          out.push({ type: 'shape', text: '', position: { x: inX(r.left), y: inY(y), w: inch(r.width), h: inch(h) },
                     shape: { fill: color, transparency: alphaOf(mcs.boxShadow) ?? null, line: null, ellipse: false, rectRadius: 0, shadow: null } });
        }
      });
      return out;
    };
    const parseInline = (element, base, runs = [], tt = (x) => x) => {
      let prevText = false;
      element.childNodes.forEach((node) => {
        const isText = node.nodeType === Node.TEXT_NODE || node.tagName === 'BR';
        if (isText) {
          const text = node.tagName === 'BR' ? '\n' : tt(node.textContent.replace(/\s+/g, ' '));
          const prev = runs[runs.length - 1];
          if (prevText && prev) prev.text += text; else runs.push({ text, options: { ...base } });
        } else if (node.nodeType === Node.ELEMENT_NODE && INLINE.includes(node.tagName) && node.textContent.trim()) {
          const cs = getComputedStyle(node);
          ['marginLeft', 'marginRight'].forEach(k => {
            if (parseFloat(cs[k]) > 0) errors.push(`인라인 <${node.tagName.toLowerCase()}> 에 ${k} 가 있습니다 — PowerPoint 는 인라인 여백을 표현하지 못합니다.`);
          });
          const t2 = cs.textTransform && cs.textTransform !== 'none' ? (s) => textTransform(s, cs.textTransform) : tt;
          const hl = highlightOf(cs);
          parseInline(node, { ...base, ...runStyle(cs), ...(hl ? { highlight: hl } : {}) }, runs, t2);
        } else if (node.nodeType === Node.ELEMENT_NODE && node.textContent.trim()) {   // 모르는 태그를 조용히 버리면 글자가 사라진다
          errors.push(`글자 안의 <${node.tagName.toLowerCase()}> 태그를 변환하지 못합니다 — "${node.textContent.trim().slice(0, 30)}" 가 빠집니다.`);
        }
        prevText = isText;
      });
      if (runs.length) {
        runs[0].text = runs[0].text.replace(/^\s+/, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      }
      return runs.filter(r => r.text.length > 0);
    };

    // 모서리: 정사각형 + 50% 이상 → 원, 그 외 둥근 사각형(반지름은 짧은 변의 절반까지)
    const radiusOf = (cs, rect) => {
      const v = parseFloat(cs.borderTopLeftRadius) || 0;
      if (!v) return { ellipse: false, rectRadius: 0 };
      const px = cs.borderTopLeftRadius.includes('%') ? (v / 100) * Math.min(rect.width, rect.height) : v;
      const half = Math.min(rect.width, rect.height) / 2;
      if (px >= half - 0.5 && Math.abs(rect.width - rect.height) < 1) return { ellipse: true, rectRadius: 0 };
      return { ellipse: false, rectRadius: inch(Math.min(px, half)) };
    };
    const shadowOf = (bs) => {
      if (!bs || bs === 'none' || /inset/.test(bs)) return null;
      const col = bs.match(/rgba?\([^)]+\)/);
      const p = bs.match(/(-?[\d.]+)px/g);
      if (!p || p.length < 2) return null;
      const [ox, oy, blur = 0] = p.map(parseFloat);
      if (!blur && !ox && !oy) return null;
      let ang = Math.atan2(oy, ox) * 180 / Math.PI; if (ang < 0) ang += 360;
      let op = 0.5; if (col) { const m = col[0].match(/,\s*([\d.]+)\)$/); if (m) op = parseFloat(m[1]); }
      return { type: 'outer', angle: Math.round(ang), blur: blur * 0.75, color: col ? rgbToHex(col[0]) : '000000', offset: Math.hypot(ox, oy) * 0.75, opacity: op };
    };

    // 배경·테두리 → 도형 (+ 부분 테두리는 선). 글자 스타일을 주면 글자를 품은 도형.
    const shapeFrom = (el, cs, rect, text = '', style = null, { fill = true, line = true } = {}) => {
      const bw = ['Top', 'Right', 'Bottom', 'Left'].map(s => (line ? parseFloat(cs[`border${s}Width`]) || 0 : 0));
      const hasBorder = bw.some(b => b > 0);
      const uniform = hasBorder && bw.every(b => b === bw[0]) &&
        ['Right', 'Bottom', 'Left'].every(s => cs[`border${s}Color`] === cs.borderTopColor);
      const out = [];
      const hasBg = fill && hasBgOf(cs);
      if (hasBg || uniform || style) {
        const rad = radiusOf(cs, rect);
        const inset = uniform ? bw[0] / 2 : 0;   // PowerPoint 윤곽선은 도형 경계 중앙에 그려짐 → 안쪽으로 보정
        out.push({
          type: 'shape', text, style,
          position: { x: inX(rect.left + inset), y: inY(rect.top + inset), w: inch(rect.width - 2 * inset), h: inch(rect.height - 2 * inset) },
          shape: {
            fill: hasBg ? rgbToHex(cs.backgroundColor) : null,
            transparency: hasBg ? alphaOf(cs.backgroundColor) : null,
            line: uniform ? { color: rgbToHex(cs.borderTopColor), width: bw[0] * 0.75 } : null,
            ellipse: rad.ellipse, rectRadius: rad.rectRadius, shadow: shadowOf(cs.boxShadow)
          }
        });
      }
      if (hasBorder && !uniform) {
        const x = rect.left, y = rect.top, w = rect.width, h = rect.height;
        const L = (b, s, x1, y1, x2, y2) => out.push({ type: 'line', x1: inX(x1), y1: inY(y1), x2: inX(x2), y2: inY(y2), width: b * 0.75,
          color: rgbToHex(cs[`border${s}Color`]), dash: cs[`border${s}Style`] === 'dashed' ? 'dash' : cs[`border${s}Style`] === 'dotted' ? 'sysDot' : 'solid' });
        if (bw[0]) L(bw[0], 'Top', x, y + bw[0] / 2, x + w, y + bw[0] / 2);
        if (bw[1]) L(bw[1], 'Right', x + w - bw[1] / 2, y, x + w - bw[1] / 2, y + h);
        if (bw[2]) L(bw[2], 'Bottom', x, y + h - bw[2] / 2, x + w, y + h - bw[2] / 2);
        if (bw[3]) L(bw[3], 'Left', x + bw[3] / 2, y, x + bw[3] / 2, y + h);
      }
      return out;
    };

    // 문단 스타일 (글자 상자·표 칸 공용)
    const paraStyle = (cs) => {
      const f = mapFont(cs.fontFamily, cs.fontWeight);
      const o = {
        fontSize: pt(cs.fontSize), fontFace: f.fontFace, bold: f.bold, color: rgbToHex(cs.color),
        align: cs.textAlign === 'start' || cs.textAlign === 'left' ? 'left' : cs.textAlign === 'end' ? 'right' : cs.textAlign,
        lineSpacing: cs.lineHeight && cs.lineHeight !== 'normal' ? pt(cs.lineHeight) : null,
        charSpacing: charSpacingOf(cs)
      };
      const tr = alphaOf(cs.color); if (tr !== null) o.transparency = tr;
      return o;
    };

    // 브라우저가 실제로 그린 줄 수 (글자 조각들의 서로 다른 세로 위치 수)
    const lineCount = (el) => {
      const r = document.createRange(); r.selectNodeContents(el);
      const tops = [...r.getClientRects()].filter(x => x.width > 0).map(x => Math.round(x.top + x.height / 2));
      return new Set(tops.map(t => Math.round(t / 4))).size;
    };

    // 단순 목록 판정: 모든 li 가 인라인 글자 또는 <p> 하나만 품을 때
    const BLOCKISH = (n) => n.nodeType === 1 && !INLINE.includes(n.tagName) && n.tagName !== 'BR';
    const simpleLi = (li) => {
      const kids = [...li.childNodes].filter(BLOCKISH);
      if (kids.length === 0) return true;
      return kids.length === 1 && kids[0].tagName === 'P' && ![...kids[0].childNodes].some(BLOCKISH);
    };
    const isSimpleList = (el) => {
      const lis = [...el.children].filter(c => c.tagName === 'LI');
      const plain = (li) => { const c = getComputedStyle(li); return !hasBgOf(c) && ['Top', 'Right', 'Bottom', 'Left'].every(k => !(parseFloat(c[`border${k}Width`]) > 0)); };
      return lis.length > 0 && lis.length === el.children.length && lis.every(li => simpleLi(li) && plain(li));
    };

    // ── 그림: 배경 이미지 원본 크기 미리 읽기 · CSS 크기/위치 → 표시 사각형 · SVG 직렬화
    const urlOf = (bg) => { const m = bg && bg !== 'none' && bg.match(/url\(["']?([^"')]+)["']?\)/); return m ? m[1] : null; };
    const natural = {};
    await Promise.all([...new Set([root, ...root.querySelectorAll('*')].map(n => urlOf(getComputedStyle(n).backgroundImage)).filter(Boolean))]
      .map(u => new Promise(res => { const im = new Image(); im.onload = () => { natural[u] = { w: im.naturalWidth, h: im.naturalHeight }; res(); }; im.onerror = () => { errors.push(`배경 이미지를 읽지 못했습니다: ${u}`); res(); }; im.src = u; })));
    // box 안에 CSS background-size / object-fit (+ position) 으로 놓인 그림 사각형(px)
    const placeImage = (box, nat, size, posStr) => {
      size = (size || 'auto').split(',')[0].trim();
      let w, h;
      if (size === 'cover' || size === 'contain' || size === 'scale-down') {
        let s = (size === 'cover' ? Math.max : Math.min)(box.w / nat.w, box.h / nat.h);
        if (size === 'scale-down') s = Math.min(1, s);
        w = nat.w * s; h = nat.h * s;
      } else if (size === 'fill') { w = box.w; h = box.h; }
      else {
        const [a, b = 'auto'] = size === 'none' ? ['auto'] : size.split(/\s+/);
        const len = (v, ref) => (v && v !== 'auto' ? (v.endsWith('%') ? parseFloat(v) / 100 * ref : parseFloat(v)) : null);
        w = len(a, box.w); h = len(b, box.h);
        if (w == null && h == null) { w = nat.w; h = nat.h; } else if (w == null) w = h * nat.w / nat.h; else if (h == null) h = w * nat.h / nat.w;
      }
      const [px = '50%', py = '50%'] = (posStr || '50% 50%').split(',')[0].trim().split(/\s+/);
      const off = (v, free) => (v.endsWith('%') ? parseFloat(v) / 100 * free : parseFloat(v));
      return { x: box.x + off(px, box.w - w), y: box.y + off(py, box.h - h), w, h };
    };
    let nameSeq = 0;
    const nextName = (p) => `${p}-${++nameSeq}`;
    // 그림 요소: 보이는 영역(= 상자 ∩ 그림) 을 틀로, 원본 전체 크기 + 자르기 위치를 같이 넘김
    const imageEl = (src, box, nat, size, posStr, cs, prefix) => {
      const img = placeImage(box, nat, size, posStr);
      const x0 = Math.max(box.x, img.x), y0 = Math.max(box.y, img.y);
      const x1 = Math.min(box.x + box.w, img.x + img.w), y1 = Math.min(box.y + box.h, img.y + img.h);
      const fr = { x: x0, y: y0, w: Math.max(x1 - x0, 1), h: Math.max(y1 - y0, 1) };
      const covers = Math.abs(fr.w - box.w) < 1 && Math.abs(fr.h - box.h) < 1;
      const rad = covers ? radiusOf(cs, { width: box.w, height: box.h }) : { ellipse: false, rectRadius: 0 };
      return { type: 'image', src, name: nextName(prefix),
               position: { x: inX(fr.x), y: inY(fr.y), w: inch(fr.w), h: inch(fr.h) },
               full: { w: inch(img.w), h: inch(img.h) }, crop: { x: inch(fr.x - img.x), y: inch(fr.y - img.y) },
               ellipse: rad.ellipse, radiusPx: rad.ellipse ? 0 : rad.rectRadius * 96, boxPx: { w: fr.w, h: fr.h } };
    };
    // 인라인 SVG → 독립 SVG 문자열 (CSS 클래스 · 변수 · currentColor 로 칠한 색을 속성으로 굳힘)
    const SVG_PROPS = ['fill', 'fill-opacity', 'fill-rule', 'stroke', 'stroke-width', 'stroke-opacity', 'stroke-dasharray', 'stroke-dashoffset',
      'stroke-linecap', 'stroke-linejoin', 'opacity', 'font-family', 'font-size', 'font-weight', 'font-style', 'text-anchor',
      'dominant-baseline', 'letter-spacing', 'visibility', 'stop-color', 'stop-opacity'];
    const serializeSvg = (svg, rect) => {
      const clone = svg.cloneNode(true);
      const src = [svg, ...svg.querySelectorAll('*')], dst = [clone, ...clone.querySelectorAll('*')];
      src.forEach((n, i) => { const c = getComputedStyle(n); dst[i].setAttribute('style', SVG_PROPS.map(k => `${k}:${c.getPropertyValue(k)}`).join(';')); });
      clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      clone.setAttribute('width', rect.width); clone.setAttribute('height', rect.height);
      if (!clone.getAttribute('viewBox')) clone.setAttribute('viewBox', `0 0 ${rect.width} ${rect.height}`);
      if (getComputedStyle(svg).transform !== 'none') errors.push('SVG 에 CSS transform 이 있습니다 — PowerPoint 그림에는 적용되지 않습니다.');
      return new XMLSerializer().serializeToString(clone);
    };

    // 차트 팔레트 = CSS 토큰
    const rootCs = getComputedStyle(document.documentElement);
    const palette = {};
    ['paper', 'bone', 'pebble', 'ink', 'charcoal', 'stone', 'ash', 'accent'].forEach(k => {
      const v = rootCs.getPropertyValue('--' + k).trim();
      if (v) palette[k] = v.replace('#', '').toUpperCase();
    });

    // ── 배경 (장표 면 색)
    const rootCsx = getComputedStyle(root);
    const background = { type: 'color', value: rgbToHex(hasBgOf(rootCsx) ? rootCsx.backgroundColor : getComputedStyle(document.body).backgroundColor) };
    const rootBg = urlOf(rootCsx.backgroundImage);
    if (rootBg && natural[rootBg]) {   // 장표 배경 그림 → PowerPoint 슬라이드 배경 (선택·이동되지 않음)
      const img = placeImage({ x: 0, y: 0, w: R.width, h: R.height }, natural[rootBg], rootCsx.backgroundSize, rootCsx.backgroundPosition);
      background.image = { src: rootBg, ...img, slideW: R.width, slideH: R.height };
    } else if (rootCsx.backgroundImage && rootCsx.backgroundImage !== 'none') {
      errors.push('장표 배경 그라데이션은 변환되지 않습니다 — 단색 또는 그림(url)만.');
    }

    const elements = [], natives = [], placeholders = [];
    const processed = new Set([root]);
    const TEXT = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'];
    const CONTAINERS = ['DIV', 'I', 'SPAN', 'LI', 'UL', 'OL', 'HEADER', 'FOOTER', 'SECTION', 'ARTICLE', 'ASIDE', 'NAV', 'MAIN', 'FIGURE'];

    root.querySelectorAll('*').forEach((el) => {
      if (processed.has(el)) return;
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) return;
      const cs = getComputedStyle(el);
      if (cs.visibility === 'hidden' || cs.display === 'none') return;

      // 모션그래픽 영상 → PowerPoint 동영상 (자리만 추출. 영상 파일은 data-video, 재생 전 그림은 data-poster)
      if (el.hasAttribute('data-video')) {
        const x0 = Math.max(rect.left, R.left), y0 = Math.max(rect.top, R.top);
        const x1 = Math.min(rect.right, R.right), y1 = Math.min(rect.bottom, R.bottom);   // 장표 밖으로 나간 부분은 잘라냄
        if (x1 - x0 < 1 || y1 - y0 < 1) { processed.add(el); return; }
        const abs = (a) => (a ? new URL(a, document.baseURI).href : null);
        elements.push({ type: 'video', name: nextName('video'), src: abs(el.getAttribute('data-video')), poster: abs(el.getAttribute('data-poster')),
                        autoplay: el.getAttribute('data-autoplay') !== 'false', loop: el.getAttribute('data-loop') !== 'false',
                        clip: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 },
                        position: { x: inX(x0), y: inY(y0), w: inch(x1 - x0), h: inch(y1 - y0) } });
        processed.add(el);
        el.querySelectorAll('*').forEach(c => processed.add(c));
        return;
      }

      // 표 · 차트 → 네이티브 자리
      const isTable = el.classList.contains('gtable') || el.hasAttribute('data-pptx-table');
      const isChart = el.hasAttribute('data-chart');
      if (isTable || isChart) {
        const box = { x: inX(rect.left), y: inY(rect.top), w: inch(rect.width), h: inch(rect.height) };
        let spec;
        if (isChart) {
          const chart = JSON.parse(el.getAttribute('data-chart'));
          // 그림 영역 = [data-plot-box] 들의 합집합 (컨테이너 기준 px) → PowerPoint 플롯 영역(수동 레이아웃)
          const pb = [...el.querySelectorAll('[data-plot-box]')].map(n => n.getBoundingClientRect());
          const plot = pb.length ? {
            x: Math.min(...pb.map(r => r.left)) - rect.left, y: Math.min(...pb.map(r => r.top)) - rect.top,
            r: Math.max(...pb.map(r => r.right)) - rect.left, b: Math.max(...pb.map(r => r.bottom)) - rect.top
          } : { x: 0, y: 0, r: rect.width, b: rect.height };
          plot.w = plot.r - plot.x; plot.h = plot.b - plot.y;
          // 막대 굵기 · 간격 실측 (PowerPoint gapWidth 계산용)
          const geom = {};
          const firstBar = el.querySelector('.bar > i, .bar > .stack, [data-bar]');
          if (firstBar) geom.barW = firstBar.getBoundingClientRect().width;
          const hb = [...el.querySelectorAll('.track > i')].map(n => n.getBoundingClientRect());
          if (hb.length) {
            geom.barH = hb[0].height;
            geom.pitch = hb.length > 1 ? hb[1].top - hb[0].top : plot.h;
            const v0 = chart.values && chart.values[0];
            geom.fullW = v0 ? hb[0].width * chart.max / v0 : plot.w;   // 값 = max 일 때 막대 길이
            geom.x0 = hb[0].left - rect.left;
            geom.c0 = hb[0].top + hb[0].height / 2 - rect.top;
          }
          spec = { kind: 'chart', box, chart, palette, plot, geom, size: { w: rect.width, h: rect.height } };
        } else {
          const top = parseFloat(cs.borderTopWidth) || 0;
          spec = {
            kind: 'table', box, palette,
            top: top ? [top * 0.75, rgbToHex(cs.borderTopColor)] : null,
            cells: [...el.children].filter(c => c.getBoundingClientRect().width > 0).map(c => {
              const r = c.getBoundingClientRect(), s = getComputedStyle(c);
              const side = (k) => { const w = parseFloat(s[`border${k}Width`]) || 0; return w ? [w * 0.75, rgbToHex(s[`border${k}Color`])] : null; };
              const ps = [...c.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li')].filter(p => !p.querySelector('p,li'));
              return {
                x: r.left - rect.left, y: r.top - rect.top, w: r.width, h: r.height,
                // 음수 margin-top(합계 행 윗선) 은 위 칸 경계로 본다
                fill: hasBgOf(s) ? rgbToHex(s.backgroundColor) : null,
                border: { t: side('Top'), b: side('Bottom'), l: side('Left'), r: side('Right') },
                pad: [pt(s.paddingLeft), pt(s.paddingRight), pt(s.paddingTop), pt(s.paddingBottom)],
                valign: s.justifyContent === 'flex-end' ? 'bottom' : s.justifyContent === 'flex-start' || s.justifyContent === 'normal' ? 'middle' : 'middle',
                paras: ps.map(p => { const pcs = getComputedStyle(p); const st = paraStyle(pcs); return { style: st, runs: parseInline(p, runStyle(pcs), [], (x) => textTransform(x, pcs.textTransform)) }; })
              };
            })
          };
        }
        natives.push(spec);
        elements.push({ type: 'native', id: natives.length - 1 });
        processed.add(el);
        el.querySelectorAll('*').forEach(c => { if (!c.closest('[data-pptx-keep]')) processed.add(c); });
        return;
      }

      // 그림 (<img>) — SVG 파일이면 벡터 그림, 아니면 object-fit 대로 자르기
      if (el.tagName === 'IMG') {
        const src = el.currentSrc || el.src;
        const pos = { x: inX(rect.left), y: inY(rect.top), w: inch(rect.width), h: inch(rect.height) };
        if (/\.svg([?#]|$)/i.test(src) || src.startsWith('data:image/svg')) elements.push({ type: 'svg', src, position: pos, name: nextName('svg') });
        else if (!el.naturalWidth) errors.push(`그림을 읽지 못했습니다: ${src}`);
        else elements.push(imageEl(src, { x: rect.left, y: rect.top, w: rect.width, h: rect.height }, { w: el.naturalWidth, h: el.naturalHeight },
                                   cs.objectFit === 'none' ? 'auto' : cs.objectFit, cs.objectPosition, cs, 'img'));
        processed.add(el);
        return;
      }
      // 인라인 SVG (아이콘 · 로고 · 일러스트) → 벡터 그림. 차트 안 SVG 는 위에서 차트로 처리됨
      if (el.tagName.toLowerCase() === 'svg') {
        elements.push({ type: 'svg', markup: serializeSvg(el, rect), name: nextName('svg'),
                        position: { x: inX(rect.left), y: inY(rect.top), w: inch(rect.width), h: inch(rect.height) } });
        processed.add(el);
        el.querySelectorAll('*').forEach(c => processed.add(c));
        return;
      }
      if (el.closest('svg')) { processed.add(el); return; }

      // 단순 목록 → 네이티브 불릿 글자 상자
      if ((el.tagName === 'UL' || el.tagName === 'OL') && isSimpleList(el)) {
        const lis = [...el.children];
        const items = [];
        lis.forEach((li, i) => {
          const src = li.querySelector('p') || li;
          const lcs = getComputedStyle(li), scs = getComputedStyle(src);
          const runs = parseInline(src, runStyle(scs), [], (x) => textTransform(x, scs.textTransform));
          const before = getComputedStyle(li, '::before').content;
          const ch = before && before !== 'none' && before !== 'normal' ? before.replace(/^["']|["']$/g, '') : '';
          if (runs.length) {
            const first = runs[0].options;
            first.bullet = ch ? { characterCode: ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'), indent: pt(lcs.paddingLeft) } : false;
            if (i > 0) first.paraSpaceBefore = pt(lcs.marginTop);
            if (i < lis.length - 1) runs[runs.length - 1].options.breakLine = true;
          }
          items.push(...runs);
          processed.add(li); li.querySelectorAll('*').forEach(c => processed.add(c));
        });
        const st = paraStyle(getComputedStyle(lis[0].querySelector('p') || lis[0]));
        elements.push({ type: 'list', items, position: { x: inX(rect.left), y: inY(rect.top), w: inch(rect.width), h: inch(rect.height) }, style: st });
        processed.add(el);
        return;
      }

      // 글자 (p · h1~h6)
      if (TEXT.includes(el.tagName)) {
        const text = el.textContent.trim();
        processed.add(el);
        el.querySelectorAll('*').forEach(c => processed.add(c));
        if (!text) {   // 빈 문단이라도 배경·테두리가 있으면 도형
          elements.push(...shapeFrom(el, cs, rect));
          return;
        }
        const st = paraStyle(cs);
        const runs = parseInline(el, runStyle(cs), [], (x) => textTransform(x, cs.textTransform));
        const boxed = hasBgOf(cs) || ['Top', 'Right', 'Bottom', 'Left'].some(s => parseFloat(cs[`border${s}Width`]) > 0);
        if (boxed) {   // 배지 · 번호 원 · 버튼 · 주석 pill → 글자를 품은 도형
          const centered = cs.justifyContent === 'center' || cs.textAlign === 'center' || (cs.display.includes('flex') && Math.abs(parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight)) < 1);
          const style = {
            fontSize: st.fontSize, fontFace: st.fontFace, bold: st.bold, color: st.color, charSpacing: st.charSpacing,
            align: centered ? 'center' : st.align, valign: cs.alignItems === 'center' || cs.display.includes('flex') ? 'middle' : 'top',
            margin: [pt(cs.paddingLeft), pt(cs.paddingRight), pt(cs.paddingBottom), pt(cs.paddingTop)]
          };
          if (lineCount(el) <= 1) { style.wrap = false; style.fit = 'none'; }   // 한 줄 배지 · 버튼은 줄바꿈 금지
          elements.push(...shapeFrom(el, cs, rect, runs.length === 1 && Object.keys(runs[0].options).length <= 6 ? runs[0].text : runs, style));
          return;
        }
        const bands = markBands(el);
        if (bands.length) {   // 띠 도형이 형광펜을 대신함 — 글자 강조색은 지움 (목록 · 표 칸 · 배지 안은 강조색 그대로)
          elements.push(...bands);
          runs.forEach(r => { delete r.options.highlight; });
        }
        elements.push({
          type: el.tagName.toLowerCase(), text: runs,
          position: { x: inX(rect.left), y: inY(rect.top), w: inch(rect.width), h: inch(rect.height) },
          style: { ...st, paraSpaceBefore: 0, paraSpaceAfter: 0,
                   valign: cs.display.includes('flex') && cs.alignItems === 'center' ? 'middle' : 'top',
                   // 실제 한 줄이거나, 화면의 줄이 모두 <br> 로 나눈 줄이면 줄바꿈 금지 → PowerPoint 가 줄을 새로 나누지 않음
                   wrap: cs.whiteSpace === 'nowrap' || lineCount(el) <= 1 || lineCount(el) === el.querySelectorAll('br').length + 1 ? false : undefined,
                   margin: [pt(cs.paddingLeft), pt(cs.paddingRight), pt(cs.paddingBottom), pt(cs.paddingTop)] }
        });
        return;
      }

      // 컨테이너 (배경·테두리 → 도형)
      if (CONTAINERS.includes(el.tagName)) {
        for (const n of el.childNodes) {
          if (n.nodeType === Node.TEXT_NODE && n.textContent.trim() && el.tagName !== 'LI') {
            errors.push(`<${el.tagName.toLowerCase()}> 안에 감싸지 않은 글자 "${n.textContent.trim().slice(0, 40)}" — <p> 로 감싸야 PowerPoint 에 나타납니다.`);
          }
        }
        const bgUrl = urlOf(cs.backgroundImage);
        if (cs.backgroundImage && cs.backgroundImage !== 'none' && !bgUrl) errors.push(`<${el.tagName.toLowerCase()}> 의 그라데이션 배경은 변환되지 않습니다 — 단색 또는 그림(url)만.`);
        if (bgUrl && natural[bgUrl]) {   // CSS 와 같은 겹침: 면 색 → 배경 그림 → 테두리
          elements.push(...shapeFrom(el, cs, rect, '', null, { line: false }));
          elements.push(imageEl(bgUrl, { x: rect.left, y: rect.top, w: rect.width, h: rect.height }, natural[bgUrl], cs.backgroundSize, cs.backgroundPosition, cs, 'bg'));
          elements.push(...shapeFrom(el, cs, rect, '', null, { fill: false }));
        } else {
          elements.push(...shapeFrom(el, cs, rect));
        }
        processed.add(el);
      }
    });

    return { background, elements, natives, placeholders, errors };
  }, { rootSelector, rootIndex, fontmap });
}

// ───────────────────────── 진입점 ─────────────────────────

async function launch(tmpDir) {
  const o = { env: { TMPDIR: tmpDir }, args: ['--allow-file-access-from-files'] };
  if (process.platform === 'darwin') o.channel = 'chrome';
  return chromium.launch(o);
}

// <br> 줄바꿈('\n')을 품은 글자 조각 → 줄마다 조각을 나누고 줄 끝에만 breakLine.
// PptxGenJS 는 '\n' 이 든 조각을 스스로 나눌 때 마지막 줄 뒤에도 문단을 끊어, 뒤따르는 조각(<mark> 등)이 다음 줄로 밀린다.
function lineRuns(text) {
  if (!Array.isArray(text)) return text;
  const out = [];
  for (const r of text) {
    const lines = r.text.split('\n');
    lines.forEach((t, k) => {
      const last = k === lines.length - 1;
      if (t || !last) out.push({ text: t, options: { ...r.options, ...(last ? {} : { breakLine: true }) } });
    });
  }
  return out;
}

// file:// URL · data: URI → 로컬 경로 / 내용
function fileOf(src) {
  if (!src.startsWith('file://')) return src;
  try { return decodeURIComponent(src.replace('file://', '')); } catch (_) { return src.replace('file://', ''); }
}
// 로컬 그림 파일 → data URI (영상 표지용 — PptxGenJS 가 base64 만 받음)
function dataUri(file) {
  const ext = path.extname(file).slice(1).toLowerCase().replace('jpg', 'jpeg');
  return `data:image/${ext || 'png'};base64,` + fs.readFileSync(file).toString('base64');
}
function svgTextOf(el) {
  if (el.markup) return el.markup;
  if (el.src.startsWith('data:')) {
    const [head, body] = el.src.split(/,(.*)/s);
    return head.includes(';base64') ? Buffer.from(body, 'base64').toString('utf8') : decodeURIComponent(body);
  }
  return fs.readFileSync(fileOf(el.src), 'utf8');
}

// SVG → 대체 PNG (2배 해상도). PowerPoint 2016+ 는 SVG 원본을, 그 이전 · 다른 앱은 이 PNG 를 보여 준다
async function svgFallback(el) {
  el.svgText = svgTextOf(el);
  const w = Math.max(1, Math.round(el.position.w * PX_PER_IN * 2)), h = Math.max(1, Math.round(el.position.h * PX_PER_IN * 2));
  el.png = await sharp(Buffer.from(el.svgText), { density: 288 }).resize(w, h, { fit: 'fill' }).png().toBuffer();
}

// 장표 배경 그림 → CSS 크기/위치대로 잘라 슬라이드 크기 그림으로 (면 색 위에 합성, 2배 해상도)
async function slideBackground(bg) {
  const { src, x, y, w, h, slideW: W, slideH: H } = bg.image;
  const S = 2;
  const input = src.startsWith('data:') ? Buffer.from(src.split(',')[1], 'base64') : fileOf(src);
  const meta = await sharp(input).metadata();
  const ix0 = Math.max(0, x), iy0 = Math.max(0, y), ix1 = Math.min(W, x + w), iy1 = Math.min(H, y + h);
  const sx = meta.width / w, sy = meta.height / h;
  const left = Math.min(meta.width - 1, Math.round((ix0 - x) * sx)), top = Math.min(meta.height - 1, Math.round((iy0 - y) * sy));
  const region = await sharp(input)
    .extract({ left, top, width: Math.max(1, Math.min(meta.width - left, Math.round((ix1 - ix0) * sx))), height: Math.max(1, Math.min(meta.height - top, Math.round((iy1 - iy0) * sy))) })
    .resize(Math.max(1, Math.round((ix1 - ix0) * S)), Math.max(1, Math.round((iy1 - iy0) * S)), { fit: 'fill' }).png().toBuffer();
  const buf = await sharp({ create: { width: Math.round(W * S), height: Math.round(H * S), channels: 3, background: '#' + bg.value } })
    .composite([{ input: region, left: Math.round(ix0 * S), top: Math.round(iy0 * S) }]).jpeg({ quality: 90 }).toBuffer();
  return 'data:image/jpeg;base64,' + buf.toString('base64');
}

// 덱 모드: deck.html 의 section.slide 를 순서대로 변환. 반환 = 장표별 natives (z 순서 포함) · SVG 원본 · 둥근 그림
async function convertDeck(htmlFile, pres, { fontmap = null, slideSelector = '.slide', tmpDir = process.env.TMPDIR || '/tmp', strict = true } = {}) {
  const browser = await launch(tmpDir);
  const filePath = path.isAbsolute(htmlFile) ? htmlFile : path.join(process.cwd(), htmlFile);
  const result = [];
  try {
    // 1340px 이하면 frame.css 가 장표를 줄여 보여줌(zoom) → 화면 폭을 넉넉히
    const page = await (await browser.newContext({ viewport: { width: 1400, height: 800 } })).newPage();
    page.on('console', (m) => { if (m.type() === 'error') console.log('브라우저 오류:', m.text()); });
    await page.goto(`file://${filePath}`, { waitUntil: 'load' });
    await page.evaluate(() => document.fonts.ready);
    const n = await page.locator(slideSelector).count();
    if (!n) throw new Error(`${htmlFile}: 장표(${slideSelector})가 없습니다.`);
    const allErrors = [];
    for (let i = 0; i < n; i++) {
      const box = await page.locator(slideSelector).nth(i).boundingBox();
      const lw = pres.presLayout.width / EMU_PER_IN, lh = pres.presLayout.height / EMU_PER_IN;
      if (Math.abs(box.width / PX_PER_IN - lw) > 0.05 || Math.abs(box.height / PX_PER_IN - lh) > 0.05) {
        allErrors.push(`장표 ${i + 1}: 크기 ${box.width}×${box.height}px 가 PowerPoint 레이아웃(${lw}×${lh}in)과 다릅니다.`);
      }
      const data = await extractSlideData(page, { rootSelector: slideSelector, rootIndex: i, fontmap });
      data.errors.forEach(e => allErrors.push(`장표 ${i + 1}: ${e}`));
      for (const el of data.elements) if (el.type === 'svg') {
        try { await svgFallback(el); } catch (e) { allErrors.push(`장표 ${i + 1}: SVG 를 그리지 못했습니다 (${e.message})`); el.type = 'skip'; }
      }
      for (const el of data.elements) if (el.type === 'video') {
        if (!fs.existsSync(fileOf(el.src))) { allErrors.push(`장표 ${i + 1}: 영상 파일이 없습니다: ${el.src}`); el.type = 'skip'; continue; }
        if (!el.poster) {   // data-poster 가 없으면 지금 화면(영상 자리)을 그대로 캡처
          const png = await page.screenshot({ clip: el.clip, fullPage: true });
          el.posterData = 'data:image/png;base64,' + png.toString('base64');
        }
      }
      const slide = pres.addSlide();
      if (data.background.image) slide.background = { data: await slideBackground(data.background) };
      else if (data.background.value) slide.background = { color: data.background.value };
      addElements(data, slide, pres);
      // native.py 용: 각 표·차트 앞에 그려진 도형 수 = 끼워 넣을 z 위치
      let z = 0;
      const natives = [];
      for (const el of data.elements) {
        if (el.type === 'native') natives.push({ ...data.natives[el.id], z });
        else if (el.type !== 'skip') z++;
      }
      const svgs = data.elements.filter(e => e.type === 'svg').map(e => ({ name: e.name, svg: e.svgText }));
      const pics = data.elements.filter(e => e.type === 'image' && e.radiusPx > 0).map(e => ({ name: e.name, r: e.radiusPx, w: e.boxPx.w, h: e.boxPx.h }));
      const videos = data.elements.filter(e => e.type === 'video').map(e => ({ name: e.name, autoplay: e.autoplay, loop: e.loop }));
      result.push({ slide: i, natives, svgs, pics, videos });
    }
    if (allErrors.length) {
      const msg = `변환 규칙 위반 ${allErrors.length}건:\n` + allErrors.map((e, k) => `  ${k + 1}. ${e}`).join('\n');
      if (strict) throw new Error(msg); else console.warn(msg);
    }
  } finally {
    await browser.close();
  }
  return result;
}

module.exports = { convertDeck, extractSlideData, addElements };
