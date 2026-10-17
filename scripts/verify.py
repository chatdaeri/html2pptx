"""PPTX 변환 결과 검증 — 합격/불합격 판정 (build.cjs 가 변환 직후 자동 실행).

    python3 verify.py deck.html deck.pptx [--render] [--compare out_dir] [--max-diff 5]

기본 검사 (PowerPoint 없이, 몇 초):
  1. 파일      zip · XML 이 깨지지 않았고 python-pptx 로 열림
  2. 무결성    PowerPoint '복구' 창을 띄우는 패턴 — 차트/표/프레젠테이션 XML 의 단일 요소 중복 · 장표 안 도형 번호(id) 중복
  3. 장표 수   HTML section.slide 수 = PPTX 장표 수
  4. 네이티브  .gtable 수 = PPTX 표 수 · data-chart 수 = PPTX 차트 수 · 차트마다 엑셀 내장
  5. 그림      HTML SVG 수 = 벡터 그림(SVG 원본 포함) 수 · 대체 PNG 정상 · 사진 수 ≥ HTML img 수 (배경 그림 포함)
  6. 글자      HTML 장표의 글자 조각(SVG·차트·영상 안 제외)이 모두 같은 PPTX 장표에 있음 — 모르는 태그로 글자가 빠지는 사고 방지
  7. 영상      HTML data-video 수 = PPTX 동영상 수 · 자동 재생 수 · 영상 파일이 비어 있지 않음 (영상이 없는 덱은 생략)
  8. 글꼴      글자에 쓰인 글꼴이 모두 파일에 들어 있음(EOT) — 테마 기본 글꼴 제외
선택:
  --render     실제 PowerPoint 로 열어 PDF 저장 (ppt_render.sh 안전 규칙) → 페이지 수 = 장표 수
  --compare D  HTML ↔ PowerPoint 비교 이미지(compare.py) → 차이 %가 --max-diff 넘는 장표 표시
종료 코드: 0 합격 · 1 불합격
"""
import os
import re
import subprocess
import sys
from html import unescape as html_unescape
from html.parser import HTMLParser
import zipfile
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
SINGLE = {   # 부모 태그 → 한 번만 나와야 하는 자식 (중복이면 PowerPoint 가 '복구' 요구)
    'c:chartSpace': ['c:date1904', 'c:lang', 'c:roundedCorners', 'c:chart', 'c:spPr', 'c:txPr', 'c:externalData'],
    'c:chart': ['c:title', 'c:autoTitleDeleted', 'c:plotArea', 'c:legend', 'c:plotVisOnly', 'c:dispBlanksAs'],
    'c:plotArea': ['c:layout', 'c:spPr'],
    'c:ser': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dLbls', 'c:cat', 'c:val', 'c:smooth'],
    'a:tcPr': ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'],
    'a:rPr': ['a:latin', 'a:ea', 'a:cs'],
    'p:presentation': ['p:sldIdLst', 'p:sldSz', 'p:notesSz', 'p:embeddedFontLst', 'p:defaultTextStyle'],
}
THEME = {'Arial', 'Calibri', 'Calibri Light', '맑은 고딕', 'Times New Roman'}


def check_xml(z):
    from lxml import etree
    bad = []
    for n in z.namelist():
        if not n.endswith('.xml'):
            continue
        try:
            root = etree.fromstring(z.read(n))
        except etree.XMLSyntaxError as e:
            bad.append(f'{n}: XML 오류 {e}')
            continue
        nsmap = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
        for parent, kids in SINGLE.items():
            pfx, tag = parent.split(':')
            for el in root.iter('{%s}%s' % (nsmap[pfx], tag)):
                cnt = Counter(etree.QName(c).localname for c in el)
                for k in kids:
                    if cnt[k.split(':')[1]] > 1:
                        bad.append(f'{n}: <{parent}> 안에 <{k}> {cnt[k.split(":")[1]]}개')
    return bad



class SlideText(HTMLParser):
    """HTML section.slide 마다 화면 글자 조각(텍스트 노드)을 모은다. SVG 안 글자 · 차트 · script/style 은 그림·네이티브로 옮겨지므로 제외."""
    VOID = {'br', 'img', 'input', 'meta', 'link', 'hr', 'source', 'wbr', 'area', 'col', 'embed', 'track', 'path', 'circle', 'rect', 'line', 'use', 'stop', 'ellipse', 'polygon', 'polyline'}

    def __init__(self):
        super().__init__()
        self.slides, self.stack, self.skip = [], [], 0

    def handle_starttag(self, tag, attrs):
        if tag in self.VOID:
            return
        a = dict(attrs)
        is_slide = tag == 'section' and re.search(r'\bslide\b', a.get('class') or '')
        if is_slide:
            self.slides.append([])
        skip = tag in ('svg', 'script', 'style', 'template') or 'data-chart' in a or 'data-video' in a
        self.stack.append((tag, skip))
        self.skip += skip

    def handle_startendtag(self, tag, attrs):
        pass

    def handle_endtag(self, tag):
        if tag in self.VOID:
            return
        while self.stack:
            t, sk = self.stack.pop()
            self.skip -= sk
            if t == tag:
                break

    def handle_data(self, d):
        t = re.sub(r'\s+', '', d)
        if t and self.slides and not self.skip:
            self.slides[-1].append(t)


def flat(x):
    return re.sub(r'\s+', '', x).lower()

def main():
    a = sys.argv[1:]
    html, pptx = a[0], a[1]
    src = open(html, encoding='utf8').read()
    res = []   # (항목, 통과?, 내용)

    try:
        z = zipfile.ZipFile(pptx)
        assert z.testzip() is None
        from pptx import Presentation
        prs = Presentation(pptx)
        res.append(('파일', True, f'{os.path.getsize(pptx) / 1048576:.1f}MB · python-pptx 로 열림'))
    except Exception as e:
        res.append(('파일', False, f'열리지 않음: {e}'))
        return report(res)

    bad = check_xml(z)
    for n in sorted(x for x in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', x)):
        ids = re.findall(r'<p:cNvPr id="(\d+)"', z.read(n).decode('utf8'))
        dup = sorted({i for i in ids if ids.count(i) > 1}, key=int)
        if dup:
            bad.append(f'{n.split("/")[-1]} 도형 번호 중복 {",".join(dup)}')
    res.append(('무결성', not bad, '단일 요소 중복 없음' if not bad else ' / '.join(bad[:5]) + (f' 외 {len(bad) - 5}건' if len(bad) > 5 else '')))

    n_html = len(re.findall(r'<section\s+class="[^"]*\bslide\b', src))
    res.append(('장표 수', n_html == len(prs.slides), f'HTML {n_html} · PPTX {len(prs.slides)}'))

    tables = charts = wb = vec = raster = 0
    bad_png = []
    for s in prs.slides:
        for sh in s.shapes:
            tables += sh.has_table
            if sh.has_chart:
                charts += 1
                wb += sh.chart.part.chart_workbook.xlsx_part is not None
            if sh.shape_type == 13:
                if b'svgBlip' in etree_bytes(sh._element):
                    vec += 1
                    if len(sh.image.blob) < 200:
                        bad_png.append(sh.name)
                else:
                    raster += 1
    h_tab = len(re.findall(r'class="[^"]*\bgtable\b', src)) + src.count('data-pptx-table')
    h_ch = src.count('data-chart=')
    imgs = re.findall(r'<img\b[^>]*>', src)
    h_svgimg = sum(bool(re.search(r'src="[^"]*(\.svg|data:image/svg)', t)) for t in imgs)
    h_vec = len(re.findall(r'<svg\b', src)) + h_svgimg
    h_raster = len(imgs) - h_svgimg
    res.append(('네이티브', h_tab == tables and h_ch == charts and wb == charts, f'표 {tables}/{h_tab} · 차트 {charts}/{h_ch} (엑셀 내장 {wb})'))
    res.append(('그림', vec == h_vec and raster >= h_raster and not bad_png,
                f'SVG 벡터 {vec}/{h_vec} · 사진 {raster} (HTML img {h_raster} + 배경 그림)' + (f' · 대체 PNG 이상: {", ".join(bad_png)}' if bad_png else '')))

    # 글자: HTML 장표의 글자 조각이 모두 같은 PPTX 장표에 있음 (태그를 몰라 글자가 빠지는 사고 방지)
    parser = SlideText()
    parser.feed(src)
    lost = []
    for i, (texts, sl) in enumerate(zip(parser.slides, prs.slides)):
        n = f'ppt/slides/slide{i + 1}.xml'
        have = flat(''.join(re.findall(r'<a:t>([^<]*)</a:t>', html_unescape(z.read(n).decode('utf8')))))
        lost += [f'{i + 1}장 "{t[:20]}"' for t in texts if flat(html_unescape(t)) not in have]
    n_txt = sum(len(t) for t in parser.slides)
    res.append(('글자', not lost, f'글자 조각 {n_txt - len(lost)}/{n_txt}' + (f' · 빠짐: {", ".join(lost[:5])}' + (f' 외 {len(lost) - 5}건' if len(lost) > 5 else '') if lost else '')))

    # 영상: HTML data-video 수 = PPTX 동영상 수 · 영상 파일이 파일 안에 들어 있고 표지 그림이 있음
    h_vid = len(re.findall(r'\bdata-video="', src))
    vids = bad_vid = auto = 0
    for n in z.namelist():
        if re.match(r'ppt/slides/slide\d+\.xml$', n):
            x = z.read(n).decode('utf8')
            vids += x.count('<a:videoFile')
            auto += x.count('cmd="playFrom(0.0)"')
    for n in z.namelist():
        if n.startswith('ppt/media/') and n.lower().endswith(('.mp4', '.mov', '.m4v', '.wmv')) and z.getinfo(n).file_size < 1000:
            bad_vid += 1
    if h_vid or vids:
        h_auto = h_vid - len(re.findall(r'\bdata-autoplay="false"', src))
        res.append(('영상', vids == h_vid and auto == h_auto and not bad_vid,
                    f'동영상 {vids}/{h_vid} · 자동 재생 {auto}/{h_auto}' + (f' · 빈 영상 파일 {bad_vid}개' if bad_vid else '')))

    used, embedded = set(), set()
    for n in z.namelist():
        if re.match(r'ppt/(slides|charts)/[^/]+\.xml$', n):
            x = z.read(n).decode('utf8')
            for m in re.finditer(r'<a:r>.*?</a:r>|<a:defRPr.*?</a:defRPr>', x, re.S):   # 실제 글자 · 차트 기본 글자에 쓰인 글꼴만
                used |= set(re.findall(r'<a:(?:latin|ea) typeface="([^"+][^"]*)"', m.group(0)))
    pres = z.read('ppt/presentation.xml').decode('utf8')
    embedded = set(re.findall(r'<p:font typeface="([^"]+)"', pres))
    eot_ok = all(z.read(n)[34:36] == b'\x4c\x50' for n in z.namelist() if n.endswith('.fntdata'))
    missing = sorted(used - embedded - THEME)
    res.append(('글꼴', not missing and eot_ok,
                f'사용 {len(used)}종 · 임베딩 {len(embedded)}종' + (f' · 빠짐: {", ".join(missing)}' if missing else '') + ('' if eot_ok else ' · EOT 아닌 글꼴 파일 있음')))

    if '--render' in a or '--compare' in a:
        out = a[a.index('--compare') + 1] if '--compare' in a else os.path.splitext(pptx)[0] + '_verify'
        os.makedirs(out, exist_ok=True)
        pdf = os.path.join(out, 'pptx.pdf')
        r = subprocess.run(['bash', os.path.join(HERE, 'ppt_render.sh'), pptx, pdf], capture_output=True, text=True)
        if r.returncode == 0:
            pages = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout).group(1))
            res.append(('PowerPoint', pages == len(prs.slides), f'복구 창 없이 열림 · PDF {pages}쪽'))
            if '--compare' in a:
                lim = float(a[a.index('--max-diff') + 1]) if '--max-diff' in a else 5.0
                c = subprocess.run([sys.executable, '-I', os.path.join(HERE, 'compare.py'), html, pptx, out, '--pptx-pdf', pdf],
                                   capture_output=True, text=True)
                diffs = [(int(m.group(1)), float(m.group(2))) for m in re.finditer(r'#(\d+) diff\s+([\d.]+)%', c.stdout)]
                if not diffs:
                    res.append(('비교', False, '비교 이미지 생성 실패 — ' + (c.stderr.strip().splitlines() or ['알 수 없음'])[-1]))
                    return report(res)
                over = [f'#{i:02d} {d}%' for i, d in diffs if d > lim]
                res.append(('비교', not over, f'{len(diffs)}장 · 최대 {max(d for _, d in diffs):.1f}% · 기준 {lim}%'
                            + (f' · 확인 필요: {", ".join(over)}' if over else '') + f' → {out}'))
        else:
            msg = (r.stderr or r.stdout).strip().splitlines()[-1] if (r.stderr or r.stdout).strip() else '알 수 없음'
            res.append(('PowerPoint', r.returncode == 3, ('건너뜀 — ' if r.returncode == 3 else '실패 — ') + msg))
    return report(res)


def etree_bytes(el):
    from lxml import etree
    return etree.tostring(el)


def report(res):
    w = max(len(k) for k, _, _ in res)
    for k, ok, msg in res:
        print(f'   {"✓" if ok else "✗"} {k.ljust(w)}  {msg}')
    failed = [k for k, ok, _ in res if not ok]
    print('   검증 ' + ('합격' if not failed else f'불합격 ({", ".join(failed)})'))
    sys.exit(1 if failed else 0)


if __name__ == '__main__':
    main()
