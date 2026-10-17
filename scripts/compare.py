"""HTML 덱 ↔ PPTX 비교 이미지 — 장표마다 왼쪽 HTML, 오른쪽 PowerPoint 렌더.

    python3 compare.py deck.html deck.pptx out_dir [--pptx-pdf rendered.pdf] [--dpi 60]

PPTX 렌더: macOS + PowerPoint 면 ppt_render.sh (실제 PowerPoint), 아니면 LibreOffice(soffice).
결과: out_dir/compare_01.png … + overview.png (6장씩) + 장표별 차이 점수(0=같음) 출력.
"""
import glob
import os
import re
import shutil
import subprocess
import sys

from PIL import Image, ImageChops, ImageDraw

CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
HERE = os.path.dirname(os.path.abspath(__file__))


# 인쇄 규칙이 없는 덱(뷰어형 · 다른 디자인 시스템)도 장표 1장 = PDF 1쪽으로 찍히게 하는 보조 CSS · 스크립트.
# 영상 자리([data-video])는 PPTX 와 같은 표지(data-poster)를 깔고, CSS 애니메이션은 첫 장면에 멈춘다(영상 표지 = 0초 장면).
PRINT_KIT = '''<style>
@page { size: 1280px 720px; margin: 0; }
@media print {
  html, body { display: block !important; width: 1280px !important; height: auto !important; margin: 0 !important; padding: 0 !important;
               gap: 0 !important; overflow: visible !important; background: none !important; }
  .slide { zoom: 1 !important; margin: 0 !important; break-after: page; break-inside: avoid; }
}
*, *::before, *::after { animation-delay: 0s !important; animation-play-state: paused !important; }
</style>
<script>document.addEventListener('DOMContentLoaded', () => document.querySelectorAll('[data-video][data-poster]').forEach((el) => {
  // 표지 = 장표 안에서 보이는 부분만 찍은 그림 → 늘리지 않고 원래 크기로, 장표 밖으로 나간 만큼 밀어서 깐다
  const r = el.getBoundingClientRect(), S = (el.closest('.slide') || document.body).getBoundingClientRect();
  el.style.background = 'url("' + el.getAttribute('data-poster') + '") ' + Math.max(0, S.left - r.left) + 'px ' + Math.max(0, S.top - r.top) + 'px / auto no-repeat border-box';   // 표지는 테두리까지 찍힌 그림
}));</script>
'''


def html_pdf(html, pdf):
    src = open(html, encoding='utf8').read()
    base = '<base href="file://' + os.path.dirname(os.path.abspath(html)) + '/">'
    m = re.search(r'<head[^>]*>', src, re.I)
    kit = src[:m.end()] + base + src[m.end():] if m else base + src
    kit = re.sub(r'</head>', PRINT_KIT + '</head>', kit, count=1, flags=re.I)
    tmp = os.path.join(os.path.dirname(os.path.abspath(pdf)), '_print.html')   # 결과 폴더에 임시 사본 (원본 덱은 건드리지 않음)
    open(tmp, 'w', encoding='utf8').write(kit)
    try:
        subprocess.run([CHROME, '--headless', '--disable-gpu', '--virtual-time-budget=10000', '--no-pdf-header-footer',
                        f'--print-to-pdf={pdf}', 'file://' + tmp], check=True, capture_output=True)
    finally:
        os.remove(tmp)


def pptx_pdf(pptx, pdf):
    r = subprocess.run(['bash', os.path.join(HERE, 'ppt_render.sh'), pptx, pdf], capture_output=True, text=True)
    if r.returncode == 0:
        return 'PowerPoint'
    print(r.stderr.strip())
    out = os.path.dirname(pdf)
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', out, pptx], check=True, capture_output=True)
    shutil.move(os.path.join(out, os.path.splitext(os.path.basename(pptx))[0] + '.pdf'), pdf)
    return 'LibreOffice'


def pages(pdf, prefix, dpi):
    subprocess.run(['pdftoppm', '-r', str(dpi), '-png', pdf, prefix], check=True)
    return sorted(glob.glob(prefix + '-*.png'))


def main():
    a = sys.argv[1:]
    html, pptx, out = a[0], a[1], a[2]
    dpi = int(a[a.index('--dpi') + 1]) if '--dpi' in a else 60
    os.makedirs(out, exist_ok=True)
    hp, pp = os.path.join(out, 'html.pdf'), os.path.join(out, 'pptx.pdf')
    html_pdf(html, hp)
    if '--pptx-pdf' in a:
        src = a[a.index('--pptx-pdf') + 1]
        if os.path.abspath(src) != os.path.abspath(pp):
            shutil.copy(src, pp)
        who = 'PDF 지정'
    else:
        who = pptx_pdf(pptx, pp)
    H, P = pages(hp, os.path.join(out, 'h'), dpi), pages(pp, os.path.join(out, 'p'), dpi)
    if len(H) != len(P):
        print(f'장표 수 다름: HTML {len(H)} · PPTX {len(P)}')
    sheets = []
    for i, (h, p) in enumerate(zip(H, P), 1):
        A, B = Image.open(h).convert('RGB'), Image.open(p).convert('RGB').resize(Image.open(h).size)
        diff = ImageChops.difference(A, B).convert('L')
        score = sum(diff.histogram()[32:]) / (A.size[0] * A.size[1]) * 100   # 눈에 띄게 다른 픽셀 %
        w, hh = A.size
        sh = Image.new('RGB', (w * 2 + 12, hh + 28), (210, 210, 210))
        sh.paste(A, (0, 28))
        sh.paste(B, (w + 12, 28))
        d = ImageDraw.Draw(sh)
        d.text((6, 8), f'#{i:02d} HTML', fill=(0, 0, 0))
        d.text((w + 18, 8), f'PPTX ({who})  diff {score:.1f}%', fill=(0, 0, 0))
        f = os.path.join(out, f'compare_{i:02d}.png')
        sh.save(f)
        sheets.append(f)
        print(f'#{i:02d} diff {score:5.1f}%')
    for k in range(0, len(sheets), 6):
        ims = [Image.open(s) for s in sheets[k:k + 6]]
        W, Hh = ims[0].size
        ov = Image.new('RGB', (W, Hh * len(ims)), (255, 255, 255))
        for j, im in enumerate(ims):
            ov.paste(im, (0, Hh * j))
        ov.save(os.path.join(out, f'overview_{k // 6 + 1}.png'))
    for f in glob.glob(os.path.join(out, 'h-*.png')) + glob.glob(os.path.join(out, 'p-*.png')):
        os.remove(f)
    print(f'비교 이미지 → {out}')


if __name__ == '__main__':
    main()
