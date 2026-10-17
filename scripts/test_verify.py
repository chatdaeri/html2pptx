"""verify.py 자체 점검 — 정상 PPTX 는 합격, 일부러 망가뜨린 PPTX 는 불합격이어야 한다.

    python3 test_verify.py deck.html deck.pptx      (deck.pptx = build.cjs 로 만든 정상 파일)
망가뜨리는 방법: ① 차트 c:txPr 중복 (실제로 PowerPoint '복구' 창을 띄웠던 패턴) ② 임베딩 글꼴 목록에서 하나 빼기 ③ 장표 하나 삭제
"""
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))


def run(html, pptx):
    r = subprocess.run([sys.executable, '-I', os.path.join(HERE, 'verify.py'), html, pptx], capture_output=True, text=True)
    return r.returncode, r.stdout


def mutate(src, dst, fn):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in zin.namelist():
            z.writestr(n, fn(n, zin.read(n)))


def main():
    html, good = sys.argv[1], sys.argv[2]
    code, out = run(html, good)
    assert code == 0, '정상 파일이 불합격:\n' + out
    tmp = tempfile.mkdtemp()

    def dup_txpr(n, d):
        if n == 'ppt/charts/chart1.xml':
            s = d.decode('utf8')
            m = re.search(r'<c:txPr>.*?</c:txPr>', s.split('</c:chart>')[1], re.S)
            return s.replace(m.group(0), m.group(0) * 2, 1).encode('utf8')
        return d

    def drop_font(n, d):
        if n == 'ppt/presentation.xml':
            return re.sub(r'<p:embeddedFont><p:font typeface="Pretendard"/>.*?</p:embeddedFont>', '', d.decode('utf8'), count=1, flags=re.S).encode('utf8')
        return d

    def drop_slide(n, d):
        if n == 'ppt/presentation.xml':
            return re.sub(r'<p:sldId [^>]*/>', '', d.decode('utf8'), count=1).encode('utf8')
        return d

    for name, fn, expect in (('txPr 중복', dup_txpr, '✗ 무결성'), ('글꼴 누락', drop_font, '✗ 글꼴'), ('장표 누락', drop_slide, '✗ 장표 수')):
        bad = os.path.join(tmp, name + '.pptx')
        mutate(good, bad, fn)
        code, out = run(html, bad)
        assert code == 1 and expect in out, f'{name}: 잡아내지 못함\n{out}'
        print(f'✓ {name} → 불합격으로 잡음')
    print('verify.py 자체 점검 통과')


if __name__ == '__main__':
    main()
