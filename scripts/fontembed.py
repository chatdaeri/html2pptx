"""PPTX 글꼴 임베딩 — 표준 라이브러리만 사용.
PowerPoint 임베디드 글꼴(ppt/fonts/*.fntdata)은 EOT(Embedded OpenType) 컨테이너.
ttf_to_eot(): TrueType → 비압축 EOT 2.1. embed(): presentation.xml 에 embeddedFontLst 추가."""
import re
import struct
import zipfile


def _tables(d):
    n = struct.unpack('>H', d[4:6])[0]
    t = {}
    for i in range(n):
        tag, _cs, off, ln = struct.unpack('>4sIII', d[12 + 16 * i:28 + 16 * i])
        t[tag.decode('latin1')] = (off, ln)
    return t


def font_names(d):
    t = _tables(d)
    off, _ = t['name']
    _fmt, cnt, so = struct.unpack('>HHH', d[off:off + 6])
    out = {}
    for i in range(cnt):
        pid, _eid, lid, nid, ln, o = struct.unpack('>HHHHHH', d[off + 6 + 12 * i:off + 18 + 12 * i])
        if pid == 3 and lid == 0x409 and nid not in out:
            out[nid] = d[off + so + o:off + so + o + ln].decode('utf-16-be')
    return out


def ttf_to_eot(d):
    t = _tables(d)
    if 'glyf' not in t:
        raise ValueError('TrueType 윤곽(glyf)이 없는 글꼴은 PowerPoint에 넣을 수 없습니다 (OTF/CFF)')
    o2 = t['OS/2'][0]
    weight, fs_type = struct.unpack('>H', d[o2 + 4:o2 + 6])[0], struct.unpack('>H', d[o2 + 8:o2 + 10])[0]
    panose = d[o2 + 32:o2 + 42]
    ur = struct.unpack('>IIII', d[o2 + 42:o2 + 58])
    fs_sel = struct.unpack('>H', d[o2 + 62:o2 + 64])[0]
    cpr = struct.unpack('>II', d[o2 + 78:o2 + 86]) if struct.unpack('>H', d[o2:o2 + 2])[0] >= 1 else (0, 0)
    csa = struct.unpack('>I', d[t['head'][0] + 8:t['head'][0] + 12])[0]
    nm = font_names(d)
    body = struct.pack('<10sBBIHH', panose, 1, fs_sel & 1, weight, fs_type, 0x504C)
    body += struct.pack('<IIII', *ur) + struct.pack('<II', *cpr) + struct.pack('<I', csa)
    body += b'\0' * 16 + b'\0\0'                      # Reserved1-4, Padding1
    for i, nid in enumerate((1, 2, 5, 4)):           # Family, Style, Version, Full
        s = nm.get(nid, '').encode('utf-16-le')
        body += (b'\0\0' if i else b'') + struct.pack('<H', len(s)) + s
    body += b'\0\0' + struct.pack('<H', 0)           # Padding5, RootStringSize
    size = 16 + len(body) + len(d)
    return struct.pack('<IIII', size, len(d), 0x00020001, 0) + body + d


def embed(pptx_in, pptx_out, faces, raw_ttf=False):
    """faces: {typeface: {'regular': bytes, 'bold': bytes, ...}} — 값은 TTF 또는 EOT 바이트."""
    zin = zipfile.ZipFile(pptx_in)
    files = {n: zin.read(n) for n in zin.namelist()}
    pres = files['ppt/presentation.xml'].decode('utf8')
    rels = files['ppt/_rels/presentation.xml.rels'].decode('utf8')
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    nxt = max(ids or [0]) + 1
    lst = '<p:embeddedFontLst>'
    k = 0
    for face, styles in faces.items():
        lst += f'<p:embeddedFont><p:font typeface="{face}"/>'
        for style in ('regular', 'bold', 'italic', 'boldItalic'):
            if style not in styles:
                continue
            k += 1
            data = styles[style]
            if not raw_ttf and data[34:36] != b'\x4c\x50':   # EOT 매직(0x504C) 없으면 변환
                data = ttf_to_eot(data)
            name = f'ppt/fonts/font{k}.fntdata'
            files[name] = data
            rid = f'rId{nxt}'
            nxt += 1
            rels = rels.replace('</Relationships>',
                                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font{k}.fntdata"/></Relationships>')
            lst += f'<p:{style} r:id="{rid}"/>'
        lst += '</p:embeddedFont>'
    lst += '</p:embeddedFontLst>'
    pres = re.sub(r'<p:embeddedFontLst>.*?</p:embeddedFontLst>', '', pres, flags=re.S)
    m = re.search(r'<p:notesSz[^>]*/>', pres)
    pres = pres[:m.end()] + lst + pres[m.end():]
    for attr in ('embedTrueTypeFonts="1"', 'saveSubsetFonts="0"'):   # 이미 있으면 값만 교체
        key = attr.split('=')[0]
        if f'{key}="' in pres:
            pres = re.sub(rf'{key}="[^"]*"', attr, pres, count=1)
        else:
            pres = pres.replace('<p:presentation ', f'<p:presentation {attr} ', 1)
    ct = files['[Content_Types].xml'].decode('utf8')
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
    files['ppt/presentation.xml'] = pres.encode('utf8')
    files['ppt/_rels/presentation.xml.rels'] = rels.encode('utf8')
    files['[Content_Types].xml'] = ct.encode('utf8')
    with zipfile.ZipFile(pptx_out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
        for n, b in files.items():
            z.writestr(n, b)
