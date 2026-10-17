# html2pptx 상세 참고

## 1. HTML 규칙 (변환 전제)

- 장표 = `section.slide` (1280×720). 한 HTML 에 여러 장. 장표 배경색 = 장표 면 색.
- 글자는 `p` · `h1~h6` · `li` 안에만. `div` 에 글자를 바로 넣으면 오류.
- 인라인(`span` 등)에 margin 이 있으면 오류 — PowerPoint 는 인라인 여백이 없다.
- 장식 선·점은 실제 요소로. `::before/::after` 는 단순 목록의 글자 불릿으로만 읽는다(`content:"•"`).
- 단순 목록 = 모든 `li` 가 인라인 글자 또는 `p` 하나이고 배경·테두리가 없음 → 네이티브 불릿 텍스트. 그 외 목록은 구조형(칸마다 도형 + 글자).
- 표: `.gtable` 또는 `[data-pptx-table]` 의 **직계 자식 = 칸**. 병합은 `grid-row/column: span N`. 칸 안에는 글자만(도형 금지).
- 차트: `[data-chart]` JSON. 그림 영역은 `[data-plot-box]` 요소(들)의 합집합 — orange-proposal 의 charts.js 가 붙인다.
- 그림: `img` · `background-image: url()` 은 화면에 보이는 영역만 틀로 잡고 원본 전체 + 자르기 값(srcRect)을 넣는다 → PowerPoint '자르기' 로 다시 조정 가능. 장표(`section.slide`) 배경 그림은 sharp 로 잘라 슬라이드 배경으로.
- SVG: 인라인 SVG 는 계산된 색(CSS 클래스 · 변수 · currentColor)을 속성으로 굳혀 독립 SVG 로 만든 뒤, sharp 로 2배 대체 PNG 를 그리고 native.py 가 SVG 원본을 `asvg:svgBlip` 으로 붙인다 (PowerPoint 2016+ 표준 방식). `data-chart` 안의 SVG 는 차트로 처리되므로 해당 없음.
- 글꼴 대응: `<meta name="pptx-fontmap" content="fonts/fontmap.json">` 또는 `--fontmap`.

## 2. fontmap.json

```json
{
  "faces": { "Inter": { "400": ["Inter 18pt", false], "600": ["Inter 18pt SemiBold", false], "700": ["Inter 18pt", true] } },
  "ea":    { "Inter 18pt": "Pretendard", "Inter 18pt SemiBold": "Pretendard SemiBold" },
  "files": { "Inter 18pt": { "regular": "Inter_18pt-Regular.ttf", "bold": "Inter_18pt-Bold.ttf" } }
}
```

- `faces`: CSS family → weight → [PowerPoint 글꼴 이름(name ID 1), bold]. 없는 weight 는 가장 가까운 값.
- `ea`: 라틴 글꼴 이름 → 한글 글꼴 이름 (PowerPoint 는 한글 글자에 `<a:ea>` 글꼴을 쓴다).
- `files`: 글꼴 이름 → regular/bold/italic/boldItalic TTF (fontmap 폴더 기준). 실제로 쓴 글꼴만 넣는다.
- PowerPoint 가 쓰는 글꼴 이름 = TTF 의 name ID 1 (예: Inter 18pt SemiBold). TrueType(glyf) 만 넣을 수 있다 — OTF(CFF) 불가.

## 3. 차트 매핑 (charts.js 데이터 → PowerPoint)

| type | PowerPoint | 맞춘 것 |
|---|---|---|
| bar | 묶은 세로 막대 | 플롯 영역 = HTML `.plot` 위치(수동 레이아웃) · gapWidth = 실측 막대 폭 · 값축 숨김 0~max · 강조 점 오렌지 · 값 레이블(강조는 General Sans 15pt) · `ref` → 같은 축 점선 꺾은선 계열 |
| stack | 누적 세로 막대 | 계열 Charcoal/Ash/Pebble · 구간 틈 = 흰 테두리 1.5pt · 강조 점 · 합계 = 선 없는 꺾은선 + 위쪽 레이블 |
| hbar | 묶은 가로 막대 | 항목 역순(첫 항목 위) · 막대 두께/간격 실측 · 값축 0~max · 레이블 막대 끝 |
| line | 꺾은선 | min/max/step 눈금 · 격자 Pebble 0.75pt · midCat(점이 플롯 가장자리) · 강조 2.25pt 오렌지 + 끝점 원 · 점선 계열 · 끝점에만 "이름 값" 레이블 · labelEvery → tickLblSkip · note → 점선 세로선 도형 |
| dual | 묶은 세로 막대 + 꺾은선(보조 축) | 막대 = 왼쪽 축 0~max · step 눈금 · 격자 Pebble 0.75pt · gapWidth = 실측 막대 폭(`[data-bar]`) · 강조 점 오렌지 · `bar.labels` 막대 위 레이블 / 선 = 막대 차트 계열을 `lineChart` 로 옮겨 보조 축(오른쪽 `valAx` crosses=max + 숨긴 `catAx`)에 연결 · Ink 1.5pt · 원 표식 · 점 위 레이블 / 단위 · 범례는 HTML 글자 · 도형(`data-pptx-keep`) |
| donut / ring | 도넛 | 12시 시작 · 조각 색 순서 · holeSize = HTML 링 두께 비율 · 조각 틈 = 흰 테두리 |

표시 형식(`format`) 끝의 `%` 는 엑셀에 `"%"` 글자로 넣는다 — charts.js 처럼 값 뒤에 기호만 붙고 ×100 하지 않는다(`xl_fmt`).

그림 영역이 차트 틀 밖으로 나가면 틀을 넓힌다(`frame_fit`) — 그러지 않으면 PowerPoint 가 플롯을 틀 안으로 줄인다.

## 4. 구조 점검 코드 (verify.py 4번 항목과 같은 내용 — 장표별로 볼 때)

```python
from pptx import Presentation
p = Presentation('deck.pptx')
for i, s in enumerate(p.slides, 1):
    t = sum(sh.has_table for sh in s.shapes); c = [sh for sh in s.shapes if sh.has_chart]
    pic = sum(sh.shape_type == 13 for sh in s.shapes)
    print(i, '표', t, '차트', len(c), '엑셀내장', sum(x.chart.part.chart_workbook.xlsx_part is not None for x in c), '그림', pic)
```

## 5. 문제 해결

| 증상 | 원인 · 해결 |
|---|---|
| PowerPoint 가 "내용에 문제 … 복구" | XML 요소 중복 또는 순서 위반. 실제 사례: 차트 `c:txPr` 중복(python-pptx 기본값 + 추가) → 교체로 해결. `--keep` 으로 `.base.pptx` 를 따로 열어 1단계/2단계 중 어디인지 가린다 |
| AppleScript 열기 오류 -9074 · ppt_render 종료 코드 4 | PowerPoint 에 확인 창(복구 등)이 떠 있다. 버튼을 누르지 말고 문구를 사용자에게 알린다 — 사용자가 창을 닫은 뒤 `verify.py` 2번(무결성) 결과부터 본다 |
| 배지 글자가 두 줄 | 한 줄 판정(브라우저가 그린 줄 수)이 1 이면 줄바꿈 끔 — 판정이 빗나가면 `white-space:nowrap` |
| 글꼴이 Calibri 로 보임 | fontmap 에 없는 family, 또는 임베딩 파일이 TTF 그대로(EOT 아님) — fontembed.py 가 EOT 로 포장 |
| 가로 막대 간격이 좁음 | 그림 영역이 틀 밖 → frame_fit 확인 |

## 6. 파일

| 파일 | 역할 |
|---|---|
| `scripts/build.cjs` | 진입점 (1·2단계 연결, fontmap 자동 탐지) |
| `scripts/html2pptx.cjs` | 측정 · 글자/도형 변환 · 표/차트 자리 추출 (pptx-skill 원본 개조) |
| `scripts/native.py` | 네이티브 표 · 차트 · ea 글꼴 · 임베딩 |
| `scripts/fontembed.py` | TTF → EOT, presentation.xml 임베딩 목록 |
| `scripts/verify.py` | 합격/불합격 검증 (build.cjs 가 자동 실행) |
| `scripts/test_verify.py` | 검증기 자체 점검 (망가뜨린 파일 3종) |
| `scripts/compare.py` | HTML ↔ PPTX 비교 이미지 |
| `scripts/ppt_render.sh` | PowerPoint(macOS) PDF 렌더 — 안전 규칙 포함 |
