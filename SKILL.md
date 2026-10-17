---
name: html2pptx
description: HTML 슬라이드 덱(section.slide, 1280×720)을 모든 요소가 편집 가능한 PPTX 로 변환한다 — 글자·도형은 텍스트 상자·도형, .gtable 은 네이티브 표(병합·강조·합계 선), data-chart 는 네이티브 차트(막대·누적·가로·선·이중축·도넛 · 데이터 편집 가능), 한글 글꼴(Pretendard) 지정, 글꼴 파일 임베딩. 선 굵기·크기 보정 없이 1px=0.75pt 그대로 옮기고 PowerPoint 렌더와 HTML 을 비교 이미지로 검증한다. pptx-skill 의 html2pptx 를 개조. "HTML 을 PPTX 로", "편집 가능한 PPT 로 변환", "html2pptx", "orange-proposal 덱 PPTX 로", "표·차트 편집되게 PPT 로" 등으로 트리거.
---

# html2pptx — HTML 덱 → 편집 가능한 PPTX

```bash
node ~/.claude/skills/html2pptx/scripts/build.cjs deck.html [out.pptx] --ppt     # 변환 + 검증(PowerPoint 렌더 · 비교 포함)
#   그 밖의 옵션: --fontmap fonts/fontmap.json · --no-embed · --loose · --keep · --no-verify
python3 ~/.claude/skills/html2pptx/scripts/verify.py deck.html out.pptx [--compare cmp/]   # 검증만 다시
```

**처음 실행하거나 실행 환경이 바뀌면 아래 사전 요구사항부터 점검한다.**
`orange-proposal` 스킬로 만든 덱은 그대로 변환된다 (규칙을 지킨 마크업 + `fonts/fontmap.json`).

## 사전 요구사항 — 변환 전에 반드시 확인

이 스킬은 아래 도구가 설치된 환경에서 동작한다. **스킬 폴더만 복사해도 의존 패키지가 설치되는 것은 아니다.** 최소 지원 버전은 아직 고정하지 않았다.

| 범위 | 설치 항목 | 역할 |
|---|---|---|
| 기본 변환 | Node.js + npm | `build.cjs` 실행 · 전역 모듈 검색 |
| 기본 변환 | npm 패키지 `pptxgenjs`, `playwright`, `sharp` | PPTX 생성 · HTML 측정 · 이미지/SVG 처리 |
| 기본 변환 | Python 3 + pip + venv | 표·차트·글꼴 처리 실행 환경 |
| 기본 변환 | Python 패키지 `python-pptx`, `lxml`, `Pillow` | 네이티브 요소 · XML · 이미지 처리 (`Pillow`는 비교에도 사용) |
| 브라우저 | macOS: Google Chrome / 그 외: Playwright Chromium | HTML 렌더링. Mac 코드는 `channel: chrome` 사용 |
| `--ppt` 검증 | Microsoft PowerPoint for Mac | 실제 PowerPoint에서 PDF 렌더링. 사용 가능한 라이선스·초기 설정 필요 |
| `--ppt` 비교 | Poppler (`pdfinfo`, `pdftoppm`) | PDF 쪽수 확인 · 장표 이미지 생성 |
| 선택 | LibreOffice (`soffice`) | `compare.py` 직접 실행 시 대체 렌더링. `--ppt`의 PowerPoint 검증을 대신하지 않음 |
| 글꼴 | 덱의 `fonts/fontmap.json` + 대응 TTF 파일 | 글꼴 지정·임베딩. `orange-proposal/fonts/`에 기본 10종 포함 |

현재 **전체 변환·비교 워크플로는 macOS 기준**이다. `compare.py`의 Chrome 경로는 `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`으로 고정돼 있고, PowerPoint 제어는 AppleScript를 사용한다. Windows·Linux에서는 기본 변환에 필요한 Chromium을 준비하고, 전체 비교에는 경로·렌더러 수정이 필요함을 먼저 안내한다. 별도 API 키는 필요 없다.

### 스킬 실행 시 누락 항목 처리

1. **변환 명령보다 먼저** 아래 읽기 전용 점검을 수행한다. 명령 자체가 없으면 해당 환경부터 누락으로 기록하고, 나머지 독립 항목은 계속 확인한다. 선택 도구는 요청한 작업에 필요한 경우만 확인한다.
2. 누락 항목이 있으면 `항목 | 필요한 이유 | 설치 명령`으로 안내한다. 사용자가 설치를 이미 승인한 범위는 실행하고, 그 외에는 패키지 목록·이유·설치 위치를 설명한 뒤 승인을 받는다. 승인 후 누락 항목만 설치한다. 단순 변환 요청을 패키지 설치 승인으로 간주하지 않는다.
3. 설치 실패 시 오류 원인을 알린다. 권한 오류를 피하려고 `sudo`, `--break-system-packages`를 임의로 사용하지 않는다. Homebrew가 없으면 먼저 해당 환경의 설치 방법을 안내한다. PowerPoint 로그인·라이선스·macOS 권한 설정은 사용자에게 안내한다.
4. 설치 후 같은 점검을 다시 수행하고 통과한 범위에서 변환한다. 필수 의존성이 없으면 변환을 시작하지 않는다. PowerPoint가 없으면 기본 변환 가능 여부와 실제 렌더 검증 미실시를 안내하고, `--ppt` 완료로 보고하지 않는다.

### 설치 여부 점검

```bash
command -v node npm python3
node --version
python3 --version
# build.cjs와 동일하게 전역 모듈 또는 기존 NODE_PATH에서 검색
NODE_PATH="${NODE_PATH:-$(npm root -g)}" node -e 'for (const p of ["pptxgenjs", "playwright", "sharp"]) { try { require(p); console.log("OK", p); } catch (e) { console.error("FAIL", p, e.message); process.exitCode = 1; } }'
# -I는 사용자 site-packages와 PYTHONPATH를 제외하므로 실제 변환과 같은 옵션으로 확인
python3 -I -c 'import pptx, lxml.etree, PIL; print("Python packages OK")'
# macOS 브라우저
test -x '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
# --ppt 요청 시 추가
command -v pdfinfo pdftoppm
test -d '/Applications/Microsoft PowerPoint.app'
# LibreOffice 대체 렌더링을 사용할 때만
command -v soffice
```

앱 파일 존재만으로 PowerPoint 라이선스·자동화 권한이 확인되지는 않는다. 실행 앱의 PowerPoint/System Events 자동화 권한, 확인 창 감지용 손쉬운 사용 권한, PowerPoint 컨테이너 임시 폴더 쓰기 권한이 필요할 수 있다. 다른 PowerPoint 문서는 저장·닫은 상태에서 렌더한다.

### 승인 후 설치 명령 (macOS 예시)

아래는 설치 예시이며 **모두 일괄 실행하지 않는다.** 이미 있는 도구는 유지한다. Homebrew는 시스템 도구 설치를 위한 선택 수단이며 변환 자체의 필수 의존성은 아니다.

```bash
# Homebrew가 이미 있을 때: 없는 도구만 선택
brew install node python
brew install --cask google-chrome
# 기본 Node 패키지: build.cjs가 전역 위치를 검색함
npm install -g pptxgenjs playwright sharp
# Python: 시스템 Python을 수정하지 않고 전용 가상환경 사용
python3 -m venv "$HOME/.venvs/html2pptx"
source "$HOME/.venvs/html2pptx/bin/activate"
python3 -m pip install python-pptx lxml Pillow
# --ppt 비교를 요청한 경우
brew install poppler
# PowerPoint 대체 렌더링이 필요한 경우만
brew install --cask libreoffice
```

- 이미 전용 가상환경이 있으면 재생성하지 않고 활성화한다. **이후 변환 때도 가상환경을 활성화한 같은 셸에서 `node …/build.cjs`를 실행한다.** 자식 프로세스가 PATH의 `python3`를 사용한다. `pip --user` 설치는 `python3 -I`에서 보이지 않을 수 있다.
- 전역 npm 설치가 불가능하면 승인받은 별도 폴더에 `npm install --prefix <설치폴더> pptxgenjs playwright sharp`로 설치하고, 변환할 때 `NODE_PATH=<설치폴더>/node_modules`를 지정한다.
- macOS는 설치된 Chrome을 사용하므로 Playwright Chromium 추가 다운로드가 필요 없다. 다른 OS의 기본 변환은 설치된 Playwright CLI로 `playwright install chromium`을 실행한 뒤 실제 브라우저 실행 여부를 확인한다.
- 글꼴 TTF와 `fontmap.json` 경로도 확인한다. 차트 축·SVG 텍스트까지 같은 모양으로 렌더하려면 사용 글꼴의 OS 설치가 필요할 수 있다(아래 알려진 한계 참고).

이 절차는 **스킬을 읽는 에이전트가 수행**한다. `build.cjs` 자체는 패키지를 자동 설치하지 않는다.

## 무엇이 무엇으로

| HTML | PowerPoint | 편집 |
|---|---|---|
| `p` · `h1~h6` (인라인 `span` · `b` · `em` · `mark` · `a` · `code` 등 포함) | 텍스트 상자 · 글자 조각별 글꼴/크기/색 · 자간 · `mark`(형광펜)는 글자 뒤 띠 도형(줄마다 1개, HTML 과 같은 두께·위치) | 글자 수정 · 띠는 따로 이동 |
| 배경·테두리가 있는 `p` (배지 · 번호 원 · 버튼 · pill) | 글자를 품은 도형 (원 · 캡슐 · 둥근 사각형) | 도형 · 글자 |
| 배경·테두리가 있는 `div` · `i` · `span`(글자 없음) · 구조형 `li` | 사각형 · 둥근 사각형 · 원(정사각 + 50%) / 한쪽 테두리는 선 | 도형 |
| `ul.bullets` · `ul.ticks` (단순 목록) | 네이티브 불릿 (`li::before` 글자 • ✓) | 목록 |
| `.gtable` (div 그리드) | **네이티브 표** — 열 너비 · 행 높이 · 병합 · 칸 색 · 가로선 굵기/색 그대로 | 표 |
| `[data-chart]` | **네이티브 차트** — 내장 엑셀(데이터 편집) · 강조 막대/조각 · 레이블 · 평균선 · 사건선 · 이중축(막대 + 오른쪽 보조 축 꺾은선) | 차트 데이터 |
| `img` (사진) | 그림 — `object-fit`·`object-position` 대로 **자르기**(원본 보존), `border-radius` 둥근 모서리 · 원 | 그림 바꾸기 · 다시 자르기 |
| 배경 이미지 `background-image: url()` | 상자: 같은 방식의 그림(면 색 → 그림 → 테두리 순서) / 장표: 슬라이드 배경 그림 | 그림 · 배경 서식 |
| SVG (인라인 · `<img src=.svg>`) | **벡터 그림**(SVG 원본 + 대체 PNG) — CSS 색·변수도 굳혀서 옮김 | 크기 · 색, '도형으로 변환' 후 조각별 편집 |
| `[data-video]` (CSS 애니메이션 모션그래픽) | **동영상**(MP4) + 재생 전 표지 그림 — 자리·크기는 HTML 상자 그대로, 장표 밖으로 나간 부분은 잘림 | 영상 바꾸기 · 재생 옵션 |

- 크기: 1280×720px = 13.333×7.5in(와이드). 1px = 0.75pt 그대로 — 축소 보정이 없어 선 굵기가 바뀌지 않는다.
- 겹침 순서: 표·차트도 HTML 과 같은 z 순서 자리에 끼워 넣는다. 차트 위에 남길 글자는 `data-pptx-keep`.
- 글꼴: `fontmap.json` 이 CSS(family + weight) → PowerPoint 글꼴 이름(+굵게) · 한글(ea) 글꼴 · 임베딩 파일을 정한다. 굵기는 가짜 굵게가 아니라 해당 굵기 글꼴(Medium · SemiBold)을 쓴다.

## 모션그래픽(영상) 넣기

PowerPoint 는 CSS 애니메이션을 못 움직이므로 **MP4 로 렌더해서 영상으로 넣는다**.

```bash
# 1) 모션 요소를 MP4 + 표지 PNG 로 (프레임마다 애니메이션을 멈춰 캡처 → ffmpeg H.264)
node scripts/record_motion.cjs deck.html --slide 1 --selector ".ring" --out assets/ring.mp4 --poster assets/ring.png --seconds 30 --fps 20
```
```html
<!-- 2) 덱 HTML 에서 그 요소를 영상 자리로 바꿈 (위치·크기는 CSS 로 지정) -->
<div class="ring" style="position:absolute; right:40px; top:110px; width:500px; height:500px"
     data-video="assets/ring.mp4" data-poster="assets/ring.png"></div>
```
- `data-poster` 를 생략하면 변환 시점의 화면을 표지로 캡처한다.
- 한 번 도는 주기(루프)가 정확히 맞는 길이로 렌더한다(회전 30초면 `--seconds 30`). 긴 모션(예: 70초 스크롤)은 보여줄 구간만 잘라 쓴다.
- MP4 는 배경이 투명할 수 없다 — 요소 뒤 배경색까지 함께 찍히므로 배경이 단색인 곳에 둔다.
- 슬라이드쇼에서 장표가 열리면 **자동 재생 · 반복 재생 · 소리 끔**으로 넣는다(PowerPoint '시작: 자동' 과 같은 타이밍). 끄려면 `data-autoplay="false"`(클릭해야 재생) · `data-loop="false"`(한 번만). 장표에 다른 애니메이션 타이밍이 이미 있으면 자동 재생을 넣지 않고 경고한다.
- 변환 후 `verify.py` 가 `동영상 N/N · 자동 재생 N/N` 과 빈 영상 파일을 검사한다.
- 한계: 영상만큼 파일이 커진다 · Keynote · Google 슬라이드는 재생이 안 될 수 있다 · 정지 화면 비교 이미지에는 표지만 보인다.

## 파이프라인

```
deck.html ─① html2pptx.cjs (Chrome 측정) → 글자·도형·선 = PptxGenJS ─┐
                    └→ 표·차트 자리 + 데이터 (natives.json) ──────────┤
                                                                     ② native.py (python-pptx)
                                                                       · 네이티브 표 · 차트
                                                                       · 한글 ea 글꼴
                                                                       · 글꼴 EOT 임베딩
                                                                     → deck.pptx
③ verify.py : 합격/불합격 — 파일 · 무결성 · 장표 수 · 네이티브 수 · 그림 · 글자 · 영상 · 글꼴 (+ --ppt: PowerPoint 렌더 · 비교)
```

## 검증 (필수) — 변환 직후 자동 실행

| # | 항목 | 합격 기준 | 잡아내는 문제 |
|---|---|---|---|
| 0 | 변환 규칙 | `build.cjs` ① 단계가 위반 없이 끝남 | 그라데이션 · 감싸지 않은 글자 · 인라인 여백 · 읽지 못한 그림 · 크기 불일치 |
| 1 | 파일 | zip · XML 정상, python-pptx 로 열림 | 깨진 파일 |
| 2 | 무결성 | 차트 · 표 · 프레젠테이션 XML 의 단일 요소 중복 없음 | PowerPoint **'복구' 창** (실제 사례: 차트 `c:txPr` 중복) |
| 3 | 장표 수 | HTML `section.slide` 수 = PPTX 장표 수 | 빠진 장표 |
| 4 | 네이티브 | `.gtable` 수 = 표 수 · `data-chart` 수 = 차트 수(전부 엑셀 내장) | 표·차트가 그림/도형으로 바뀜, 편집 불가 |
| 4b | 그림 | HTML SVG 수 = 벡터 그림 수(SVG 원본 포함) · 대체 PNG 정상 · 사진 수 ≥ HTML `img` 수 | SVG 가 빠지거나 그림으로만 들어감 |
| 5 | 글자 | HTML 장표의 글자 조각(SVG · 차트 · 영상 안 제외)이 모두 같은 PPTX 장표에 있음 | 모르는 태그(예: 예전의 `mark`) 안 글자가 조용히 사라짐 |
| 6 | 영상 | `data-video` 수 = 동영상 수 = 자동 재생 수(`data-autoplay="false"` 제외) · 빈 영상 파일 없음 | 영상 누락 · 클릭해야 재생 |
| 7 | 글꼴 | 글자에 쓰인 글꼴이 모두 EOT 로 임베딩 | 받는 PC 에서 글꼴 바뀜 |
| 8 | PowerPoint (`--ppt`) | 실제 PowerPoint 가 복구 창 없이 열고 PDF 쪽수 = 장표 수 | 1~7 로 못 잡는 손상 |
| 9 | 비교 (`--ppt`) | HTML ↔ PowerPoint 장표별 차이 ≤ 5% (보통 1~3.5%) · `out_compare/` 이미지 | 모양 · 위치 · 줄바꿈 차이 |

- 비교용 HTML 인쇄는 덱에 인쇄 규칙(`@page`)이 없어도 장표 1장 = 1쪽으로 찍는다. 영상 자리에는 표지(`data-poster`)를 깔고 CSS 애니메이션은 첫 장면에 멈춘다.
- 무결성 검사는 장표 안 도형 번호(id) 중복도 본다. PptxGenJS 가 영상에 겹치는 번호를 줘 PowerPoint 복구 창이 뜬 사례가 있어(2026-10-17), `native.py` 가 번호를 다시 매긴다.

- 불합격이면 `build.cjs` 가 종료 코드 2 로 끝난다. ✗ 항목을 고치고 다시 변환한다.
- **7번 비교 이미지는 차이 %와 상관없이 전 장 눈으로 확인한다** — 숫자는 어디를 볼지 알려 줄 뿐이다.
- PowerPoint 단계는 다른 문서가 열려 있으면 건너뛴다(✓ 표시 + "건너뜀"). 이때는 문서를 닫고 `verify.py … --compare` 로 다시 돌린다.
- 검증기 자체 점검: `python3 scripts/test_verify.py deck.html deck.pptx` (일부러 망가뜨린 파일 3종을 불합격으로 잡는지)

## PowerPoint 렌더 안전 규칙 (`ppt_render.sh`)

- **PowerPoint 에 다른 문서가 열려 있으면 실행하지 않는다** (종료 코드 3 → `compare.py` 는 LibreOffice 로 대신).
  파일에 문제가 있으면 PowerPoint 가 '복구' 창을 띄우거나 비정상 종료해 열린 문서가 모두 닫힐 수 있다 — 2026-10-06 실제 발생.
- 이 파일 이름의 프레젠테이션만 열고 · 저장하고 · 닫는다. `active presentation` 을 쓰지 않는다.
- 확인 창(복구 등)이 떠 있으면 **버튼을 누르지 않고** 문구만 읽어 알린 뒤 멈춘다 (종료 코드 4). 버튼을 누르면 PowerPoint 가 종료될 수 있다.

## 알려진 차이 · 한계

- 여러 줄 글자의 줄바꿈 위치가 한 단어 정도 다를 수 있다 (한 줄 글자 · 모든 줄을 `<br>` 로 나눈 글자는 줄바꿈 금지로 고정).
- `mark` 형광펜: 일반 글자 상자 안이면 HTML 에서 잰 자리에 띠 도형을 글자 뒤에 깐다 — PowerPoint 에서 글자를 고쳐 줄이 바뀌면 띠는 따라오지 않는다. 목록 · 표 칸 · 배지 안의 `mark` 는 글자 높이 전체를 칠하는 강조색으로 대신한다.
- 큰 제목 글자는 PowerPoint 에서 몇 px 위에 그려질 수 있다(고정 줄 간격의 글자 기준선 차이). 띠와 글자 사이가 그만큼 벌어진다.
- 모르는 인라인 태그 안의 글자는 버리지 않고 오류로 알린다. `verify.py` 의 '글자' 항목이 HTML 글자 조각이 모두 PPTX 에 있는지 다시 확인한다.
- 라틴 글꼴에 없는 기호(General Sans 의 → 등)는 PowerPoint 가 다른 글꼴로 그린다.
- 변환 불가 → 오류로 알림: 그라데이션 · 인라인 여백. inset 그림자 · `overflow:hidden` 자르기 · 모서리별 다른 반지름 · 반복 배경(repeat) 은 옮겨지지 않는다.
- SVG: 안의 글자는 PowerPoint 가 설치된 글꼴로 그린다 · SVG 자체의 CSS transform 은 무시 · 그림 파일은 로컬 경로 또는 data: 만(웹 주소 불가).
- 차트: 둥근 막대 없음(직각), 축 항목 하나만 굵게 불가, 가로 막대 항목 이름은 오른쪽 정렬.
- 차트 축 눈금 · 항목 이름: PowerPoint(Mac)가 임베딩 글꼴을 쓰지 않아, Inter 가 설치되지 않은 PC 에서는 Calibri 로 그린다(글자 폭 약 20% 좁음). 데이터 레이블 · 텍스트 상자는 임베딩 글꼴로 정상 — 2026-10-14 실험 확인.
- 글꼴 임베딩은 PowerPoint 형식(EOT 포장)으로 넣는다 — 원본 TTF 를 그대로 넣으면 PowerPoint 가 무시한다. 10종 약 5MB.

상세 규칙 · 차트 매핑 · 문제 해결: `reference.md`
