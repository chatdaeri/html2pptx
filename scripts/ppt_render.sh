#!/bin/bash
# PowerPoint(macOS)로 PPTX 를 열어 PDF 로 저장 — 실제 PowerPoint 렌더링 검증용.
#   bash ppt_render.sh deck.pptx out.pdf
# 안전 규칙: 이 파일 이름의 프레젠테이션만 열고 · 저장하고 · 닫는다 (다른 열린 문서는 건드리지 않음).
# PowerPoint 샌드박스 안 임시 폴더를 거쳐 권한 창 없이 동작한다.
# 종료 코드: 0 성공 · 2 렌더 실패 · 3 다른 문서가 열려 있어 건너뜀 · 4 확인 창(복구 등) — 사용자 확인 필요
set -euo pipefail
src="$1"; out="$2"
# 안전장치: 다른 문서가 열려 있으면 실행하지 않는다. 파일 손상 시 PowerPoint 가 '복구' 창을 띄우거나
# 비정상 종료될 수 있고, 그러면 열려 있던 다른 문서까지 닫힌다 (2026-10-06 실제 발생).
if pgrep -xq "Microsoft PowerPoint"; then
  open_n=$(osascript -e 'tell application "Microsoft PowerPoint" to count of presentations' 2>/dev/null || echo 1)
  if [[ "$open_n" != "0" && "${PPT_RENDER_FORCE:-}" != "1" ]]; then
    echo "PowerPoint 에 문서 ${open_n}개가 열려 있어 렌더를 건너뜁니다. 작업 문서를 저장 · 닫은 뒤 다시 실행하세요." >&2
    exit 3
  fi
fi
# 확인 창(복구 등) 문구 읽기 — 읽기만 하고 버튼은 절대 누르지 않는다 (누르면 PowerPoint 가 종료될 수 있음)
dialog_text() {
  osascript -e 'tell application "System Events" to tell process "Microsoft PowerPoint"
    set t to ""
    repeat with w in (every window whose subrole is "AXDialog")
      try
        set t to t & (value of every static text of w as text) & " | 버튼: " & (name of every button of w as text)
      end try
    end repeat
    return t
  end tell' 2>/dev/null || true
}
if pgrep -xq "Microsoft PowerPoint"; then
  d=$(dialog_text)
  if [[ -n "$d" ]]; then echo "PowerPoint 에 확인 창이 떠 있어 렌더를 멈춥니다 (사용자 확인 필요): $d" >&2; exit 4; fi
fi
box="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/tmp/html2pptx"
mkdir -p "$box"
name="render_$$_$(basename "$src")"
cp "$src" "$box/$name"
rm -f "$box/${name%.pptx}.pdf"
res=$(osascript <<EOF 2>&1
tell application "Microsoft PowerPoint"
  try
    open POSIX file "$box/$name"
  on error e
    return "OPEN-ERROR " & e
  end try
  set ok to false
  repeat 30 times
    if (name of every presentation) contains "$name" then
      set ok to true
      exit repeat
    end if
    delay 1
  end repeat
  if not ok then return "OPEN-ERROR not opened"
  set p to presentation "$name"
  save p in (POSIX file "$box/${name%.pptx}.pdf") as save as PDF
  close p saving no
  return "OK"
end tell
EOF
)
if [[ "$res" != "OK" ]]; then
  d=$(dialog_text)
  # 확인 창이 남아 있으면 파일을 지우지 않는다 (창이 그 파일을 붙잡고 있음). 창은 사용자가 닫는다.
  if [[ -n "$d" ]]; then echo "PowerPoint 가 파일을 열지 못하고 확인 창을 띄웠습니다 — 파일 손상 가능. 창은 누르지 말고 사용자에게 알리세요: $d" >&2; exit 4; fi
  rm -f "$box/$name"; echo "PowerPoint 렌더 실패: $res" >&2; exit 2
fi
rm -f "$box/$name"
mv "$box/${name%.pptx}.pdf" "$out"
echo "PowerPoint PDF → $out"
