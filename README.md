# html2pptx

HTML 슬라이드 덱(`section.slide`, 1280×720)을 모든 요소가 편집 가능한 PPTX로 변환하는 Claude Code 스킬입니다.

- 글자·도형: 텍스트 상자·도형
- `.gtable`: 네이티브 표 · `[data-chart]`: 네이티브 차트(데이터 편집 가능)
- `[data-video]`: 모션그래픽 MP4를 슬라이드에 넣고 자동 재생
- 한글 글꼴 지정과 글꼴 임베딩
- 변환 후 자동 검증(`verify.py`)과 PowerPoint 렌더 비교(`--ppt`, macOS)

```bash
node scripts/build.cjs deck.html out.pptx --fontmap fonts/fontmap.json --ppt
python3 scripts/verify.py deck.html out.pptx
node scripts/record_motion.cjs deck.html --slide 1 --selector ".ring" --out assets/ring.mp4 --poster assets/ring.png
```

요구 사항: Node(전역 `pptxgenjs`·`playwright`·`sharp`), Python 3 (`python-pptx`·`lxml`·`Pillow`), ffmpeg. 자세한 내용은 `SKILL.md`와 `reference.md`를 봅니다.
