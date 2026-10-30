# skill-test

## pptx-to-dynamic-html

파워포인트(.pptx) 슬라이드를 원본 디자인을 유지한 채 스크롤형 다이나믹 HTML 한 페이지로 변환하는 스킬입니다. 제안발표용으로, 원본 충실도를 가장 중요하게 보고 효과는 가볍게 유지합니다.

- PPT 등장 애니메이션을 순서대로 재현
- 애니메이션이 없는 슬라이드는 읽는 순서대로 자동 등장 효과 적용 (`--no-auto-animate`로 끄기)
- 말풍선 도형은 꼬리가 가리키는 영역에 마우스를 올렸을 때만 표시
- 화살표는 흐르는 형태로 표현 (연결선은 대시가 흐르고, 블록 화살표는 빛이 지나감)
- 마우스를 올린 요소는 살짝 커지며 하이라이트 (`--no-hover`로 끄기)
- 마우스 움직임에 따른 가벼운 패럴랙스
- 차트·표는 원본 렌더링을 잘라 이미지로 넣어 모양 유지

### 사용법

Claude에 .pptx를 첨부하고 "웹페이지로 만들어줘"처럼 요청하면 됩니다. 직접 실행할 때는:

```bash
python scripts/extract_pptx.py deck.pptx WORK
python scripts/render_reference.py deck.pptx WORK     # LibreOffice 필요
python scripts/build_html.py WORK out.html --title "덱 제목"
python scripts/compare.py out.html WORK                # 원본과 나란히 비교
```

필요한 것: Python 패키지 python-pptx, lxml, pillow, playwright / LibreOffice, pdftoppm

### 폴더 구성

- `SKILL.md`: 스킬 지침
- `scripts/`: 추출·렌더링·빌드·비교 스크립트
- `assets/template.html`: 페이지 틀 (효과 CSS와 JS)
