---
name: pptx-editing
description: >-
  Build, edit, and QA PowerPoint (.pptx) decks with python-pptx without the common
  failures that make a deck look broken or lose work. Use this whenever a task
  involves a .pptx file, a PowerPoint / 파워포인트 / 발표자료 / 슬라이드 / 덱 /
  lab-meeting deck / conference slides — including building from a template,
  editing slides, tables, figures, and especially SPEAKER NOTES (발표자 노트),
  fitting images without overflow, and rendering to PDF/PNG for a visual check.
  Trigger even if the user only says "이 슬라이드 고쳐줘" / ".pptx" — naive
  python-pptx edits silently drop notes, overflow the slide, or get wiped on rebuild.
---

# PPTX editing (pointer)

PowerPoint `.pptx` 덱을 **손편집을 지우지 않고** 만들고 고치는 워크플로.

- **트리거:** `.pptx` / 발표자료 / 슬라이드 작업. 사용자가 내부 사정을 말하지 않아도 트리거.
- **핵심 함정:** ① **재빌드가 손편집을 지운다** — 덱을 통째로 새로 쓰면 PowerPoint에서 직접
  넣은 노트·옮긴 상자·타이핑한 숫자가 사라진다. ② **발표자 노트가 조용히 실패한다** —
  템플릿 notes master에 body placeholder가 없으면 `notes_text_frame`이 `None`이다(템플릿이
  노트를 못 쓰는 게 아니라 placeholder를 주입하면 된다). ③ **넘침을 안 알려준다** — 이미지·
  텍스트가 슬라이드 경계를 넘어도 예외가 없다. ④ **autofit을 못 믿는다** — 렌더러마다 다르다.
- **정본 구현:** 공개 오픈소스 스킬
  **[`kangdacool/pptx-editing-skill`](https://github.com/kangdacool/pptx-editing-skill)** (MIT).
  규칙 + 실제로 돌아가는 도구 + 자기검사가 함께 있다. 작업 전 확보해 적용·검증한다.
