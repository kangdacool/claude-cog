---
name: docx-editing
description: >-
  Build, edit, and QA Word (.docx) documents with python-docx so they follow
  academic typesetting conventions instead of the defaults. Use this whenever a
  task involves a .docx file, Word / 워드 / 원고 / manuscript / 보고서 / 브리프 /
  초록, or writing a builder that emits .docx — journal tables, markdown that
  becomes real paragraphs, figure captions, and a render-and-look QA step.
  Trigger even if the user only says "원고 만들어줘" / ".docx" — the default
  python-docx path produces one paragraph per line, literal ** markers, and
  gridded tables that violate journal convention.
---

# DOCX editing (pointer)

Word `.docx` 를 **학술 조판 관습대로** 만드는 워크플로.

- **트리거:** `.docx` / 원고 / 보고서 / 브리프 작업, 또는 `.docx`를 내는 빌더를 짤 때.
- **핵심 함정:** ① **마크다운을 그대로 넣으면 문서가 세로로 늘어진다** — 하드랩(`\n`)이 문단
  경계로 오해되어 3,800단어가 413문단이 된다(빈 줄만 문단 경계다). ② **`**굵게**`가 글자로
  남는다.** ③ **저널 표에 세로줄이 그어진다** — `table.style = 'Table Grid'`가 주범이고,
  출판된 저널 표는 세로줄이 하나도 없다. ④ **docx→PDF 렌더가 원인 없이 죽는다** — Word COM을
  맨몸 `Dispatch`로 잡으면 «이미 떠 있는» 인스턴스에 붙고, 그게 멈춰 있으면 파일이 깨진 것처럼
  보인다(`DispatchEx`가 답). ⑤ **장르를 틀리면 표 전체를 다시 짠다** — 저널 원고의 표와
  기관 브리프의 표는 다른 물건이다.
- **⭐ 부품이 아니라 «골격»부터.** 표지·쪽나눔·표번호·캡션·게이트를 매번 처음부터 엮으면
  그러다 부품까지 다시 짜게 된다 — 실측으로 확인된 실패 양상이다(코어 §6).
- **정본 구현:** 공개 오픈소스 스킬
  **[`kangdacool/docx-editing-skill`](https://github.com/kangdacool/docx-editing-skill)** (MIT).
  작업 전 확보해 적용·검증한다.
