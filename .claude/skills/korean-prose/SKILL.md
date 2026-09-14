---
name: korean-prose
description: >-
  Audit and fix Korean (한국어) academic or report prose for the defects a machine
  or translation-minded writer makes repeatedly: josa after numbers (1.20로 → 1.20으로),
  「의」 chains, subject–predicate mismatch, translationese (상관된다 · 을 필요로 하다),
  double negation, redundancy (가장 높은 최고치), leftover «» marks. Use whenever you
  write, edit, or review Korean prose — 논문 · 원고 · 보고서 · 초록 · 요약 — in .md,
  .txt, .docx, or .hwpx, before handing it to a human, and when calibrating style
  against a target journal's published articles.
---

# Korean prose (pointer)

한국어 문장을 **쓰고 나서 반드시 돌리는** 감사. 맞춤법 검사기는 통과하는데 한국어 화자가 읽으면 걸리는 결함을 잡는다.

- **트리거:** 한국어 원고·보고서·초록을 쓰거나 고친 뒤, 사람에게 넘기기 전.
- **핵심 함정:**
  1. **숫자 뒤 조사는 읽는 소리로 정한다.** 「1.20」은 「영」으로 끝나므로 「1.20으로」, 「12.5%」는 「퍼센트」로 끝나므로 「12.5%였다」.
  2. **주어와 마지막 서술어만 이어 읽는다.** 「두 결과는 …의 평가이다」에서 결과는 평가가 아니다.
  3. **건수를 점수로 삼지 않는다.** 줄이려고 멀쩡한 문장을 비틀면 새 결함이 생긴다.
  4. **게재본에서 흔한 표현은 문체다.** 규칙을 넓히기 전에 투고할 학술지와 편집 수준이 높은 같은 분야 학술지의 게재본으로 잰다.
- **정본 구현:** 공개 오픈소스 스킬
  **[`kangdacool/korean-prose-skill`](https://github.com/kangdacool/korean-prose-skill)** (MIT).
  작업 전 확보해 적용·검증한다.
