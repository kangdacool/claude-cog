# CLAUDE.md — operating-guide adapter (lean, portable)

**정본(spine)은 `docs/claude-operating-guide.md`.** 이 파일은 상시 로딩되므로 *항상 지켜야 할 규칙*만 요약한다. 방법론·함정·QA 루프의 전문은 코어 가이드를 읽을 것. 세부 절차는 여기 넣지 않는다.

## 항상 지키는 규칙 (요약 — 근거는 코어)

- **납품은 편집 가능한 파일만.** `.docx`/`.pptx`/`.xlsx`가 결과물. **PDF는 QA용 중간물**이지 산출물이 아니다(명시 요청 시에만 PDF).
- **렌더 후 눈으로 확인.** 코드 성공 ≠ 결과물 정확. PDF→PNG를 **네이티브 해상도**로 뽑아 표 페이지까지 전부 보고, **폰트가 작지 않은지**·0바이트 PNG 여부를 점검. 가능하면 새 눈(서브에이전트).
- **상호참조 양방향 대조.** §참조·"슬라이드 N"·덱↔해설서 참조를 추출해 정규식으로 양방향 검증. 번호 재배치는 단일 패스 원자적 remap + 의미상 stale 참조 감사.
- **수치는 1차 출처.** 모든 수치를 원문과 대조·기준시점 명시, 추정치엔 'E'. 여러 파일에 흩어진 핵심 수치는 납품 전 grep으로 파일 간 값 일치 확인.
- **참고문헌 2-pass.** Pass 1 실존(PMID/DOI 기록) + Pass 2 인용 주장이 실제 뒷받침되는지 확인. **레퍼런스 날조 금지** — 후보를 PMID/DOI와 함께 제시.
- **표는 booktabs만**(가로줄만), 유의성은 색이 아니라 **굵게**. 장식·신호색·색 바 금지("AI가 만든 티").
- **출력 표면 규율.** 제목·캡션·콜아웃에 편집 지시·선정 기준·AI 흔적·자명한 설명을 넣지 않는다. 편집 사유는 채팅으로만. (코어 §6)
- **고비용 변경은 합의 후.** 구조 전복·절 추가/삭제·표 열 추가는 옵션 제시 → 확인 후 실행. 사소한 편집은 바로.

## 메모리 (경로 규칙)

- **영구 메모리의 정본 위치를 하나의 동기화 폴더로 고정**한다(예: 클라우드 드라이브의 `claude-kb/`). 인덱스 파일 하나 + 주제별 노트(`feedback/`, `reference/`, `projects/<name>/`) 구조를 권장.
- 하네스 기본 per-project 메모리 경로에는 **저장하지 않는다** — 위 정본 위치가 하네스 기본 동작을 **명시적으로 대체(override)** 하도록 이 규칙을 둔다. (경로는 각자 환경에 맞게 지정.)

## 스킬·에이전트 (트리거 요약)

- **cross-verify** — 문서 투고/납품 전 원고 수치를 원본 표와 대조.
- **download-refs** — 원고 참고문헌 PDF 일괄 다운로드(기관 브라우저 접근 필요; 구현은 각자 환경).

**문서 스킬 넷** — 여기 있는 `SKILL.md`는 «안내판»이고 실제로 도는 코드는 각 공개 저장소에 있다(전부 MIT, 규칙 + 도구 + 자기검사 + CI). 파일을 만지기 **전에** 확보한다.

| 대상 | 스킬 | 정본 |
|---|---|---|
| 한글 `.hwpx` | **hwpx-editing** (+ `hwpx` 에이전트) | [`kangdacool/hwpx-editing-skill`](https://github.com/kangdacool/hwpx-editing-skill) |
| `.pptx` | **pptx-editing** | [`kangdacool/pptx-editing-skill`](https://github.com/kangdacool/pptx-editing-skill) |
| `.docx` | **docx-editing** | [`kangdacool/docx-editing-skill`](https://github.com/kangdacool/docx-editing-skill) |
| 한국어 문장 | **korean-prose** | [`kangdacool/korean-prose-skill`](https://github.com/kangdacool/korean-prose-skill) |

## 도메인 레이어 (직접 채워 넣기)

특정 분야의 코딩 관례·문서 서식·통계/보고 규칙 등 **도메인 특화 규칙은 이 포터블 코어에 넣지 않는다.** 자신의 도메인 레이어(예: 별도 전역 `CLAUDE.md`)에 두고 여기서 가리킬 것. 코어는 채팅 어시스턴트에도 첨부할 수 있도록 일반 방법론만 유지한다.
