<div align="center">

# 🔬 Claude Operating Guide — 연구실 운영 킷

**한 연구실이 실제로 쓰는 Claude Code 세팅을 통째로.**
방법론 문서 + 훅 + 서브에이전트 + 감사 도구 + 설치 스크립트.

_The working Claude Code setup of a research lab: the operating guide, the hooks that
enforce it, the audit tools that check the output, and a one-command installer._

[![CI](https://github.com/kangdacool/claude-cog/actions/workflows/ci.yml/badge.svg)](https://github.com/kangdacool/claude-cog/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)
![Works with](https://img.shields.io/badge/agents-Claude%20Code%20%C2%B7%20Codex%20%C2%B7%20Cursor%20%C2%B7%20Gemini-orange)

</div>

---

## 이게 무엇인가

논문·보고서·발표덱을 만드는 연구실에서, 같은 실수가 반복되지 않도록 **규칙을 코드로
바꿔 온 기록**입니다. 산문 가이드 하나가 아니라 네 층입니다.

| 층 | 무엇 | 왜 |
|---|---|---|
| **가이드** | `docs/claude-operating-guide.md` (270줄) | 방법론의 «왜». 채팅 Claude 에 이 파일 하나만 첨부해도 동작합니다 |
| **규약** | `CLAUDE.md.template` (618줄) | 상시 로딩되는 운영 규약. IMRaD 규율·참고문헌 2-pass·파이프라인 구조·통계 보고 서식 |
| **훅** | `hooks/` (6개) | 규칙이 «지켜지게» 만드는 장치. 산출물이 나오면 감사를 자동으로 돌리고, 조용히 망가지는 명령을 막습니다 |
| **도구** | `tools/` (41개) | 감사 19 · 규율 8 · 제작 5 · R 5 · 유틸 4 + 자기검사 |

**한 줄 요약: 자주 일어나고 조용히 틀리는 것은 규칙이 아니라 검사로 막습니다.**
규칙으로 다섯 번 실패한 명령 형태가 훅을 걸고 나서야 멈췄습니다 — 그 경험이 이 저장소의 뼈대입니다.

## 설치

```powershell
git clone https://github.com/kangdacool/claude-cog
cd claude-cog
pwsh -File install.ps1 -WhatIf     # 무엇이 바뀌는지만 본다
pwsh -File install.ps1             # 묻고 적용
```

`~/.claude` 의 셋만 건드리고 **전부 백업**합니다 — `CLAUDE.md` 스텁 · `skills`/`agents` 정션 ·
`settings.json` 의 훅(**병합**이라 기존 설정과 남의 훅이 보존됩니다). 멱등이라 다시 돌려도
「바꿀 것이 없습니다」로 끝납니다. 자세한 것은 [`SETUP.md`](SETUP.md).

> Windows 전용입니다(정션 + PowerShell). 다른 OS라면 `hooks/hooks.manifest.json` 을 보고
> 심볼릭 링크로 같은 것을 걸면 됩니다 — 도구 자체는 OS를 가리지 않습니다.

**채팅 Claude 에서 쓰려면** `docs/claude-operating-guide.md` 한 파일만 첨부하고
*"이 가이드의 규칙을 따라 작업해줘"* 라고 하십시오. 그게 전부입니다.

## 무엇이 들어 있나

```
docs/claude-operating-guide.md   방법론 정본 — 이 한 파일이 «왜»의 전부
CLAUDE.md                        Claude Code 용 lean 어댑터(상시 규칙 요약)
CLAUDE.md.template               [채우기] 자리를 채워 자기 연구실 규약으로
hooks/                           훅 6개 + 매니페스트 + 자기검사
tools/                           감사·규율·제작 도구 41개
.claude/skills/                  워크플로 포인터 (아래 «문서 포맷 스킬»)
.claude/agents/                  서브에이전트 5 — 원고감사·파이프라인감사·문헌조사 등
install.ps1  SETUP.md            설치와 되돌리기
```

### 훅 — 규칙이 지켜지게 만드는 층

| 훅 | 언제 | 무엇을 막나 |
|---|---|---|
| `deliverable_guard.py` | 산출물(.docx/.pptx/.hwpx)이 만들어질 때 | 조판 결함이 그대로 나가는 것. 장르별 감사를 자동으로 돌립니다 |
| `heredoc_guard.py` | Bash 실행 전 | 조용히 망가지거나 **셸이 영원히 멈추는** 명령 형태 |
| `docx_kit_guard.py` | 문서 조립 코드를 쓸 때 | 이미 있는 도구를 못 찾고 다시 짜는 것 |
| `log-*.js` (3) | 프롬프트·계획·응답 종료 | 무엇을 시켰고 어떤 모델이 답했는지 세션 로그로 |

**막지 않고 알리는 훅이 있습니다.** 막으면 그때 훅을 끄게 되고, 꺼진 가드는 없는 가드입니다.

### 도구 — 산출물을 «사람 눈으로 보기 전에» 거르는 층

단일 진입점 `tools/audit.py` 가 파일 장르(docx·pptx·hwpx·md·tex)로 검사를 라우팅하고,
**건너뛴 검사와 그 이유까지** 보고합니다.

```bash
python tools/audit.py 원고.docx        # 장르로 라우팅
python tools/audit.py --list           # 장르 x 검사 전체 지도
```

글자 크기 · 칸 폭 · **인용되지 않은 표·그림** · 본문↔표 수치 일치 · 용어 충돌 ·
낡은 파생 산출물 · 목표 저널 문체 적합 · 표면 유출(편집 흔적이 산출물에 새는 것).

그 밖에 `memory_health.py`(교훈 코퍼스가 부풀지 않게) · `precommit_scan.py`(커밋 전
개인식별정보 스캔) · `build_guard.py`(손편집을 재생성이 덮어쓰지 않게) ·
`journal_display_census.py`(목표 지면의 **실제** 표·그림 규범을 세어 봅니다).

**자기검사가 11개** 있고, 그중 절반은 «울리지 않아야 하는 경우»를 증명합니다 —
오탐이 많은 검사는 꺼지고, 꺼진 검사는 없는 검사이기 때문입니다.

## 문서 스킬 — 코드는 따로 있습니다

`.claude/skills/*/SKILL.md` 는 «언제 무엇을 여는가»를 알려주는 안내판이고, 실제로 도는
코드는 아래 네 저장소입니다(전부 MIT, 규칙 + 도구 + 자기검사 + CI).

| 포맷 | 저장소 | 무엇을 막아 주나 |
|---|---|---|
| 한글 `.hwpx` | **[`hwpx-editing-skill`](https://github.com/kangdacool/hwpx-editing-skill)** | 재압축·`linesegarray` 함정으로 한글이 파일을 못 여는 것 |
| `.pptx` | **[`pptx-editing-skill`](https://github.com/kangdacool/pptx-editing-skill)** | 재빌드가 손편집·발표자 노트를 지우는 것, 경계 넘침 |
| `.docx` | **[`docx-editing-skill`](https://github.com/kangdacool/docx-editing-skill)** | 줄마다 문단, 저널 표 세로줄, 렌더가 조용히 죽는 것 |
| 한국어 문장 (`.md` `.txt` `.docx` `.hwpx`) | **[`korean-prose-skill`](https://github.com/kangdacool/korean-prose-skill)** | 숫자 뒤 조사, 「의」 연쇄, 주술 불일치, 번역투 — 맞춤법 검사기는 통과하는 결함 |

⚠️ **새 포맷 스킬을 공개하면 이 표와 `CLAUDE.md` 의 스킬 절에 함께 추가하십시오.**
안 하면 이 저장소가 «현관»인데 문패가 하나 빠진 상태가 됩니다.

## 무엇이 «없나» — 일부러

- **교훈 원문과 프로젝트 메모리.** 사람·기관·미발표 결과가 들어 있습니다.
  `docs/claude-operating-guide.md` 가 그것의 익명화 증류본이고, 그게 옳은 구조입니다.
- **기관 접근·개인 계정에 매인 도구**(브라우저 자동화·메일 발송·참고문헌 PDF 수집).
- **특정 자료를 전제하는 분석 도구.** 남의 환경에서 안 돕니다.

## 도메인 레이어는 직접

분야별 코딩 관례·문서 서식·통계 보고 규칙은 이 포터블 코어에 넣지 않습니다.
자기 도메인 레이어(별도 전역 `CLAUDE.md`)에 두고 여기서 가리키십시오.

## 라이선스

MIT — [LICENSE](LICENSE). 출처·범위는 [NOTES.md](NOTES.md).
