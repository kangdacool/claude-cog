# SETUP — 이 킷을 기계에 물리기

```powershell
git clone https://github.com/kangdacool/claude-cog
cd claude-cog
pwsh -File install.ps1 -WhatIf     # 무엇이 바뀌는지만 본다
pwsh -File install.ps1             # 묻고 적용 (-Force 면 안 묻는다)
```

경로는 아무 데나 좋습니다 — **설치기가 자기 위치에서 계산**합니다.

## 무엇을 바꾸나 (셋 다 백업합니다: `.bak_<날짜시각>`)

| | |
|---|---|
| `~/.claude/CLAUDE.md` | 이 저장소의 `CLAUDE.md` 를 `@import` 하는 **1줄 스텁** |
| `~/.claude/skills`, `~/.claude/agents` | 이 저장소로 가는 **정션** |
| `~/.claude/settings.json` | `hooks/hooks.manifest.json` 의 훅 6개를 **병합** |

**안 건드리는 것**: `.credentials.json`, 모델·테마 등 개인 설정, 매니페스트에 없는 기존 훅.
이미 물려 있으면 「바꿀 것이 없습니다」로 끝납니다(멱등).

끝나면 **Claude Code 를 다시 시작**하십시오 — 훅과 스킬은 시작할 때 읽습니다.

## 확인은 스크립트가 «돌려서» 합니다

설치기가 마지막에 정션·스텁을 다시 읽고 자기검사를 실제로 실행합니다. 하나라도 어긋나면
exit 1 입니다. **「걸었다」고 찍는 것과 「걸렸다」는 것은 다릅니다** — 개발 중 실제로
`적용: 정션 skills` 를 찍고도 안 걸린 판이 있었습니다(PowerShell 루프 클로저 함정).

## 쓰기 시작하기

1. **`CLAUDE.md.template` 을 자기 규약으로.** `[채우기]` 표시가 붙은 곳 — 메모리 위치와
   프로젝트 매핑표 — 만 바꾸면 나머지(IMRaD 규율·참고문헌 2-pass·파이프라인 구조·통계
   보고 서식)는 그대로 돌아갑니다. 다 채웠으면 `~/.claude/CLAUDE.md` 스텁이 그 파일을
   가리키게 하십시오.
2. **메모리 폴더를 만듭니다.** 동기화되는 한 곳에 `MEMORY.md`(인덱스) + `feedback/` +
   `projects/<name>/`. 하네스 기본 per-project 경로에는 두지 마십시오 — 기기마다 갈라집니다.
3. **산출물을 만들면 감사를 돌립니다.** `python tools/audit.py <파일>`
   (`deliverable_guard` 훅이 자동으로도 돌립니다).

## 되돌리기

`~/.claude` 의 `*.bak_<날짜시각>` 를 되돌리고 정션 둘을 지웁니다.

```powershell
foreach ($n in "skills","agents") { $p = "$HOME\.claude\$n"; if (Test-Path $p) { (Get-Item $p).Delete() } }
```

⚠️ **정션은 `Remove-Item -Recurse` 로 지우지 마십시오** — 링크를 따라가 **원본을 지웁니다.**
`.Delete()` 는 링크만 지웁니다.

## 다른 OS

Windows 전용입니다(정션 + PowerShell). 다른 OS라면 `hooks/hooks.manifest.json` 을 읽어
심볼릭 링크와 `settings.json` 항목을 같은 모양으로 만들면 됩니다 — **도구 자체는 OS를
가리지 않습니다**(CI 가 Linux 에서 자기검사 전수를 돌립니다).
