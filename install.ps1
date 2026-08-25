<#
  install.ps1 -- 이 저장소를 이 기계의 Claude Code 에 «물린다».

      pwsh -File install.ps1            무엇이 바뀌는지 보여주고 묻는다
      pwsh -File install.ps1 -Force     묻지 않는다
      pwsh -File install.ps1 -WhatIf    아무것도 안 바꾸고 계획만 인쇄

  왜 필요한가
      배선이 절대경로 9곳에 박혀 있었다 -- CLAUDE.md 스텁 1 · 정션 2 · 훅 6.
      OneDrive 가 다른 드라이브에 있거나 계정명이 다르면 전부 깨지는데,
      **에러가 아니라 「훅이 안 도는 상태」로 조용히 시작된다.** 그게 최악이다.
      이 스크립트는 자기 위치에서 경로를 «계산»해서 건다.

  무엇을 바꾸나 (셋 다 되돌릴 수 있게 백업한다)
      ~/.claude/CLAUDE.md        1줄 스텁 -> 이 저장소의 claude-config/CLAUDE.md
      ~/.claude/skills, agents   이 저장소로 가는 정션
      ~/.claude/settings.json    hooks.manifest.json 의 훅을 «병합»(다른 설정 보존)

  무엇을 «안» 바꾸나
      .credentials.json · 모델·테마 등 개인 설정 · 기존 훅 중 이 매니페스트에 없는 것.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param([switch]$Force)

$ErrorActionPreference = "Stop"
$ROOT = $PSScriptRoot

# 레이아웃 둘을 다 안다 -- 파일을 두 벌로 만들지 않으려고.
#   정본 저장소: claude-config/{hooks,skills,agents} + claude-config/CLAUDE.md
#   공개 킷    : hooks/ + .claude/{skills,agents} + CLAUDE.md.template
if (Test-Path (Join-Path $ROOT "claude-config\hooks")) {
    $CFG     = Join-Path $ROOT "claude-config"
    $HOOKDIR = Join-Path $CFG "hooks"
    $SKILLS  = Join-Path $CFG "skills"
    $AGENTS  = Join-Path $CFG "agents"
    $MD      = Join-Path $CFG "CLAUDE.md"
} else {
    $CFG     = $ROOT
    $HOOKDIR = Join-Path $ROOT "hooks"
    $SKILLS  = Join-Path $ROOT ".claude\skills"
    $AGENTS  = Join-Path $ROOT ".claude\agents"
    $MD      = Join-Path $ROOT "CLAUDE.md"
}
$CLAUDE = Join-Path $HOME ".claude"
$STAMP    = Get-Date -Format "yyMMdd_HHmmss"
$changes  = @()

function Say($m)  { Write-Host "  $m" }
function Head($m) { Write-Host ""; Write-Host "== $m" -ForegroundColor Cyan }

# ── 0. 전제 확인 ───────────────────────────────────────────────────────────
Head "전제"
foreach ($p in @($CFG, $HOOKDIR, $SKILLS, $AGENTS)) {
    if (-not (Test-Path $p)) { throw "없습니다: $p  (저장소 루트에서 실행하십시오)" }
}
$manifestPath = Join-Path $HOOKDIR "hooks.manifest.json"
if (-not (Test-Path $manifestPath)) { throw "없습니다: $manifestPath" }
$manifest = (Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json).hooks
Say "저장소: $ROOT"
Say "훅 매니페스트: $($manifest.Count)개"
if (-not (Test-Path $CLAUDE)) { New-Item -ItemType Directory $CLAUDE | Out-Null }

foreach ($exe in @("python", "node")) {
    $found = (Get-Command $exe -ErrorAction SilentlyContinue)
    if ($found) { Say "$exe : $($found.Source)" }
    else { Write-Warning "$exe 를 찾을 수 없습니다 -- 그 훅은 조용히 안 돕니다." }
}

# ── 1. CLAUDE.md 스텁 ──────────────────────────────────────────────────────
Head "CLAUDE.md 스텁"
$stubPath = Join-Path $CLAUDE "CLAUDE.md"
$canon    = $MD -replace '\\', '/'
$stub     = "@$canon"
$cur      = if (Test-Path $stubPath) { (Get-Content $stubPath -Raw -Encoding UTF8).Trim() } else { $null }
if ($cur -eq $stub) {
    Say "이미 맞음"
} else {
    if ($cur) { Say "현재: $($cur -split "`n" | Select-Object -First 1)" }
    Say "바꿈: $stub"
    $changes += @{ what = "CLAUDE.md 스텁"; act = {
        if (Test-Path $stubPath) { Copy-Item $stubPath "$stubPath.bak_$STAMP" }
        Set-Content -Path $stubPath -Value $stub -Encoding UTF8 -NoNewline
    } }
}

# ── 2. 정션 (skills, agents) ───────────────────────────────────────────────
Head "정션"
foreach ($name in @("skills", "agents")) {
    $link   = Join-Path $CLAUDE $name
    $target = if ($name -eq "skills") { $SKILLS } else { $AGENTS }
    $item   = Get-Item $link -ErrorAction SilentlyContinue
    # ⚠️ PS7 에서 정션의 .Target 은 «문자열»이다 -- $item.Target[0] 은 첫 글자("D")를 집는다.
    #    @(...) 로 감싸야 원소가 된다. 이걸 틀리면 멀쩡한 정션을 매번 다시 건다.
    $curTarget = if ($item -and $item.LinkType) { @($item.Target)[0] } else { $null }
    if ($curTarget -and
        (Resolve-Path $curTarget -ErrorAction SilentlyContinue).Path -eq (Resolve-Path $target).Path) {
        Say "$name : 이미 맞음"
        continue
    }
    if ($item -and -not $item.LinkType) {
        Say "$name : 실제 폴더가 있습니다 -> _backup_$STAMP 으로 «옮깁니다»(지우지 않습니다)"
    } elseif ($item) {
        Say "$name : 다른 곳을 가리킵니다 -> 다시 겁니다"
    } else {
        Say "$name : 새로 겁니다"
    }
    # ⚠️ .GetNewClosure() 가 «없으면» 이 블록은 실행 시점의 $link/$target/$name -- 즉
    #    루프 마지막 값(agents) -- 을 본다. 스킬 정션이 조용히 안 걸리고 로그만
    #    「적용」이라 찍힌다(2026-08-25 가짜 HOME 시험에서 실제로 그랬다).
    #    R 의 for 루프 클로저와 같은 함정이다.
    $changes += @{ what = "정션 $name"; act = {
        $it = Get-Item $link -ErrorAction SilentlyContinue
        if ($it -and -not $it.LinkType) {
            $bk = Join-Path $CLAUDE "_backup_$STAMP"
            if (-not (Test-Path $bk)) { New-Item -ItemType Directory $bk | Out-Null }
            Move-Item $link (Join-Path $bk $name)
        } elseif ($it) {
            (Get-Item $link).Delete()          # 링크만 지운다(대상은 안전)
        }
        New-Item -ItemType Junction -Path $link -Target $target | Out-Null
    }.GetNewClosure() }
}

# ── 3. 훅 병합 ─────────────────────────────────────────────────────────────
Head "훅"
$setPath = Join-Path $CLAUDE "settings.json"
$set = if (Test-Path $setPath) { Get-Content $setPath -Raw -Encoding UTF8 | ConvertFrom-Json }
       else { [pscustomobject]@{} }
$existing = @()
if ($set.PSObject.Properties.Name -contains "hooks") {
    foreach ($ev in $set.hooks.PSObject.Properties.Name) {
        foreach ($e in $set.hooks.$ev) { foreach ($h in $e.hooks) { $existing += $h.command } }
    }
}
$todo = @()
foreach ($m in $manifest) {
    $script = Join-Path $HOOKDIR $m.script
    if (-not (Test-Path $script)) { Write-Warning "훅 파일 없음, 건너뜀: $($m.script)"; continue }
    $cmd = '{0} "{1}"' -f $m.runner, ($script -replace '\\', '/')
    if ($existing -contains $cmd) { Say "이미 걸림: $($m.script)"; continue }
    $stale = $existing | Where-Object { $_ -like "*$($m.script)*" }
    if ($stale) { Say "경로가 낡음, 다시 검: $($m.script)" }
    else        { Say "새로 검: $($m.event)/$($m.matcher) $($m.script)" }
    $todo += @{ m = $m; cmd = $cmd; stale = $stale }
}
if ($todo.Count -gt 0) {
    $changes += @{ what = "훅 $($todo.Count)개"; act = {
        Copy-Item $setPath "$setPath.bak_$STAMP" -ErrorAction SilentlyContinue
        $s = if (Test-Path $setPath) { Get-Content $setPath -Raw -Encoding UTF8 | ConvertFrom-Json }
             else { [pscustomobject]@{} }
        if ($s.PSObject.Properties.Name -notcontains "hooks") {
            $s | Add-Member -NotePropertyName hooks -NotePropertyValue ([pscustomobject]@{})
        }
        foreach ($t in $todo) {
            $ev = $t.m.event
            # 낡은 경로의 같은 훅을 먼저 걷어낸다
            if ($s.hooks.PSObject.Properties.Name -contains $ev) {
                foreach ($entry in $s.hooks.$ev) {
                    $entry.hooks = @($entry.hooks | Where-Object { $_.command -notlike "*$($t.m.script)*" })
                }
                $s.hooks.$ev = @($s.hooks.$ev | Where-Object { $_.hooks.Count -gt 0 })
            } else {
                $s.hooks | Add-Member -NotePropertyName $ev -NotePropertyValue @()
            }
            $entry = [pscustomobject]@{ hooks = @([pscustomobject]@{ type = "command"; command = $t.cmd }) }
            if ($null -ne $t.m.matcher) {
                $entry | Add-Member -NotePropertyName matcher -NotePropertyValue $t.m.matcher
            }
            $s.hooks.$ev = @($s.hooks.$ev) + $entry
        }
        $s | ConvertTo-Json -Depth 12 | Set-Content $setPath -Encoding UTF8
    } }
}

# ── 3b. git pre-commit (개인식별정보가 이력에 들어가기 «전에» 막는다) ──────
# 한 번 커밋되면 이력에서 지우기가 매우 어렵다 -- 들어가기 전이 유일하게 싼 지점이다.
# 실사고 2026-08-25: 환자 성명이 프로젝트 메모리에 6일 있었다.
Head "git pre-commit"
$ghDir = Join-Path $ROOT "claude-config\githooks"
if (-not (Test-Path $ghDir)) { $ghDir = Join-Path $ROOT "githooks" }
if (-not (Test-Path (Join-Path $ghDir "pre-commit"))) {
    Say "pre-commit 파일이 없습니다 -- 건너뜁니다"
} elseif (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Say "git 을 찾을 수 없습니다 -- 건너뜁니다"
} else {
    $want = ($ghDir -replace '\\', '/')
    $curHP = (git config --global --get core.hooksPath) 2>$null
    if ($curHP -eq $want) {
        Say "이미 걸림 (core.hooksPath)"
    } elseif ([string]::IsNullOrWhiteSpace($curHP)) {
        Say "전역 core.hooksPath 를 이 저장소로 겁니다: $want"
        Say "  (해제: git config --global --unset core.hooksPath)"
        $changes += @{ what = "git pre-commit"; act = {
            git config --global core.hooksPath $want
        }.GetNewClosure() }
    } else {
        # ⚠️ 남의 설정을 덮지 않는다. 이미 쓰는 훅이 있으면 그쪽이 우선이다.
        Write-Warning ("core.hooksPath 가 이미 다른 곳을 가리킵니다: $curHP`n" +
                       "  덮지 않았습니다. 두 훅을 합치려면 그 폴더의 pre-commit 에서 " +
                       "$want/pre-commit 을 호출하십시오.")
    }
}

# ── 4. 적용 ────────────────────────────────────────────────────────────────
Head "요약"
if ($changes.Count -eq 0) { Say "바꿀 것이 없습니다 -- 이미 물려 있습니다."; }
else {
    $changes | ForEach-Object { Say "· $($_.what)" }
    Say "백업 접미사: .bak_$STAMP"
    $go = $Force -or $PSCmdlet.ShouldProcess("이 기계의 ~/.claude", "위 항목 적용")
    if (-not $Force -and -not $WhatIfPreference) {
        $go = (Read-Host "적용할까요? (y/N)") -eq "y"
    }
    if ($go -and -not $WhatIfPreference) {
        foreach ($c in $changes) { & $c.act; Say "적용: $($c.what)" }
    } else { Say "적용하지 않았습니다."; return }
}

# ── 5. 검증 -- 걸렸다고 «말하지» 말고 돌려 본다 ────────────────────────────
Head "검증"
$fail = 0
foreach ($n in @("skills", "agents")) {
    $ok = Test-Path (Join-Path (Join-Path $CLAUDE $n) "*")
    Say ("{0} 정션: {1}" -f $n, $(if ($ok) { "OK" } else { $fail++; "실패" }))
}
$stubOk = (Test-Path $stubPath) -and ((Get-Content $stubPath -Raw -Encoding UTF8).Trim() -eq $stub)
Say ("CLAUDE.md 스텁: {0}" -f $(if ($stubOk) { "OK" } else { $fail++; "실패" }))

foreach ($t in @("hooks/heredoc_guard_selftest.py",
                 "hooks/docx_kit_guard_selftest.py",
                 "skills/docx-editing/scripts/docx_kit_selftest.py",
                 # 개인식별정보 게이트 -- 회귀 사례가 실제로 샜던 문장이다
                 "../agent/tools/precommit_scan_selftest.py",
                 "../tools/precommit_scan_selftest.py")) {
    $p = Join-Path $CFG $t
    if (-not (Test-Path $p)) { continue }
    & python $p *> $null
    Say ("{0}: {1}" -f (Split-Path $t -Leaf), $(if ($LASTEXITCODE -eq 0) { "통과" } else { $fail++; "실패" }))
}

Write-Host ""
if ($fail -gt 0) { Write-Host "★ $fail 건 실패" -ForegroundColor Red; exit 1 }
Write-Host "물렸습니다. Claude Code 를 다시 시작하십시오(훅·스킬은 시작 시 읽습니다)." -ForegroundColor Green
