#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PostToolUse 훅 — 방금 쓴 «산출물»을 그 자리에서 감사한다. 장르는 audit.py가 정한다.

    등록: settings.json 의 PostToolUse (matcher: Bash|Write)

왜 있나
    도구를 만드는 것과 그 도구가 «제때 불리는 것»은 다른 문제다. 이 시스템에는 감사 17개와
    그것을 장르별로 라우팅하는 audit.py가 있었지만, **audit.py를 부르는 것이 없었다** —
    즉 전부 «기억해야 도는» 검사였다. 규칙을 더 크게 쓰는 대신 «순간»에 묶는다.
    (reference/tool_taxonomy.md 의 승격 규칙: 중요한데 자꾸 안 돌아가는 check는 guard로.)

    선례는 hwpx_guard.py다. 그 훅의 표현 그대로: "에이전트는 검증을 «돌리기로 선택해야»
    돌린다. 이 훅은 그 선택을 없앤다." 이 파일은 같은 일을 .hwpx 하나가 아니라 **모든
    산출물 형식**에 대해 하고, 무엇을 돌릴지는 스스로 정하지 않고 audit.py에 묻는다.
    ⇒ 검사 목록의 정본은 여전히 REGISTRY 한 곳이다. 훅은 «순간»만 제공한다.

무엇을 검사하지 않는가
    · 오래된 파일 — mtime이 FRESH_SECONDS 밖이면 이번 명령이 만든 게 아니다.
    · COM/렌더가 필요한 검사 — `--fast`가 뺀다. 저장할 때마다 한글·PowerPoint를 띄우면
      훅이 느려서 결국 «끄게» 되고, 꺼진 가드는 없는 가드다. 뺀 것은 [SKIP]으로 이름이
      남으므로 «안 돈 것»과 «통과한 것»이 구별된다.
    · audit.py 자신을 부르는 명령 — 무한 반복 방지.
    · 중간 산출물 — 장르가 안 잡히면(audit.py가 rc 0으로 조용히 지나감) 아무 말 안 한다.

조용한 것이 기본이다. 통과하면 아무것도 출력하지 않는다.
"""
import json
import os
import re
import subprocess
import sys
import time

# 한국어 사유를 stderr로 낸다. Windows cp949 콘솔에서 그대로 쓰면 UnicodeEncodeError로
# 훅이 죽고, **죽은 훅은 «조용히 통과»와 구별되지 않는다**. hwpx_guard와 같은 가드를 둔다.
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError, OSError):
    pass

FRESH_SECONDS = 180        # 이 명령이 만졌다고 볼 시간 창
MAX_FILES = 3              # 한 번에 감사할 파일 수 상한 (훅 지연 방지)
PER_FILE_TIMEOUT = 90      # audit.py는 여러 검사를 도므로 hwpx_guard보다 넉넉히

# ⚠️ 하드코딩하면 다른 기계에서 «에러 없이 아무것도 안 하는» 훅이 된다.
#    레이아웃이 둘이다: 정본은 <repo>/claude-config/hooks/ + <repo>/agent/tools/,
#    공개 킷은 <repo>/hooks/ + <repo>/tools/. 둘 다에서 찾는다.
_HERE = os.path.dirname(os.path.abspath(__file__))
AUDIT = ""
for _cand in (
    os.path.join(os.path.dirname(os.path.dirname(_HERE)), "agent", "tools", "audit.py"),
    os.path.join(os.path.dirname(_HERE), "tools", "audit.py"),
):
    if os.path.exists(_cand):
        AUDIT = _cand
        break
else:                                   # 못 찾아도 죽지 않는다 -- 훅은 조용히 통과한다
    AUDIT = os.path.join(os.path.dirname(_HERE), "tools", "audit.py")

# ⭐ «조판된 산출물»만. 청중에게 가는 물건이라는 뜻이다.
#
# .md/.tex를 넣었다가 뺐다(2026-08-20 실측). 메모리 코퍼스 61개 중 35개가 울렸는데, 그 지적이
# «틀려서»가 아니라 «내부 노트에는 해당되지 않아서»다 — 색인 파일의 대시 통일이나 ↑·≠ 사용은
# 결함이 아니다. 판정 기준은 참/거짓이 아니라 **청중에게 가는 물건인가**이다.
# 그 상태로 두면 .md를 고칠 때마다 우는 훅이 되고, 우는 훅은 하루 만에 꺼진다.
# 원고 소스(.md/.tex)는 .docx로 빌드될 때 이 훅을 다시 지나가므로 «놓치는» 것도 아니다.
# .xlsx/.csv도 제외 — 산출물이 아니라 «입력 자료원»이다(audit.py의 data 장르).
# 필요하면 손으로: python <이 저장소>/tools/audit.py <파일>
EXTS = (".pptx", ".docx", ".hwpx")

QUOTED = [re.compile(r'"([^"]+?\.(?:pptx|docx|hwpx))"'),
          re.compile(r"'([^']+?\.(?:pptx|docx|hwpx))'")]
BARE = re.compile(r'(?:^|[\s=])((?:[A-Za-z]:)?[^\s"\';|&><]+\.(?:pptx|docx|hwpx))')


def candidates(command):
    """명령문에서 산출물 경로 후보를 뽑는다. 따옴표 경로를 먼저 -- 공백 경로가 잘리지 않게."""
    found, seen = [], set()
    for rx in QUOTED:
        for m in rx.finditer(command):
            if m.group(1) not in seen:
                seen.add(m.group(1)); found.append(m.group(1))
    stripped = command
    for rx in QUOTED:
        stripped = rx.sub(" ", stripped)
    for m in BARE.finditer(stripped):
        if m.group(1) not in seen:
            seen.add(m.group(1)); found.append(m.group(1))
    return found


def fresh_targets(raws, cwd):
    now, targets = time.time(), []
    for raw in raws:
        path = raw if os.path.isabs(raw) else os.path.join(cwd, raw)
        try:
            if not os.path.isfile(path):
                continue
            if now - os.path.getmtime(path) > FRESH_SECONDS:
                continue                            # 이번 명령이 만든 게 아니다
        except OSError:
            continue
        targets.append(path)
        if len(targets) >= MAX_FILES:
            break
    return targets


def main():
    try:
        event = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        return 0                                    # 훅 입력이 이상하면 조용히 통과

    tool = event.get("tool_name")
    ti = event.get("tool_input") or {}
    cwd = event.get("cwd") or os.getcwd()

    if tool == "Bash":
        command = ti.get("command") or ""
        # audit.py를 부르는 명령에서 다시 audit을 돌리면 무한 반복이 된다.
        if not command or "audit.py" in command or "deliverable_guard" in command:
            return 0
        raws = candidates(command)
    elif tool in ("Write", "Edit"):
        fp = ti.get("file_path") or ""
        raws = [fp] if fp.lower().endswith(EXTS) else []
    else:
        return 0

    targets = fresh_targets(raws, cwd)
    if not targets or not os.path.exists(AUDIT):
        return 0

    problems = []
    for path in targets:
        try:
            r = subprocess.run([sys.executable, AUDIT, "--fast", path],
                               capture_output=True, text=True,
                               encoding="utf-8", errors="replace",
                               timeout=PER_FILE_TIMEOUT)
        except (subprocess.TimeoutExpired, OSError):
            continue                                # 감사기가 못 돌면 통과시킨다
        if r.returncode == 0:
            continue
        fails = [ln.strip() for ln in (r.stdout or "").splitlines()
                 if ln.lstrip().startswith("[FAIL]") or ln.lstrip().startswith("실패:")]
        detail = [ln.rstrip() for ln in (r.stdout or "").splitlines()
                  if ln.startswith("        ")][:6]
        if fails:
            problems.append((path, fails, detail))

    if not problems:
        return 0

    print("산출물 감사 실패 — 방금 쓴 파일이 그 장르의 검사를 통과하지 못했습니다. "
          "이 상태로 전달하면 사용자가 받는 파일에 그대로 남습니다.", file=sys.stderr)
    for path, fails, detail in problems:
        print("\n  %s" % path, file=sys.stderr)
        for ln in fails[:5]:
            print("    %s" % ln, file=sys.stderr)
        for ln in detail:
            print("    %s" % ln.strip()[:150], file=sys.stderr)
    # 경로를 «해결된 것»으로 찍는다 -- 레이아웃이 둘이라 문자열로 박으면 한쪽에서 틀린다.
    print("\n  전체 출력(느린 검사 포함): python %s <파일>" % AUDIT, file=sys.stderr)
    return 2                                        # stderr가 Claude에게 전달된다


if __name__ == "__main__":
    sys.exit(main())
