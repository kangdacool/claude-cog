#!/usr/bin/env python3
"""Self-test for heredoc_guard.py. Half the cases exist to prove it does NOT block."""
import json
import subprocess
import sys
import os

GUARD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "heredoc_guard.py")

BLOCK = [
    "python - <<'PY'\nprint(1)\nPY",
    'python - <<"PY"\nprint(1)\nPY',
    "cat > f.txt <<EOF\nhello\nEOF",
    "python3 - <<- EOF\nx=1\nEOF",
    "cd /tmp && python - <<'PYEOF'\nimport io\nPYEOF",
    # 큰따옴표 안의 백틱 = 명령치환. 마크다운 코드표기가 든 산문을 셸로 넘기면 사라진다.
    'python -c "s=s.replace(\'a\',\'`data/` 설명\')"',
    'git commit -m "fix `dist/` path"',
    # `python -` = stdin으로 코드 먹이기. 이 환경엔 stdin이 없어 «영원히 매달린다».
    # 아래 다섯은 2026-08-25에 실제로 매달렸던 명령의 «형태»다(경로만 줄였다).
    "python -",
    "python - 2>/dev/null; sed -i 's|a|b|' f",
    'cd /d && python - "MEMORY.md" 2>/dev/null; echo x',
    "cd /d && sed -i 's|a|b|' f; python - \"noop\" 2>/dev/null; echo x",
    "python3 -",
    # python -c 인자 속 한글 — 셸 인자를 거치며 깨진다(2026-09-30 실측)
    "python -c \"s=s.replace('a','내부 변수명')\"",
    "cd /d/x && python -c 'print(\"한글\")'",
]
ALLOW = [
    "ls -la",
    "python scratchpad/fix.py",
    'grep -n "a<<b" file.txt',            # literal text, not a heredoc
    "python -c \"print('hi')\"",
    "echo $((1 << 3))",                   # arithmetic shift
    'cat <<< "here string"',              # here-string, single line, safe
    'python -c "print(1)" <<< "x"',       # here-string with a quoted word (옛 오탐)
    "python -c 's = \"`literal`\"'",      # 작은따옴표 안의 백틱은 리터럴
    'git commit -m "fix \\`dist\\` path"',  # 이스케이프된 백틱
    "git commit -m 'x' && git push",
    "",                                   # empty
    # ── `python -` 가드가 «막으면 안 되는» 것들 ────────────────────────────
    # 진짜 옵션: 뒤에 글자가 붙는다.
    "python -m pip install x",
    "python --version",
    "python -X faulthandler s.py",
    "py -3 s.py",
    # 리터럴로 등장하는 것. 이 문자열을 막으면 이 가드를 «설명하는 커밋»조차 못 한다.
    "git commit -m 'python - 로 stdin을 물렸다'",
    "sed -i 's|python -|python3 |' f.py",
    'grep -rn "python -" hooks/',
    # 파이프로 stdin이 «진짜» 있는 경우: 정상 작동하므로 막지 않는다.
    "cat s.py | python -",
    # python -c 는 ASCII 면 통과. 같은 명령의 «다른 부분»에 있는 한글은 셸이 그대로 넘긴다.
    "python -c \"print('\\uccab')\"",
    'python -c "print(1)"; grep -n "한글" f.txt',
]


def run(cmd, tool="Bash"):
    # ⚠️ 가드가 stderr를 utf-8로 고정하므로 여기서도 utf-8로 읽는다. text=True만 쓰면
    #    로케일(cp949)로 디코드하려다 UnicodeDecodeError가 난다 -- 반환코드 검사는
    #    통과하는데 시험이 죽어서, 「33/33 통과」와 예외가 같이 찍힌다(2026-08-25).
    p = subprocess.run([sys.executable, GUARD],
                       input=json.dumps({"tool_name": tool,
                                         "tool_input": {"command": cmd}}),
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return p.returncode


fail = 0
for c in BLOCK:
    if run(c) != 2:
        print("MISS (should block):", c.split("\n")[0]); fail += 1
for c in ALLOW:
    if run(c) != 0:
        print("FALSE POSITIVE:", c.split("\n")[0]); fail += 1
# a non-Bash tool must always pass, and malformed input must fail open
if run("python - <<'PY'\nx\nPY", tool="Write") != 0:
    print("MISS: non-Bash tool was blocked"); fail += 1
p = subprocess.run([sys.executable, GUARD], input="not json", capture_output=True, text=True)
if p.returncode != 0:
    print("MISS: malformed input did not fail open"); fail += 1

total = len(BLOCK) + len(ALLOW) + 2
print("\n%d/%d cases correct" % (total - fail, total))
sys.exit(1 if fail else 0)
