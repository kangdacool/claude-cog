#!/usr/bin/env python3
"""
PreToolUse:Bash guard — refuse shell heredocs.

WHY. A heredoc (`python - <<'PY' ... PY`) passes its body through the tool call, the Bash
tool, the shell and finally the interpreter. On this setup something in that chain rewrites
escapes and non-ASCII: `\\n` arrives as a real newline, `\\b` as a backspace, `κ` as `?`,
en-dashes as hyphens. It fails SILENTLY -- the symptom is a syntax error on an unrelated
line, a regex that matches nothing while reporting success, or a manuscript sentence whose
Greek letters turned into Latin ones.

On 2026-08-15/16 this happened eight times in one session, and the rule against it was
already written down. A rule that is known and broken eight times is not a rule problem, so
this blocks instead.

WHAT TO DO INSTEAD. Write the script or the text to a file with the Write tool, then run
the file. One extra tool call; never corrupts.

    Write  scratchpad/fix.py   (the code)
    Bash   python scratchpad/fix.py

FAILS OPEN. Any unexpected input, parse error or exception allows the command through. A
guard that blocks work because it crashed is worse than the problem it prevents.

Disable: remove the PreToolUse block from ~/.claude/settings.json.
"""
import json
import re
import sys

# 이 훅의 메시지에는 한국어가 있고, 콘솔이 cp949면 깨져서 «읽을 수 없는 차단 사유»가 된다
# (2026-08-25 실측: 새로 넣은 메시지가 통째로 깨져 나왔다). 차단은 이유가 읽혀야 뜻이 있다.
for _s in (sys.stderr, sys.stdout):
    if hasattr(_s, "reconfigure"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# `<<` + optional dash + optional quote + a word, and then END OF LINE, because a heredoc's
# body starts on the next line. Requiring end-of-line is what separates a real heredoc from
# the same characters appearing inside a quoted argument: `grep -n "a<<b" file.txt` is not a
# heredoc and blocking it would be an over-reach. Excludes `<<<` (here-string: one line, safe)
# and shell redirections like `2<&1`. A trailing redirection after the delimiter is allowed.
# (?<!<) 가 없으면 `cmd <<< "x"` 의 2·3번째 `<`가 heredoc으로 잡힌다 -- 뒤의 `(?![<])`는
# 그 자리에서 통과하기 때문이다. here-string은 한 줄이라 안전하므로 막으면 안 된다.
# (2026-08-19 자기검사로 발견한 기존 오탐.)
HEREDOC = re.compile(
    r"(?<!<)<<-?[^\S\n]*(?![<])(['\"]?)([A-Za-z_][A-Za-z0-9_]*)\1[^\S\n]*(?:\d?[<>&|][^\n]*)?$",
    re.M)

# ── stdin으로 코드를 먹이는 «다른» 형태: `python -` ────────────────────────────
#
# WHY. heredoc과 같은 부류이고 결과는 더 나쁘다. 이 환경에는 stdin이 없으므로
# `python -`는 «영원히 매달린다» -- 셸이 반환하지 않고, python 프로세스가 CPU를 태우며
# 남는다. 2026-08-25 실측: 한 세션에서 5번 썼고, 8월 23일부터 셸 3개와 python 4개가
# 누적 250 CPU-시간을 태우고 있었다. 그리고 그 부하를 «다른 파이프라인의 이상»으로
# 오진해 남의 SuperLearner 작업(17시간)을 죽이는 2차 사고까지 냈다.
#
# 이것도 규칙으로는 이미 실패했다 -- 세션 기록에 "python - blocked on stdin THREE times,
# my own repeated mistake"라고 적힌 «뒤에» 두 번 더 했다. 그래서 막는다.
#
# ⚠️ **«명령 위치»에서만 잡는다.** 문자열 전체를 훑으면 따옴표 안의 리터럴까지 잡혀
#    정상 작업이 막힌다 -- 실제로 걸릴 것들:
#        git commit -m "...python - 로 stdin을 물렸다..."
#        sed -i 's|python -|x|' file
#    그래서 시작·`;`·`&&`·`||`·줄바꿈 뒤만 본다. 구분자에서 `|`는 «뺐다»: sed 치환
#    패턴의 구분자로 흔하고, 게다가 `cat x.py | python -`는 stdin이 진짜 있어 정상
#    작동하므로 막을 이유가 없다(그 경우도 `python x.py`로 쓰면 그만이다).
# ⚠️ `-m` `-c` `-X` `--version` 처럼 뒤에 글자가 붙는 것은 잡지 않는다(진짜 옵션이다).
STDIN_CODE = re.compile(r"(?:^|[;\n]|&&|\|\|)\s*(python3?|py)\s+-(?=\s|$)")

# ── `python -c "..."` 인자에 비ASCII(한글 등) — heredoc 과 같은 훼손, 다른 입구 ──────────
#
# WHY. 2026-09-30 한 코호트 결과 덱: `python -c` 로 빌더에 한 줄을 넣었는데 인자 속 한글이
# 인코딩을 거치며 깨져(«���� ������») 파일에 그대로 기록됐고, 같은 줄의 게이트 정규식은
# 아무것도 못 잡는 «죽은 검사»가 된 채 통과했다. 같은 세션에 python -c 의 한글 치환이
# «substring not found» 로 두 번 더 실패했다. 셸 인자는 이 기계에서 한글을 보존하지 않는다.
# ⚠ «-c 인자 안»만 본다 — 같은 명령의 다른 부분(`; grep "한글" f`)은 셸이 그대로 넘기므로 막지 않는다.
#    \uXXXX 이스케이프로 쓴 한글은 ASCII 라 통과한다(그 방법은 이 날 제대로 동작했다).
PY_C_ARG = re.compile(r"(?:^|[;\n]|&&|\|\|)\s*(?:python3?|py)\s+-c\s+(['\"])")


def py_c_nonascii(cmd):
    """python -c 의 코드 인자에 비ASCII 문자가 있으면 그 앞뒤를 돌려준다."""
    for m in PY_C_ARG.finditer(cmd):
        q, i = m.group(1), m.end()
        j = i
        while j < len(cmd):
            if cmd[j] == "\\" and q == '"':
                j += 2
                continue
            if cmd[j] == q:
                break
            j += 1
        arg = cmd[i:j]
        for k, ch in enumerate(arg):
            if ord(ch) > 127:
                return arg[max(0, k - 25):k + 25]
    return None


def backtick_in_double_quotes(cmd):
    """큰따옴표 안의 «이스케이프되지 않은» 백틱을 찾는다. 있으면 그 앞뒤 30자를 돌려준다.

    WHY. 백틱은 큰따옴표 안에서 명령치환이다. 마크다운 코드표기(`data/`)나 한국어 산문이
    든 문자열을 `python -c "..."`나 `git commit -m "..."`으로 넘기면 셸이 그 부분을
    실행하려 든다. 2026-08-19 한 세션에서 세 번 당했다 -- 커밋 메시지의 `dist/`가
    "command not found"가 되고, 로그에 적으려던 `profile_columns.py`가 통째로 사라져
    문장에 구멍이 뚫린 채 파일에 기록됐다. heredoc과 같은 부류다: 셸을 거치며 텍스트가
    조용히 훼손된다.

    작은따옴표 안의 백틱은 리터럴이라 안전하므로 보지 않는다."""
    i, n = 0, len(cmd)
    in_single = in_double = False
    while i < n:
        ch = cmd[i]
        if ch == "\\":
            i += 2
            continue
        if ch == "'" and not in_double:
            in_single = not in_single
        elif ch == '"' and not in_single:
            in_double = not in_double
        elif ch == "`" and in_double:
            return cmd[max(0, i - 30):i + 30]
        i += 1
    return None


def _run(payload):
    """검사 본체. dispatch.py 와 main() 이 공유한다 — stdin 읽기는 밖에서 한다."""
    try:
        if payload.get("tool_name") != "Bash":
            return 0
        cmd = payload.get("tool_input", {}).get("command", "") or ""

        snippet = backtick_in_double_quotes(cmd)
        if snippet:
            sys.stderr.write(
                "BLOCKED: 큰따옴표 안에 백틱이 있습니다 -- 셸이 그 부분을 실행합니다.\n\n"
                "  ...%s...\n\n"
                "백틱은 큰따옴표 안에서 명령치환입니다. 마크다운 코드표기나 산문이 든\n"
                "문자열을 python -c \"...\" 나 git commit -m \"...\" 로 넘기면 그 구절이\n"
                "사라지거나 'command not found'가 됩니다. 조용히 실패하므로, 파일에는\n"
                "구멍 뚫린 문장이 그대로 기록됩니다.\n\n"
                "이렇게 하십시오:\n"
                "  - 파일 내용을 쓸 때: Write 또는 Edit 도구를 쓴다(셸을 거치지 않는다).\n"
                "  - 커밋 메시지: -F 로 메시지 파일을 넘긴다.\n"
                "  - 정말 셸에 넘겨야 하면: 작은따옴표로 감싸거나 \\` 로 이스케이프한다.\n"
                % snippet.replace("\n", " "))
            return 2

        snip = py_c_nonascii(cmd)
        if snip:
            sys.stderr.write(
                "BLOCKED: python -c 인자에 비ASCII 문자(한글 등)가 있습니다.\n\n"
                "  ...%s...\n\n"
                "셸 인자를 거치며 한글이 깨져 파일에 그대로 기록됩니다(2026-09-30 실측:\n"
                "빌더의 게이트 한 줄이 깨진 채 저장돼 «아무것도 못 잡는 검사»가 됐다).\n\n"
                "이렇게 하십시오:\n"
                "  1. 스크립트를 Write 도구로 파일에 쓴다(한글이 UTF-8 로 그대로 간다).\n"
                "  2. python scratchpad/fix.py 로 실행한다.\n"
                "  (꼭 한 줄이어야 하면 한글을 \\uXXXX 이스케이프로 쓴다.)\n"
                % snip.replace("\n", " "))
            return 2

        m = STDIN_CODE.search(cmd)
        if m:
            sys.stderr.write(
                "BLOCKED: `%s -` reads code from stdin, and this environment has none.\n\n"
                "The command will HANG FOREVER: the shell never returns and the python\n"
                "process spins, burning CPU until something kills it. 2026-08-25 실측 --\n"
                "한 세션에서 5번, 셸 3개와 python 4개가 누적 250 CPU-시간을 태우고 있었다.\n"
                "그리고 그 부하를 다른 파이프라인의 이상으로 오진해 남의 계산 17시간을\n"
                "죽이는 2차 사고까지 났다.\n\n"
                "Do this instead:\n"
                "  1. Write the script to a file with the Write tool.\n"
                "  2. Run that file:  python scratchpad/fix.py\n\n"
                "One-liner면 `python -c \"...\"` 를 쓰십시오(그건 막지 않습니다).\n"
                % m.group(1))
            return 2

        m = HEREDOC.search(cmd)
        if not m:
            return 0
        sys.stderr.write(
            "BLOCKED: this command uses a shell heredoc (<<%s).\n\n"
            "Heredoc bodies get their escapes and non-ASCII rewritten somewhere between the\n"
            "tool call and the interpreter on this machine -- \\n becomes a real newline, \\b a\n"
            "backspace, and Greek letters are replaced. It fails silently, so the damage shows\n"
            "up later as a syntax error, a regex that matches nothing, or corrupted prose.\n\n"
            "Do this instead:\n"
            "  1. Write the script (or text) to a file with the Write tool.\n"
            "  2. Run that file with Bash.\n\n"
            "For a genuinely one-line command, use ordinary quoting -- no heredoc needed.\n"
            % m.group(2))
        return 2                                   # non-zero blocks and shows stderr
    except Exception:
        return 0                                   # fail open


def check(payload):
    """(종료코드, 메시지). dispatch.py 전용 — stderr 를 «가로채» 돌려준다."""
    import contextlib
    import io
    buf = io.StringIO()
    try:
        with contextlib.redirect_stderr(buf):
            code = _run(payload)
    except Exception:
        return 0, ""                               # fail open
    return code, buf.getvalue()


def main():
    try:
        payload = (json.loads(sys.stdin.buffer.read().decode("utf-8", errors="replace")) if hasattr(sys.stdin, "buffer") else json.load(sys.stdin))  # Windows stdin 기본 cp949 — 한글 payload 가 깨졌다(2026-09-28)
    except Exception:
        return 0                                   # fail open
    return _run(payload)


if __name__ == "__main__":
    sys.exit(main())
