# -*- coding: utf-8 -*-
"""reject_ledger.py 자기시험.

⚠ CORE §3: 「게이트도 시험한다 — 통과할 것·실패할 것을 하나씩 넣어 본다.」
   안 하면 「항상 통과하는」 또는 「항상 실패하는」 게이트를 쓰게 된다.
   그래서 여기서 「거부되어야 할 입력」과 「통과해야 할 입력」을 둘 다 넣는다.

Usage: python reject_ledger_selftest.py     (exit 0 = 통과)
"""
import io
import os
import sys
import tempfile
import contextlib

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import reject_ledger as RL  # noqa: E402

FAILS = []


def check(cond, msg):
    if not cond:
        FAILS.append(msg)
        print("  FAIL  %s" % msg)
    else:
        print("  ok    %s" % msg)


def run(argv):
    """main()을 돌리고 (exit_code, stdout) 반환."""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            rc = RL.main(argv)
    except SystemExit as e:
        return e.code, buf.getvalue()
    return rc, buf.getvalue()


def main():
    tmp = tempfile.mkdtemp()
    RL.LEDGER = os.path.join(tmp, "ledger.jsonl")
    P = ["--project", "T"]

    print("1) 세 줄이 빠지면 「거부」되어야 한다 (실패해야 할 입력)")
    rc, out = run(["add"] + P + ["--candidate", "c0", "--evidence", "verified"])
    check(isinstance(rc, str) and "기각이 성립하지 않는다" in rc,
          "criterion·known-good 누락 -> SystemExit + 이유 설명")
    check(not os.path.exists(RL.LEDGER) or "c0" not in open(RL.LEDGER, encoding="utf-8").read(),
          "거부된 입력은 「기록되지 않는다」")

    print("2) 세 줄이 다 있으면 통과해야 한다 (통과해야 할 입력)")
    ok = ["--criterion", "기준A", "--known-good", "선례에 대면 통과", "--evidence", "verified"]
    rc, out = run(["add"] + P + ["--candidate", "c1"] + ok)
    check(rc == 0, "완전한 입력 -> 기록됨")
    check("연속 1회" in out, "연속 계수 1")

    print("3) 연속 3회에서 「경고」가 나와야 한다 (임계 시험)")
    run(["add"] + P + ["--candidate", "c2"] + ok)
    rc, out = run(["add"] + P + ["--candidate", "c3"] + ok)
    check("연속 3회" in out, "연속 계수 3")
    check("기준을 감사" in out and "known-good" in out, "임계에서 「기준 감사」 경고 출력")
    rc, out = run(["status"] + P)
    check(rc == 1, "status는 임계 초과 시 exit 1")

    print("4) 「생존」이 계수를 되돌려야 한다")
    run(["survive"] + P + ["--candidate", "winner"])
    rc, out = run(["status"] + P)
    check(rc == 0, "생존 기록 후 status exit 0")
    check("연속 기각 0" in out, "연속 계수가 0으로 초기화")

    print("5) 근거가 「추측」이면 표시되어야 한다")
    run(["add"] + P + ["--candidate", "c4", "--criterion", "기준B",
                       "--known-good", "미대조", "--evidence", "guess"])
    rc, out = run(["status"] + P)
    check("추측" in out and "c4" in out, "추측 기반 기각을 재검토 후보로 표시")

    print("6) 다른 프로젝트와 「섞이지」 않아야 한다")
    run(["add", "--project", "OTHER", "--candidate", "x"] + ok)
    rc, out = run(["status"] + P)
    check("연속 기각 1" in out, "프로젝트별로 계수가 분리된다")

    print()
    if FAILS:
        print("FAILED %d" % len(FAILS))
        return 1
    print("all passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
