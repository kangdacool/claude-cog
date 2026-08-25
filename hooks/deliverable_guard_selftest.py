#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deliverable_guard 셀프테스트.

가드의 진짜 실패 모드는 «못 잡는 것»이 아니라 «엉뚱한 데서 우는 것»이다 -- 우는 가드는
꺼지고, 꺼진 가드는 없는 가드다. 그래서 절반은 «울리면 안 되는» 경우다.

    python deliverable_guard_selftest.py
"""
import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError, OSError):
    pass

import deliverable_guard as G

fails = []


def check(label, got, want):
    if got != want:
        fails.append("%s\n      기대 %r\n      실제 %r" % (label, want, got))


# ---- 1. 경로 추출 --------------------------------------------------------
check("따옴표 안 공백 경로", G.candidates('python build.py "out/최종 보고서.hwpx"'),
      ["out/최종 보고서.hwpx"])
check("맨경로", G.candidates("python make_deck.py output/deck.pptx"), ["output/deck.pptx"])
check("여러 개", G.candidates("cp a.docx b.docx"), ["a.docx", "b.docx"])
check("절대경로(윈도우)", G.candidates(r"python x.py D:\out\report.docx"), [r"D:\out\report.docx"])  # noleak: 시험용 가짜 경로

# 울리면 안 되는 것: «청중에게 가는 물건»이 아니다
check("xlsx는 산출물 아님(입력 자료원)", G.candidates("python load.py data/raw.xlsx"), [])
check("py는 산출물 아님", G.candidates("python run.py script.py"), [])
check("md는 원고 «소스»지 산출물이 아님", G.candidates("python build.py draft.md"), [])
check("확장자 없음", G.candidates("ls output/"), [])

# ---- 2. 신선도(mtime) ----------------------------------------------------
tmp = tempfile.mkdtemp(prefix="dg_selftest_")
fresh = os.path.join(tmp, "fresh.docx")
stale = os.path.join(tmp, "stale.docx")
for p in (fresh, stale):
    with open(p, "w", encoding="utf-8") as f:
        f.write("x\n")
old = time.time() - (G.FRESH_SECONDS + 60)
os.utime(stale, (old, old))

check("방금 쓴 파일은 대상", G.fresh_targets([fresh], tmp), [fresh])
check("오래된 파일은 대상 아님", G.fresh_targets([stale], tmp), [])
check("없는 파일은 대상 아님", G.fresh_targets([os.path.join(tmp, "nope.docx")], tmp), [])

# 상한: MAX_FILES를 넘으면 자른다 (훅 지연 방지)
many = []
for i in range(G.MAX_FILES + 2):
    p = os.path.join(tmp, "m%d.docx" % i)
    with open(p, "w", encoding="utf-8") as f:
        f.write("x\n")
    many.append(p)
check("파일 수 상한", len(G.fresh_targets(many, tmp)), G.MAX_FILES)

# ---- 3. 무한 반복 방지 ---------------------------------------------------
# audit.py를 부르는 명령에서 다시 audit을 돌리면 훅이 자기를 부른다.
import io
import json


def run_event(payload):
    """main()을 실제 stdin/stderr로 돌려 반환코드를 본다."""
    old_in, old_err = sys.stdin, sys.stderr
    sys.stdin = io.StringIO(json.dumps(payload))
    sys.stderr = io.StringIO()
    try:
        return G.main()
    finally:
        sys.stdin, sys.stderr = old_in, old_err


check("audit.py 호출 자체는 통과",
      run_event({"tool_name": "Bash", "cwd": tmp,
                 "command": "x", "tool_input": {"command": "python audit.py %s" % fresh}}), 0)
check("Bash 아닌 도구는 통과",
      run_event({"tool_name": "Read", "cwd": tmp, "tool_input": {"file_path": fresh}}), 0)
# 깨진 stdin -- 훅이 죽으면 «조용히 통과»와 구별되지 않으므로 반드시 rc 0으로 살아나와야 한다
_oi, _oe = sys.stdin, sys.stderr
sys.stdin, sys.stderr = io.StringIO("not json"), io.StringIO()
try:
    check("JSON 아닌 stdin은 조용히 통과", G.main(), 0)
finally:
    sys.stdin, sys.stderr = _oi, _oe

# ---- 4. 등록 정합 --------------------------------------------------------
check("audit.py 경로 존재", os.path.exists(G.AUDIT), True)

# ---- 결과 ---------------------------------------------------------------
print("15 cases (8건은 «울리면 안 되는» 경우)")
if fails:
    print("  FAIL %d건" % len(fails))
    for f in fails:
        print("    - %s" % f)
    sys.exit(1)
print("  OK  all pass")
