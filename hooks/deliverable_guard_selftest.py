#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deliverable_guard 셀프테스트.

가드의 진짜 실패 모드는 «못 잡는 것»이 아니라 «엉뚱한 데서 우는 것»이다 -- 우는 가드는
꺼지고, 꺼진 가드는 없는 가드다. 그래서 절반은 «울리면 안 되는» 경우다.

    python deliverable_guard_selftest.py
"""
import atexit
import os
import shutil
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


ran = []


def check(label, got, want):
    # ⚠ 개수는 «세어서» 찍는다. 예전엔 요약줄에 「15 cases」가 문장으로 박혀 있었고,
    #   사례를 9개 더한 순간 조용히 거짓이 됐다(2026-09-11). [[core]] §3 — 행수·개수를
    #   문장에 박지 말 것. 박아야 하면 「문서가 말한 숫자 vs 실제」를 대조할 것.
    ran.append(label)
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
# ⚠ fixture 를 OS 임시 폴더에 만들면 «새 규칙이 그걸 정확히 걸러낸다»(2026-09-11).
#   규칙이 맞으므로 규칙을 풀지 말고 시험 자리를 옮긴다 — 임시 루트 «밖»에 만든다.
tmp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "._dg_fixture")
shutil.rmtree(tmp, ignore_errors=True)
os.makedirs(tmp, exist_ok=True)
atexit.register(lambda: shutil.rmtree(tmp, ignore_errors=True))
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

# ---- 2b. «청중에게 가는 자리»인가 (2026-09-11 오발에서 나옴) -----------------
# 실제 오발: 스크래치패드의 되돌림용 사본 `before_strip.hwpx` 를 감사하고 검증 명령을
# «막았다». 거기서 찾은 결함은 내가 이미 제거한 학회 공고문의 따옴표 — 산출물에는 없는
# 결함으로 작업을 세운 것이다. 확장자도 신선도도 둘 다 «참»이었으므로, 이 축이 없으면 못 막는다.
for label, p, want in [
    ("스크래치패드 사본",
     r"C:\Users\U\AppData\Local\Temp\claude\s\scratchpad\before_strip.hwpx", False),  # noleak
    ("_work 되돌림본", r"D:\proj\_work\before_strip.hwpx", False),  # noleak
    ("_archive", r"D:\proj\_archive\old.docx", False),  # noleak
    ("백업 확장자", r"D:\proj\manuscript\draft.bak.docx", False),  # noleak
    ("사본 표기", r"D:\proj\manuscript\draft-백업.docx", False),  # noleak
    ("워드 잠금파일", r"D:\proj\manuscript\~$draft.docx", False),  # noleak
    ("진짜 산출물", r"D:\proj\manuscript\draft_260911.docx", True),  # noleak
    ("진짜 투고본", r"D:\proj\submission\main.hwpx", True),  # noleak
    ("이름에 temp 가 든 «정상» 폴더", r"D:\proj\template\deck.pptx", True),  # noleak
    ("원본 표기", r"D:\proj\manuscript\draft_원본.docx", False),  # noleak
]:
    check("자리 판정 — " + label, G.is_deliverable_path(p), want)

# 임시 자리는 «이름»이 아니라 구조로 본다 — 폴더를 뭐라 부르든 OS 임시 루트 아래면 임시다.
# (첫 판은 `[Tt]e?mp` 정규식이었다. 그러면 임시 루트가 다른 이름일 때 뚫린다.)
_tmproot = tempfile.gettempdir()
check("자리 판정 — OS 임시 루트 아래(이름 무관)",
      G.is_deliverable_path(os.path.join(_tmproot, "aZq9", "deck.pptx")), False)
check("자리 판정 — 임시 루트 밖의 같은 이름",
      G.is_deliverable_path(os.path.join("D:", "proj", "aZq9", "deck.pptx")), True)

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
print("%d cases (그중 «자리 판정» %d건 — 2026-09-11 오발에서 나옴)"
      % (len(ran), sum(1 for x in ran if x.startswith("자리 판정"))))
if fails:
    print("  FAIL %d건" % len(fails))
    for f in fails:
        print("    - %s" % f)
    sys.exit(1)
print("  OK  all pass")
