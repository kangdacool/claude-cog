#!/usr/bin/env python3
"""Selftest for journal_fit.strip_working_text.

Why this exists (2026-08-24, 한 원고): the tool counted Korean working notes carried in
`<!-- ... -->` and reported em-dash at 29.7/10k, "ABOVE every paper", when the manuscript body
held exactly one em-dash and that one sat inside a comment. A style measurement taken on text
the author would never submit is worse than no measurement: it sends you rewriting prose that
is already fine.

Run:  python journal_fit_selftest.py       (exit 1 on any failure)
"""
import importlib.util
import os
import sys

_spec = importlib.util.spec_from_file_location(
    "journal_fit", os.path.join(os.path.dirname(os.path.abspath(__file__)), "journal_fit.py"))
JF = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(JF)

FAIL = []


def check(name, cond, detail=""):
    if cond:
        print("  ok   %s" % name)
    else:
        FAIL.append(name)
        print("  FAIL %s  %s" % (name, detail))


# 1. HTML comments go, including multi-line ones carrying the very characters we measure.
src = ("Teachers were exposed more often.\n"
       "<!-- 작업주석 — 이 줄은 투고본에 안 들어간다;\n"
       "     여러 줄이고 줄표도 있다 — -->\n"
       "The gap widened.\n")
clean, dropped = JF.strip_working_text(src)
check("HTML comment removed", "작업주석" not in clean, repr(clean))
check("em-dash inside comment not counted", "—" not in clean, repr(clean))
check("prose kept", "The gap widened." in clean and "Teachers were exposed" in clean)
check("dropped fraction reported", 0.3 < dropped < 0.9, "dropped=%.2f" % dropped)

# 2. A CJK line with plenty of Latin is still a working note (the 0.15-ratio version missed these).
src2 = ("Results are reported per point.\n"
        "⚠ 2-pass 검증 완료 7편: [1](ILO PDF) · [3][4][6][10](PMC) — 확인함\n"
        "The association held.\n")
clean2, _ = JF.strip_working_text(src2)
check("Latin-heavy CJK note removed", "PMC" not in clean2, repr(clean2))
check("surrounding prose kept", "The association held." in clean2)

# 3. Clean English is untouched — the filter must not eat real prose.
src3 = "We fitted a model; the estimate was 1.14. It held after adjustment.\n"
clean3, dropped3 = JF.strip_working_text(src3)
check("clean English untouched", clean3.strip() == src3.strip(), repr(clean3))
check("nothing dropped from clean text", dropped3 < 0.01, "dropped=%.3f" % dropped3)

print("\n%s" % ("ALL PASS" if not FAIL else "FAILED: " + ", ".join(FAIL)))
sys.exit(1 if FAIL else 0)
