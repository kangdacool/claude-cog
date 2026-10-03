#!/usr/bin/env python3
"""
manuscript_register.py — detect CONFESSIONAL register and generated-prose rhythm in a
manuscript body, without touching the academic hedging that must survive.

WHY THIS EXISTS (2026-08-15). A methods manuscript went through a full
number check (162/162 against source CSVs), a cross-reference check, and an independent
reviewer-eye audit. All three passed the prose. The user then asked one question — "did
you use words like 'honesty' and 'we fixed it' where a plain statement would do?" — and
the answer was yes, six times, in the paper's most sensitive paragraphs:

    "and doing so misled us once"
    "One caution emerged in obtaining these: a first version ..."
    "The first version of this comparison said the opposite"
    "That this is not a universal small-sample floor was not assumed but checked"

Every one of those was TRUE and every one was a liability. The rule they violate:

    ⭐ In a methods paper, honesty belongs in the DESIGN — pre-stated gates, failures
    reported in tables, limitations named. In the PROSE it should read as fact, not as
    confession. "The paired comparison isolates the factor" is the same disclosure as
    "comparing the other way misled us once", minus the defensiveness a referee hears.

The failure mode is specific and predictable: it appears when the SESSION has spent
effort catching its own errors. That vigilance leaks into the manuscript's voice. A
number checker cannot see it — every sentence is true.

WHAT THIS DELIBERATELY DOES NOT FLAG
    The correction is one-directional and easy to overdo. These are ACADEMIC register
    and must survive untouched; none is matched here:
      * stated limitations   ("the scope boundary is bracketed rather than located")
      * pre-registration     ("the rule was fixed before the confirming experiment")
      * precision contrasts  ("conditional rather than marginal", "a bound, not a
                              threshold") — antithesis is reported as a COUNT, not as a
                              hit list, because most instances are load-bearing
      * negative results     ("six remedies failed"), uncertainty, hedged magnitude
    Only first-person process narration and self-correction anecdote are called hits.

Usage:
    python manuscript_register.py FILE [--start "## Introduction"] [--end "## Display items"]
    # exits 1 if any CONFESSION hits; density metrics are reported either way

As a library:
    from manuscript_register import scan
    report = scan(text)      # -> dict with 'confession', 'process', 'density'
"""
import argparse
import io
import os
import re
import sys

# Windows 기본 콘솔은 cp949 라 ⚠·«» 에서 UnicodeEncodeError 로 «검사기가 죽는다».
# 그러면 audit.py 가 [FAIL] 로 집계해 산문에 문제가 없는데 있는 것처럼 보인다.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def read_any(path):
    """본문 문자열. .md/.txt/.tex 는 그대로, .docx/.hwpx 는 추출해서.

    ⚠ 추출기 사본을 만들지 않는다 — audit_text_consistency 의 것을 재사용한다.
      2026-09-05 이전에는 이 파일이 평문만 읽었다. 그런데 audit.py REGISTRY 에는
      manuscript/supplement/brief(전부 .docx 장르)로 «등록»돼 있었다 — 즉 돌 수 없는
      곳에 등록만 돼 있었고, auto=False 라 아무도 그 사실을 몰랐다.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".docx", ".hwpx"):
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from audit_text_consistency import read_any as _read
        body, cells = _read(path)
        return "\n".join(list(body) + list(cells))
    return io.open(path, encoding="utf-8").read()

# --- hits: reported line by line, and an exit-1 condition ---------------------------
CONFESSION = [
    (r"misled us|led us astray|we were wrong|we had mis-?specified",
     "self-correction anecdote"),
    (r"(a |the )?first (version|attempt|pass)\b(?=[^.]*\b(said|gave|reported|showed)\b)",
     "earlier-attempt narration"),
    (r"one caution emerged|a caution emerged|it turned out that we",
     "editorialised aside"),
    (r"was not assumed but|rather than assumed|we did not simply assume",
     "asserting one's own rigour instead of reporting the check"),
    (r"\bhonest(ly)?\b|\bto be fair\b|\bin fairness\b|\bwe must admit\b|\bfrankly\b",
     "claiming honesty rather than being plain"),
    (r"worth (stating|reporting|saying)|is worth noting|plainly put|it should be said",
     "rhetorical framing"),
    (r"the (test|thing|part) that (matters|counts)|the real question is",
     "rhetorical emphasis"),
]
PROCESS = [
    # "fixed" and "corrected" are the trap. "We fixed the acceptance rule before the
    # confirming run" is PRE-REGISTRATION and must survive; "we fixed the bug" is
    # confession. The first version of this pattern matched bare `fixed` and flagged the
    # pre-registration sentence in the very manuscript it was written for -- the exact
    # over-correction this file warns against. Match the object, not the verb.
    (r"\b(we|us|our)\b(?=[^.]*\b(misled|realis|realiz|discovered|noticed|caught|"
     r"traced it|had to)\b)", "first-person process narration"),
    (r"\bwe\b[^.]{0,40}\b(fixed|corrected|repaired)\b\s+(it|them|this|that|the\s+"
     r"(bug|error|mistake|code|script|problem|issue))\b", "first-person repair narration"),
]

# --- density: reported as counts only; NOT failures ---------------------------------
DENSITY = {
    "em-dash": r"—",
    "antithesis (rather than / X, not Y)": r"rather than|,\s*not\s+\w|\bnot\s+\w+\s+but\b",
    "numeral-opener sentence": r"(?m)^\s*(One|Two|Three|Four|Five|Six|Seven|Eight|Nine|Ten)\b",
    "cleft opener (What X is/does)": r"(?m)^\s*What\b[^.]{0,60}\b(is|does|costs)\b",
}


def scan(text):
    lines = text.split("\n")
    out = {"confession": [], "process": [], "density": {}, "words": len(text.split())}
    for bucket, rules in (("confession", CONFESSION), ("process", PROCESS)):
        for i, ln in enumerate(lines, 1):
            for pat, why in rules:
                if re.search(pat, ln, re.I):
                    out[bucket].append((i, why, ln.strip()))
                    break
    for label, pat in DENSITY.items():
        n = len(re.findall(pat, text, re.I))
        per = out["words"] // n if n else 0
        out["density"][label] = (n, per)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--start", default=None, help="only scan from this marker onward")
    ap.add_argument("--end", default=None, help="stop at this marker")
    a = ap.parse_args()

    t = read_any(a.file)
    span = "--start/--end \uc9c0\uc815"
    if a.start and a.start in t:
        t = t[t.index(a.start):]
    elif not a.start and not a.end:
        ## ⛔ 2026-09-08 — 여기가 «본문»을 문서 전체로 세고 있었다. 실측: 같은 원고를
        ##    journal_fit 은 7,090단어, 이 도구는 9,786단어로 읽었다(참고문헌·표셀 포함).
        ##    분모가 38% 부풀면 AI 티 밀도가 실제보다 «좋아 보인다» — 재는 도구가
        ##    스스로 축소 보고한 것이다.
        ##    ⚠ 구간 검출을 «복사하지 않고» journal_fit 것을 재사용한다(read_any 와 같은
        ##       패턴). 복사하면 두 도구가 갈라지고, 갈라지면 어느 쪽이 맞는지 아무도 모른다.
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from journal_fit import auto_span
        t, span = auto_span(t)
    if a.end and a.end in t:
        t = t[:t.index(a.end)]
    r = scan(t)

    print("body: %d words   [span: %s]\n" % (r["words"], span))
    for label, (n, per) in r["density"].items():
        note = "  <- clusters read as generated prose" if per and per < 130 else ""
        print("  %-38s %3d%s%s" % (label, n,
                                   ("  (one per %d words)" % per) if per else "", note))
    ## ⚠ 2026-09-08 — 위 130은 «절대 바닥»이지 «저널 기준»이 아니다. 실측:
    ##    IJEqH OA 12편의 «X, not Y 최대치»가 3.8/10k = one per 2,632 words 였다 —
    ##    이 문턴이 20배 느슨하다. 그래서 이 검사가 조용해도 그 저널에서는 이상치일 수 있다.
    ##    ⛔ 130을 감으로 고치지 않는다 — «너무 많다»는 저널마다 다르기 때문이다.
    print("  ⚠ 위 기준(one per 130)은 절대 바닥이다. «목표 저널 대비» 판정은")
    print("    journal_fit.py --corpus 가 한다 — 그쪽이 조용하기 전까지 이 검사의 침묵은 보증이 아니다.")
    print()
    bad = r["confession"] + r["process"]
    if not bad:
        print("no confessional or first-person process narration found")
        return 0
    print("CONFESSIONAL REGISTER — %d line(s). The content is probably true and belongs;" % len(bad))
    print("the phrasing is what a referee hears as defensive. Restate as fact.\n")
    for ln, why, txt in bad:
        print("  line %-5d %s" % (ln, why))
        print("      %s" % txt[:100])
    return 1


if __name__ == "__main__":
    sys.exit(main())
