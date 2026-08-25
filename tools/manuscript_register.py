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
import re
import sys

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

    t = io.open(a.file, encoding="utf-8").read()
    if a.start and a.start in t:
        t = t[t.index(a.start):]
    if a.end and a.end in t:
        t = t[:t.index(a.end)]
    r = scan(t)

    print("body: %d words\n" % r["words"])
    for label, (n, per) in r["density"].items():
        note = "  <- clusters read as generated prose" if per and per < 130 else ""
        print("  %-38s %3d%s%s" % (label, n,
                                   ("  (one per %d words)" % per) if per else "", note))
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
