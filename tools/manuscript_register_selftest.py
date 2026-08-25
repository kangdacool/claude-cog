#!/usr/bin/env python3
"""Self-test for manuscript_register.py.

Half of these cases exist to prove the scanner does NOT fire. A register checker that
flags stated limitations or pre-registration language would push a manuscript toward
sounding more certain than its evidence, which is worse than the problem it fixes.
"""
import sys

from manuscript_register import scan

MUST_FLAG = [
    "because comparing it the other way misled us once.",
    "One caution emerged in obtaining these: a first version reported 0.953.",
    "The first version of this comparison said the opposite.",
    "That this is not a small-sample floor was not assumed but checked.",
    "What the rule turns on is worth stating plainly.",
    "Under S2, the test that matters, the ordering was sharper.",
    "We traced it to a variance scaling error after we noticed the ratio.",
    "To be honest, the grid is coarse.",
    # the repair sense of "fixed" -- still a hit
    "We fixed the bug and re-ran the cells.",
]

MUST_NOT_FLAG = [
    # stated limitations
    "The scope boundary is bracketed rather than located.",
    "The surcharge for stratified schedules is measured but not modeled.",
    "Below 1,500 observations the threshold is unsupported.",
    # pre-registration and negative results
    "The rule was fixed before the confirming experiment was run.",
    "Two gates fired, both on unpaired comparisons.",
    "Six remedies were tested and none repaired it.",
    # precision contrasts that carry meaning
    "Sigma is conditional rather than marginal, which understates the load.",
    "This is a bound, not a second threshold.",
    "It governs the interval and not the point estimate.",
    # ordinary first person doing science
    "We simulated longitudinal mechanisms with treatment-confounder feedback.",
    "We report relative bias, coverage, and the standard-error ratio.",
    # the pre-registration sense of "fixed" -- the false positive this scanner shipped with
    "Before the second structure was run we fixed the acceptance rule.",
    "We fixed the threshold by a rule stated in advance.",
]

fail = 0
for s in MUST_FLAG:
    r = scan(s)
    if not (r["confession"] or r["process"]):
        print("MISS  (should flag): %s" % s)
        fail += 1
for s in MUST_NOT_FLAG:
    r = scan(s)
    if r["confession"] or r["process"]:
        why = (r["confession"] + r["process"])[0][1]
        print("FALSE POSITIVE (%s): %s" % (why, s))
        fail += 1

print("\n%d/%d cases correct" % (len(MUST_FLAG) + len(MUST_NOT_FLAG) - fail,
                                 len(MUST_FLAG) + len(MUST_NOT_FLAG)))
sys.exit(1 if fail else 0)
