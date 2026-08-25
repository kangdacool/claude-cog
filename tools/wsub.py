# -*- coding: utf-8 -*-
"""Whitespace-tolerant find/replace for generator sources.

Why this exists (2026-08-22): manuscript and poster generators hold prose wrapped at 79 columns,
and the wrap position moves every time a sentence is edited. Anchoring on the exact line breaks
failed five times in a row. `sub()` matches on words with `\\s+` between them, so a patch written
against the *rendered* text finds the *wrapped* source.

⚠ The trap this file now closes (hit four times on 2026-08-24): the old version accumulated edits
in memory and only wrote at the end, so a single failed anchor at the end of a batch called
`sys.exit()` and **silently discarded every successful edit before it**. The prints said
"ok ... ok ... ok ... ANCHOR FAIL" and the file was untouched. Each `sub()` now writes to disk
immediately, so a late failure keeps the earlier work and tells you exactly where it stopped.

Usage:
    from wsub import Patcher
    p = Patcher("R/manuscript/05_manuscript.py")
    p.sub("old text as it reads", "new text with real line breaks", "label")
    p.write()          # optional; every sub() has already written

Anchors must match exactly once. Zero matches or several both stop the run -- a patch that could
land in two places is a patch you did not mean.

⚠ Indentation, the trap of 2026-08-25: the pattern is matched from its first *word*, so the
leading spaces of the first line are NOT part of the match and stay in the file. Write the
replacement's FIRST line flush-left and indent only the lines after it:

    p.sub("out = dl = None\n    timeout = 0",          # anchor: leading spaces don't matter
          "out = dl = None\n    timeout = 7200.0")     # <- first line NOT indented

Indenting that first line silently produced `        out = ...` under a 4-space block and the
file no longer parsed. For .py targets that is now caught: every sub() re-parses the result and,
if the edit broke the syntax, rolls back that one edit and names the label. Pass check=False to
turn it off (e.g. a batch that must pass through an invalid intermediate state).
"""
import ast
import io
import os
import re
import sys


class Patcher(object):
    def __init__(self, path, autowrite=True, check=None):
        self.path = path
        self.src = io.open(path, encoding="utf-8").read()
        self.n = 0
        self.autowrite = autowrite
        self.done = []
        # check=None -> decide by extension. Only Python can be checked this cheaply, and a
        # generator source is exactly where a silently-broken file costs the most.
        self.check = (os.path.splitext(path)[1].lower() == ".py") if check is None else check
        if self.check:
            try:
                ast.parse(self.src)
            except SyntaxError:
                # Already broken before we touched it; don't blame the first edit for it.
                self.check = False
                print("  note: %s does not parse to begin with -- syntax gate off"
                      % os.path.basename(path))

    def _flush(self):
        io.open(self.path, "w", encoding="utf-8").write(self.src)

    def sub(self, old, new, label):
        pat = re.compile(r"\s+".join(re.escape(t) for t in old.split()))
        hits = list(pat.finditer(self.src))
        if len(hits) != 1:
            # Keep what already worked. Discarding it is how a whole batch gets lost silently.
            if self.autowrite and self.n:
                self._flush()
            sys.exit("ANCHOR FAIL (%s): %d matches%s"
                     % (label, len(hits),
                        ("\n  kept %d earlier edit(s): %s" % (self.n, ", ".join(self.done)))
                        if self.n else ""))
        m = hits[0]
        before = self.src
        self.src = self.src[:m.start()] + new + self.src[m.end():]
        if self.check:
            try:
                ast.parse(self.src)
            except SyntaxError as e:
                self.src = before          # roll back this edit only
                if self.autowrite and self.n:
                    self._flush()
                sys.exit("SYNTAX BREAK (%s): %s at line %s\n"
                         "  this edit was rolled back; %d earlier edit(s) kept%s\n"
                         "  most likely: the replacement\'s FIRST line carries indentation that is "
                         "already in the file (see the module docstring)"
                         % (label, e.msg, e.lineno, self.n,
                            (": " + ", ".join(self.done)) if self.n else ""))
        self.n += 1
        self.done.append(label)
        if self.autowrite:
            self._flush()
        print("  ok", label)

    def write(self):
        self._flush()
        print("wrote %d edits" % self.n)
