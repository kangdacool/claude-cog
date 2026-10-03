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

⚠ `sub()` does NOT reach prose split by *implicit concatenation* -- the shape a docx/pptx builder
uses most:

    ans(doc, 'The evidence body starts at High '
             'because the design is randomised.')

Between the two halves sit a quote, a newline and indentation, so an anchor written as the finished
sentence never matches. Measured 2026-08-28: 17 of 17 attempts missed, silently, because `sub()`
reported ANCHOR FAIL per call and the batch just looked "not found". Use `sub_literal()` there --
it tokenises, joins each adjacent-STRING run into its *value*, and swaps the whole run for a fresh
literal re-wrapped at the original column. Anchor on the sentence as the reader sees it.

    p.sub_literal("The evidence body starts at High because the design is randomised.",
                  "The evidence body starts at High.", "shorten")

⚠ BULK-EDITING STRING VALUES: never .strip() a fragment of an implicit concatenation.
The trailing space of one fragment IS the word break at the join:
    '…review performance '  +  'quality…'   → strip → '…review performancequality…'
Measured 2026-09-03: a cleanup script that stripped every literal glued 12 word boundaries
across four builders, and the damage is invisible in the source -- each half looks fine alone.
If you must transform values in bulk, operate on the JOINED value (_string_runs() gives it),
or preserve leading/trailing whitespace per fragment. Then scan for glue: a join whose left
side ends and right side begins with non-space is suspect, because this codebase wraps only
at spaces.

Rule of thumb: editing **prose that a builder prints** -> `sub_literal`. Editing **code** (calls,
arguments, control flow) -> `sub`.

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
import tokenize


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
            self._anchor_fail(label, len(hits))
        m = hits[0]
        self._commit(self.src[:m.start()] + new + self.src[m.end():], label)

    def _anchor_fail(self, label, n_hits):
        # Keep what already worked. Discarding it is how a whole batch gets lost silently.
        if self.autowrite and self.n:
            self._flush()
        sys.exit("ANCHOR FAIL (%s): %d matches%s"
                 % (label, n_hits,
                    ("\n  kept %d earlier edit(s): %s" % (self.n, ", ".join(self.done)))
                    if self.n else ""))

    def _commit(self, new_src, label):
        before = self.src
        self.src = new_src
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

    def _string_runs(self):
        """[(start, end, joined value, column, inside_brackets)] for every run of adjacent
        STRING tokens -- i.e. every implicitly concatenated literal, wrapped or not.

        `inside_brackets` says whether the run sits inside ( [ { . It decides whether the
        replacement may be split over several lines: implicit concatenation across a line break is
        only legal inside brackets."""
        lines = self.src.splitlines(keepends=True)
        starts = [0]
        for ln in lines:
            starts.append(starts[-1] + len(ln))

        def off(pos):
            row, col = pos
            return starts[row - 1] + col

        toks = list(tokenize.generate_tokens(io.StringIO(self.src).readline))
        depth_at, depth = [], 0
        for t in toks:                      # bracket depth *before* each token
            depth_at.append(depth)
            if t.type == tokenize.OP:
                if t.string in "([{":
                    depth += 1
                elif t.string in ")]}":
                    depth -= 1
        out, i = [], 0
        while i < len(toks):
            if toks[i].type != tokenize.STRING:
                i += 1
                continue
            j, parts, last = i, [], i
            while j < len(toks) and toks[j].type in (
                    tokenize.STRING, tokenize.NL, tokenize.INDENT):
                if toks[j].type == tokenize.STRING:
                    parts.append(toks[j])
                    last = j
                j += 1
            try:
                val = "".join(ast.literal_eval(t.string) for t in parts)
            except Exception:
                i += 1
                continue
            out.append((off(parts[0].start), off(toks[last].end), val, parts[0].start[1],
                        depth_at[i] > 0))
            i = last + 1
        return out

    @staticmethod
    def _relit(value, col, wrap=True, width=96):
        """Re-wrap a value as a Python literal that keeps lines under `width` columns.

        wrap=False emits one line however long: outside brackets a split literal is a SyntaxError.
        """
        if not wrap:
            return repr(value)
        room = max(30, width - col)
        chunks, cur = [], ""
        for word in value.split(" "):
            if cur and len(cur) + len(word) + 1 > room:
                chunks.append(cur + " ")
                cur = word
            else:
                cur = (cur + " " + word) if cur else word
        chunks.append(cur)
        return ("\n" + " " * col).join(repr(c) for c in chunks)

    def sub_literal(self, old, new, label):
        """Replace a string literal by its *joined value* -- reaches implicit concatenation.

        `old` is the sentence as the builder prints it, with the halves already joined. The whole
        adjacent-STRING run is swapped for a fresh literal re-wrapped at the original column, so
        the source stays readable. Same discipline as sub(): exactly one match or it stops, and a
        broken .py is rolled back.

        ⚠ An `old` that is a *substring* of a literal will not match -- this compares whole
        values. That is deliberate: a partial match inside a run would have to guess where to cut
        the literal, and guessing is how prose gets sliced mid-word.
        """
        hits = [r for r in self._string_runs() if r[2] == old]
        if len(hits) != 1:
            self._anchor_fail(label, len(hits))
        start, end, _val, col, in_brackets = hits[0]
        self._commit(self.src[:start] + self._relit(new, col, wrap=in_brackets)
                     + self.src[end:], label)

    def write(self):
        self._flush()
        print("wrote %d edits" % self.n)
