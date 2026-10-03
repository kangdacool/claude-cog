#!/usr/bin/env python3
"""Selftest for wsub.Patcher — especially that a late failure does not discard earlier edits.

Run:  python wsub_selftest.py       (exit 1 on any failure)
"""
import ast
import importlib.util
import io
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("wsub", os.path.join(HERE, "wsub.py"))
W = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(W)

FAIL = []


def check(name, cond, detail=""):
    if cond:
        print("  ok   %s" % name)
    else:
        FAIL.append(name)
        print("  FAIL %s  %s" % (name, detail))


def _parses(text):
    import ast
    try:
        ast.parse(text)
        return True
    except SyntaxError:
        return False


def tmpfile(text, suffix=".txt"):
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    io.open(path, "w", encoding="utf-8").write(text)
    return path


# 1. Wrapped source is found by an anchor written as one line.
src = "alpha beta\ngamma delta epsilon\n"
p1 = tmpfile(src)
pat = W.Patcher(p1)
pat.sub("beta gamma delta", "REPLACED", "wrap")
got = io.open(p1, encoding="utf-8").read()
check("matches across a line break", "REPLACED" in got, repr(got))
check("writes immediately (no write() needed)", "REPLACED" in io.open(p1, encoding="utf-8").read())

# 2. A late failure keeps the earlier edits. Run in a subprocess because sub() exits the process.
p2 = tmpfile("one two\nthree four\n")
script = tmpfile("")
io.open(script, "w", encoding="utf-8").write(
    "import sys\n"
    "sys.path.insert(0, %r)\n" % HERE +
    "from wsub import Patcher\n"
    "p = Patcher(%r)\n" % p2 +
    "p.sub('one two', 'FIRST', 'a')\n"
    "p.sub('nonexistent anchor', 'X', 'b')\n"
    "p.write()\n")
r = subprocess.run([sys.executable, script], capture_output=True, text=True)
after = io.open(p2, encoding="utf-8").read()
check("late failure exits non-zero", r.returncode != 0, "rc=%s" % r.returncode)
check("earlier edit survives the failure", "FIRST" in after, repr(after))
check("failure names the surviving edits", "kept 1 earlier edit" in (r.stdout + r.stderr),
      (r.stdout + r.stderr)[-160:])

# 3. An anchor matching twice is refused (a patch that could land in two places is not meant).
p3 = tmpfile("same same\n")
script3 = tmpfile("")
io.open(script3, "w", encoding="utf-8").write(
    "import sys\n"
    "sys.path.insert(0, %r)\n" % HERE +
    "from wsub import Patcher\n"
    "Patcher(%r).sub('same', 'X', 'dup')\n" % p3)
r3 = subprocess.run([sys.executable, script3], capture_output=True, text=True)
check("ambiguous anchor refused", r3.returncode != 0 and "2 matches" in (r3.stdout + r3.stderr),
      (r3.stdout + r3.stderr)[-120:])
check("ambiguous anchor changes nothing", io.open(p3, encoding="utf-8").read() == "same same\n")

# 4. The syntax gate: an edit that breaks a .py is rolled back, not saved.
#    This is the 2026-08-25 failure -- the replacement\'s first line carried indentation that was
#    already in the file, so the block ended up double-indented and the module stopped parsing.
PY = "def f():\n    a = 1\n    b = 2\n    return a + b\n"
p4 = tmpfile(PY, suffix=".py")
script4 = tmpfile("")
io.open(script4, "w", encoding="utf-8").write(
    "import sys\n"
    "sys.path.insert(0, %r)\n" % HERE +
    "from wsub import Patcher\n"
    "p = Patcher(%r)\n" % p4 +
    "p.sub('a = 1', 'a = 1', 'harmless')\n"
    "p.sub('b = 2', '    b = 2', 'double indent')\n")   # <- first line wrongly indented
r4 = subprocess.run([sys.executable, script4], capture_output=True, text=True)
o4 = r4.stdout + r4.stderr
check("syntax break refused", r4.returncode != 0 and "SYNTAX BREAK" in o4, o4[-160:])
check("syntax break names the label", "double indent" in o4, o4[-160:])
check("broken edit rolled back (file still parses)",
      io.open(p4, encoding="utf-8").read() == PY, repr(io.open(p4, encoding="utf-8").read()))

# 5. A non-.py target is not syntax-checked (prose generators, .md, .txt).
p5 = tmpfile("hello world\n")
W.Patcher(p5).sub("world", "def (", "no check on txt")
check("non-py target skips the gate", "def (" in io.open(p5, encoding="utf-8").read())

# ------------------------------------------------------------------ sub_literal
# 6. The motivating shape: prose split by implicit concatenation. sub() cannot reach it
#    (a quote + newline + indent sit between the halves); sub_literal() can.
BUILDER = (
    "def build(doc):\n"
    "    ans(doc, 'The evidence body starts at High '\n"
    "             'because the design is randomised.')\n")
p6 = tmpfile(BUILDER, suffix=".py")
script6 = tmpfile("")
io.open(script6, "w", encoding="utf-8").write(
    "import sys\n"
    "sys.path.insert(0, %r)\n" % HERE +
    "from wsub import Patcher\n"
    "Patcher(%r).sub('The evidence body starts at High because the design is randomised.',\n"
    "                'X', 'via sub')\n" % p6)
r6 = subprocess.run([sys.executable, script6], capture_output=True, text=True)
check("sub() cannot reach implicit concatenation (that is why sub_literal exists)",
      r6.returncode != 0 and "0 matches" in (r6.stdout + r6.stderr), (r6.stdout + r6.stderr)[-120:])

pat6 = W.Patcher(p6)
pat6.sub_literal("The evidence body starts at High because the design is randomised.",
                 "The evidence body starts at High.", "shorten")
got6 = io.open(p6, encoding="utf-8").read()
check("sub_literal reaches it", "'The evidence body starts at High.'" in got6, repr(got6))
check("sub_literal leaves the file parsing", "ans(doc," in got6 and got6.count("ans(") == 1,
      repr(got6))

# 7. Inside brackets a long replacement is re-wrapped, not left as one 300-column line.
long_txt = " ".join(["word%02d" % i for i in range(40)])
p7 = tmpfile("def f():\n    ans(doc, 'short')\n", suffix=".py")
W.Patcher(p7).sub_literal("short", long_txt, "rewrap")
got7 = io.open(p7, encoding="utf-8").read()
check("long value is re-wrapped under 100 columns",
      max(len(ln) for ln in got7.splitlines()) < 100,
      "longest=%d" % max(len(ln) for ln in got7.splitlines()))
_c7 = [n.value for n in ast.walk(ast.parse(got7)) if isinstance(n, ast.Constant)]
check("re-wrapped literal still carries the same value", long_txt in _c7, repr(_c7)[:120])

# 7b. OUTSIDE brackets it must stay on one line -- a split literal there is a SyntaxError, not a
#     style problem. (Found by this selftest on 2026-08-31, before the tool was ever used.)
p7b = tmpfile("def f():\n    t = 'short'\n    return t\n", suffix=".py")
W.Patcher(p7b).sub_literal("short", long_txt, "no-wrap outside brackets")
got7b = io.open(p7b, encoding="utf-8").read()
check("outside brackets the literal is not split", got7b.count("'") == 2, repr(got7b[:120]))
check("outside brackets the file still parses", _parses(got7b), repr(got7b[:120]))

# 8. Two literals with the same value are refused -- same rule as sub().
p8 = tmpfile("a = 'dup'\nb = 'dup'\n", suffix=".py")
script8 = tmpfile("")
io.open(script8, "w", encoding="utf-8").write(
    "import sys\n"
    "sys.path.insert(0, %r)\n" % HERE +
    "from wsub import Patcher\n"
    "Patcher(%r).sub_literal('dup', 'X', 'ambiguous')\n" % p8)
r8 = subprocess.run([sys.executable, script8], capture_output=True, text=True)
check("sub_literal refuses an ambiguous value",
      r8.returncode != 0 and "2 matches" in (r8.stdout + r8.stderr), (r8.stdout + r8.stderr)[-120:])
check("sub_literal changes nothing when ambiguous",
      io.open(p8, encoding="utf-8").read() == "a = 'dup'\nb = 'dup'\n")

# 9. A *substring* of a literal does not match -- whole values only, by design.
p9 = tmpfile("a = 'alpha beta gamma'\n", suffix=".py")
script9 = tmpfile("")
io.open(script9, "w", encoding="utf-8").write(
    "import sys\n"
    "sys.path.insert(0, %r)\n" % HERE +
    "from wsub import Patcher\n"
    "Patcher(%r).sub_literal('beta', 'X', 'partial')\n" % p9)
r9 = subprocess.run([sys.executable, script9], capture_output=True, text=True)
check("sub_literal refuses a partial value (no guessing where to cut)",
      r9.returncode != 0 and "0 matches" in (r9.stdout + r9.stderr), (r9.stdout + r9.stderr)[-120:])

for f in (p1, p2, p3, p4, p5, p6, p7, p7b, p8, p9,
          script, script3, script4, script6, script8, script9):
    try:
        os.remove(f)
    except OSError:
        pass

print("\n%s" % ("ALL PASS" if not FAIL else "FAILED: " + ", ".join(FAIL)))
sys.exit(1 if FAIL else 0)
