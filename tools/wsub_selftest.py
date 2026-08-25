#!/usr/bin/env python3
"""Selftest for wsub.Patcher — especially that a late failure does not discard earlier edits.

Run:  python wsub_selftest.py       (exit 1 on any failure)
"""
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

for f in (p1, p2, p3, p4, p5, script, script3, script4):
    try:
        os.remove(f)
    except OSError:
        pass

print("\n%s" % ("ALL PASS" if not FAIL else "FAILED: " + ", ".join(FAIL)))
sys.exit(1 if FAIL else 0)
