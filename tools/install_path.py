"""Make the lab's shared Python tools importable on this machine.

The tools live in OneDrive and sync across machines, but Python has no way to
find them unless site-packages is told where they are. This writes one .pth file
so that `import manuscript_table` (and brief_builder, flowchart_generator, ...)
just works - from any project, any directory, without a sys.path.insert in every
consuming script.

Hardcoding `sys.path.insert(0, "D:/onedrive/...")` into project script   # noleaks is the
alternative, and it is worse: it is an absolute path in version-controlled
project code, it breaks on any machine with a different layout, and it has to be
repeated in every file that imports.

    python install_path.py          # install
    python install_path.py --check  # report only

Run once per machine (and again if Python is upgraded to a new minor version,
since site-packages moves).
"""
import argparse
import os
import site
import sys

PTH_NAME = "kang_lab_tools.pth"
TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))


def site_packages():
    cands = [p for p in site.getsitepackages() if p.endswith("site-packages")]
    if not cands:
        cands = [site.getusersitepackages()]
    return cands[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="report status, change nothing")
    args = ap.parse_args()

    sp = site_packages()
    pth = os.path.join(sp, PTH_NAME)
    print("python        : %s" % sys.version.split()[0])
    print("tools dir     : %s" % TOOLS_DIR)
    print("site-packages : %s" % sp)
    print("path file     : %s%s" % (pth, "" if os.path.exists(pth) else "  (missing)"))

    if os.path.exists(pth):
        current = open(pth, encoding="utf-8").read().strip()
        if os.path.normcase(current) == os.path.normcase(TOOLS_DIR.replace("\\", "/")):
            print("\nalready installed and pointing here")
        else:
            print("\ninstalled but points elsewhere: %s" % current)

    if args.check:
        pass
    else:
        try:
            with open(pth, "w", encoding="utf-8") as f:
                f.write(TOOLS_DIR.replace("\\", "/") + "\n")
            print("\nwrote %s" % pth)
        except OSError as e:
            print("\ncould not write (%s)." % e)
            print("Either run as administrator, or set PYTHONPATH=%s instead." % TOOLS_DIR)
            return 1

    # prove it - a fresh interpreter, so we are not fooled by this process's sys.path
    import subprocess
    r = subprocess.run([sys.executable, "-c",
                        "import manuscript_table, os;"
                        "print(os.path.dirname(manuscript_table.__file__))"],
                       capture_output=True, text=True, cwd=os.path.expanduser("~"))
    if r.returncode == 0:
        print("verified      : `import manuscript_table` resolves to %s" % r.stdout.strip())
        return 0
    print("VERIFY FAILED : %s" % (r.stderr.strip().splitlines() or ["?"])[-1])
    return 1


if __name__ == "__main__":
    sys.exit(main())
