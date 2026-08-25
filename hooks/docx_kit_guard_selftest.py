#!/usr/bin/env python3
"""Self-test for docx_kit_guard.py. 절반은 «울리지 않는 것»을 증명한다."""
import json
import os
import subprocess
import sys

GUARD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docx_kit_guard.py")

BUILD = "from docx import Document\nd = Document()\nd.add_paragraph('x')\nd.save('a.docx')\n"
READ = "from docx import Document\nd = Document('a.docx')\nprint(len(d.paragraphs))\n"
KIT = ("import sys\nsys.path.insert(0, 'x')\nfrom docx_kit import render_markdown\n"
       "from docx import Document\nd = Document()\nd.add_paragraph('x')\nd.save('a.docx')\n")

# (파일경로, 내용, 울려야 하나)
CASES = [
    # ── 울려야 한다 ──────────────────────────────────────────────────────
    ("proj/tools/build_report.py", BUILD, True),
    ("proj/make_totale.py", "import docx\nd = docx.Document()\nd.add_heading('T')\nd.save('t.docx')\n", True),
    # ── 울리면 «안 된다» ─────────────────────────────────────────────────
    ("proj/tools/audit_fonts.py", READ, False),            # 읽기만 하는 감사
    ("proj/tools/build_report.py", KIT, False),            # kit을 이미 쓴다
    ("skills/docx-editing/scripts/md_to_docx.py", BUILD, False),   # 정본
    ("agent/tools/manuscript_table.py", BUILD, False),     # shim
    ("agent/tools/docx_kit.py", BUILD, False),             # 진입점 자신
    ("proj/tools/build_deck.py", "from pptx import Presentation\n", False),  # 다른 형식
    ("proj/notes.md", BUILD, False),                       # .py 아님
    ("proj/tools/plot.py", "import pandas as pd\n", False),  # docx 무관
]


def run(path, content, tool="Write"):
    p = subprocess.run([sys.executable, GUARD],
                       input=json.dumps({"tool_name": tool,
                                         "tool_input": {"file_path": path,
                                                        "content": content}}),
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return p.returncode


bad = []
for path, content, should in CASES:
    rc = run(path, content)
    fired = rc == 2
    ok = fired == should
    print(f"  [{'OK  ' if ok else '실패'}] {'울림' if fired else '조용':4s} "
          f"(기대 {'울림' if should else '조용'})  {path}")
    if not ok:
        bad.append(path)

# 다른 도구는 무시해야 한다
if run("proj/tools/build_report.py", BUILD, tool="Bash") == 2:
    print("  [실패] Bash 호출에도 울린다"); bad.append("tool_name")
else:
    print("  [OK  ] Write/Edit 이외에는 울리지 않는다")

print()
if bad:
    print(f"★ {len(bad)}건 실패: {bad}")
    sys.exit(1)
print(f"{len(CASES) + 1}/{len(CASES) + 1} 통과")
