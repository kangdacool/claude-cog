#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PostToolUse 훅 — .docx를 «만드는» 코드를 쓰는데 docx-editing 스킬을 안 쓰면 말해 준다.

    등록: settings.json 의 PostToolUse (matcher: Write|Edit)

왜 있나
    2026-08-25에 docx 제작 코드를 `docx-editing` 스킬로 모았다. 그런데 스킬은 «내가
    부를 때만» 뜬다 -- 그날 실패 셋이 전부 「맞는 답이 이미 있는데 못 찾음」이었고,
    그중 하나는 스킬을 만든 «직후»에도 kit을 모른 채 마크다운 파서를 다시 짠 것이었다.
    도구를 만드는 것과 그 도구가 «제때 불리는 것»은 다른 문제다
    (reference/tool_taxonomy.md 승격 규칙 · deliverable_guard.py 선례).

    이 훅은 그 «순간»을 만든다: python-docx로 문서를 «조립하는» 파일을 쓰는 순간.

무엇을 잡나 (셋 다 참일 때만)
    ① .py 파일이고
    ② python-docx를 «조립»에 쓴다 (add_paragraph / add_heading / .save)
    ③ docx_kit 을 안 쓴다
무엇을 안 잡나
    · 스킬 자신의 scripts/ 와 agent/tools 의 shim (정본과 껍데기)
    · docx를 «읽기만» 하는 감사 도구 (Document(path)만 쓰고 조립하지 않는다)
    · .docx를 안 만지는 파일

막지 않고 «알린다». 직접 조립해야 할 정당한 경우가 있고, 막으면 그때 훅을 끄게 된다 --
꺼진 가드는 없는 가드다. 대신 이름을 부르고 정확한 import 한 줄을 준다.
"""
import json
import os
import re
import sys

for _s in (sys.stderr, sys.stdout):
    if hasattr(_s, "reconfigure"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# 조립 신호 — 이 중 하나라도 있으면 «만드는» 코드다. 읽기 전용 감사와 가르는 축이다.
BUILDS = re.compile(r"\.add_paragraph\(|\.add_heading\(|\.add_table\(|\.add_picture\(|"
                    r"doc\.save\(|document\.save\(")
USES_DOCX = re.compile(r"^\s*(?:from\s+docx[\s.]|import\s+docx\b)", re.M)
HAS_KIT = re.compile(r"docx_kit|docx-editing")

# 정본·껍데기는 당연히 직접 쓴다.
EXEMPT = re.compile(r"skills[/\\]docx-editing[/\\]scripts[/\\]|"
                    r"agent[/\\]tools[/\\](md_to_docx|manuscript_table|brief_builder)\.py$|"
                    r"docx_kit")

MSG = """\
이 파일은 python-docx로 문서를 «조립»하는데 docx-editing 스킬을 쓰지 않습니다.

  {path}

직접 짜기 전에 이미 있는 것을 보십시오:

    import os, sys
    sys.path.insert(0, os.path.expanduser(
        os.path.join("~", ".claude", "skills", "docx-editing", "scripts")))
    from docx_kit import (manuscript_shell, brief_shell,          # 골격
                          render_markdown, add_journal_table,     # 부품
                          brief_table, render_docx)

⭐ **골격부터 봅니다.** 원고 totale 은 manuscript_shell, 국문 브리프는 brief_shell —
표지·쪽나눔·표번호·그림캡션·숫자 게이트까지 배치가 들어 있습니다. 부품만 쓰면 그 배치를
매번 다시 엮게 되고, 그러다 부품까지 다시 짜게 됩니다(실측: 직접 짠 빌더 53개 중 51개가
kit 에 있는 「굵게 처리」를 재구현했습니다).

장르가 표를 정합니다 — 저널 원고는 add_journal_table(세로줄 없음), 국문 브리프는
brief_table(격자·색 헤더). 자세한 것은 SKILL.md.

⚠️ 정말 직접 조립해야 하면 그대로 진행하십시오. 이 훅은 막지 않습니다.
"""


def main():
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return 0
    try:
        if payload.get("tool_name") not in ("Write", "Edit"):
            return 0
        ti = payload.get("tool_input", {}) or {}
        path = ti.get("file_path") or ""
        if not path.endswith(".py") or EXEMPT.search(path.replace("\\", "/")):
            return 0

        text = ti.get("content") or ti.get("new_string") or ""
        if not text and os.path.exists(path):
            try:
                text = open(path, encoding="utf-8", errors="replace").read()
            except OSError:
                return 0
        # Edit는 조각만 오므로, 조각에 신호가 없으면 파일 전체를 본다.
        if not BUILDS.search(text) and os.path.exists(path):
            try:
                text = open(path, encoding="utf-8", errors="replace").read()
            except OSError:
                pass

        if not (USES_DOCX.search(text) and BUILDS.search(text)):
            return 0
        if HAS_KIT.search(text):
            return 0

        sys.stderr.write(MSG.format(path=path))
        return 2                                   # stderr가 Claude에게 전달된다
    except Exception:
        return 0                                   # 실패는 조용히 통과


if __name__ == "__main__":
    sys.exit(main())
