#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md 파일을 렌더해서 «기본 브라우저»로 연다. 설치할 것이 없다.

왜 있나 (2026-08-22, 사용자 요청):
  이 랩의 작업 문서가 거의 다 .md다 — SESSION_LOG · notes/ · ANALYSIS_PLAN · 제출 초록.
  그런데 읽으려면 VS Code나 Cursor를 띄워야 했고 사용자가 「읽기만 하고 싶은데 IDE가 뜬다」고
  했다. Typora(유료)·Obsidian(vault 개념) 대신, 이미 있는 파이썬으로 렌더해 브라우저에 띄운다.

무엇을 신경 썼나:
  · **한글 가독성** — 본문 맑은 고딕, 코드는 D2Coding→Consolas. 줄간격을 넉넉히.
  · **표** — 이 랩 문서는 표가 많다. markdown-it의 table 확장을 켠다.
  · 이 랩이 쓰는 기호(«» 「」 ⚠ 🔴 ✅ ⛔ ⭐ ═ ▼)가 깨지지 않게 UTF-8 고정.
  · 렌더 결과는 tempdir에 둔다. 원본 옆에 .html을 만들지 않는다(OneDrive가 동기화한다).
  · 라이브러리가 없어도 «돌아간다» — markdown_it > mistune > 내장 최소 변환기 순으로 내려간다.

USAGE
    python view_md.py FILE.md          # 렌더해서 브라우저로 열기
    python view_md.py FILE.md --print  # HTML 경로만 출력(열지 않음)
    python view_md.py --assoc          # .md 더블클릭 연결 방법 안내
"""
import html as _html
import os
import re
import sys
import tempfile
import webbrowser

CSS = """
:root { color-scheme: light dark; }
* { box-sizing: border-box; }
body {
  margin: 0 auto; padding: 44px 28px 120px; max-width: 900px;
  font-family: "Malgun Gothic","맑은 고딕",-apple-system,"Segoe UI",sans-serif;
  font-size: 16.5px; line-height: 1.78; color: #1c1e21; background: #fff;
  word-break: keep-all; overflow-wrap: anywhere;
}
h1,h2,h3,h4 { line-height: 1.35; margin: 1.9em 0 .6em; font-weight: 700; }
h1 { font-size: 1.85em; border-bottom: 2px solid #d7dbe0; padding-bottom: .3em; }
h2 { font-size: 1.42em; border-bottom: 1px solid #e6e9ec; padding-bottom: .25em; }
h3 { font-size: 1.16em; }
h1:first-child { margin-top: 0; }
p, li { margin: .5em 0; }
ul, ol { padding-left: 1.5em; }
code { font-family: "D2Coding",Consolas,"Courier New",monospace; font-size: .89em;
       background: #f1f3f5; padding: .12em .38em; border-radius: 4px; }
pre { background: #f6f8fa; padding: 14px 16px; border-radius: 8px; overflow-x: auto;
      border: 1px solid #e6e9ec; }
pre code { background: none; padding: 0; font-size: .86em; line-height: 1.55; }
blockquote { margin: 1em 0; padding: .1em 1.1em; border-left: 4px solid #c9ced4;
             color: #47505a; background: #fafbfc; }
table { border-collapse: collapse; margin: 1.1em 0; width: 100%; font-size: .95em; }
th, td { border: 1px solid #dde1e6; padding: 7px 11px; text-align: left;
         vertical-align: top; }
th { background: #f2f4f6; font-weight: 700; }
tr:nth-child(even) td { background: #fbfcfd; }
hr { border: none; border-top: 1px solid #e0e4e8; margin: 2.2em 0; }
a { color: #1257a8; }
img { max-width: 100%; }
.meta { color: #7a838d; font-size: .82em; margin-bottom: 26px;
        border-bottom: 1px dashed #dde1e6; padding-bottom: 12px; }
@media (prefers-color-scheme: dark) {
  body { background: #14171a; color: #dfe3e7; }
  h1 { border-color: #333a41; } h2 { border-color: #262c32; }
  code { background: #22272c; } pre { background: #1a1f24; border-color: #2b3238; }
  blockquote { background: #191d21; border-left-color: #3a434b; color: #aab3bc; }
  th, td { border-color: #2e353c; } th { background: #1e2429; }
  tr:nth-child(even) td { background: #181c20; }
  hr { border-top-color: #2b3238; } a { color: #6fa8ea; }
  .meta { color: #79828b; border-bottom-color: #2b3238; }
}
"""


def _render(text):
    """markdown_it > mistune > 내장 최소 변환기. 없다고 죽지 않는다."""
    try:
        from markdown_it import MarkdownIt
        md = MarkdownIt("commonmark", {"html": False, "linkify": True, "typographer": False})
        md.enable(["table", "strikethrough"])
        return md.render(text), "markdown-it"
    except ImportError:
        pass
    try:
        import mistune
        return mistune.create_markdown(plugins=["table", "strikethrough"])(text), "mistune"
    except ImportError:
        pass
    return _fallback(text), "내장(최소)"


def _fallback(text):
    """라이브러리가 하나도 없을 때. 표·제목·코드블록·목록만 처리한다."""
    out, in_code, in_tbl = [], False, False
    for ln in text.split("\n"):
        if ln.strip().startswith("```"):
            out.append("</pre>" if in_code else "<pre>")
            in_code = not in_code
            continue
        if in_code:
            out.append(_html.escape(ln))
            continue
        if ln.strip().startswith("|") and ln.strip().endswith("|"):
            cells = [c.strip() for c in ln.strip().strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                continue                      # 구분행
            tag = "th" if not in_tbl else "td"
            if not in_tbl:
                out.append("<table>")
                in_tbl = True
            out.append("<tr>" + "".join("<%s>%s</%s>" % (tag, _inline(c), tag)
                                        for c in cells) + "</tr>")
            continue
        if in_tbl:
            out.append("</table>")
            in_tbl = False
        m = re.match(r"^(#{1,4})\s+(.*)$", ln)
        if m:
            n = len(m.group(1))
            out.append("<h%d>%s</h%d>" % (n, _inline(m.group(2)), n))
        elif re.match(r"^\s*[-*]\s+", ln):
            out.append("<li>%s</li>" % _inline(re.sub(r"^\s*[-*]\s+", "", ln)))
        elif ln.strip().startswith(">"):
            out.append("<blockquote>%s</blockquote>" % _inline(ln.strip().lstrip("> ")))
        elif re.match(r"^-{3,}$|^={3,}$", ln.strip()):
            out.append("<hr>")
        elif ln.strip():
            out.append("<p>%s</p>" % _inline(ln))
    if in_tbl:
        out.append("</table>")
    return "\n".join(out)


def _inline(s):
    s = _html.escape(s)
    s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"<em>\1</em>", s)
    return s


def build(path):
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    body, engine = _render(text)
    st = os.stat(path)
    import datetime
    mtime = datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")
    page = (
        "<!doctype html><html lang='ko'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<title>%s</title><style>%s</style></head><body>"
        "<div class='meta'>%s · 수정 %s · %s</div>%s</body></html>"
        % (_html.escape(os.path.basename(path)), CSS,
           _html.escape(os.path.abspath(path)), mtime, engine, body))

    # 원본 옆이 아니라 tempdir에 쓴다 — OneDrive가 .html까지 동기화하면 지저분해진다.
    out = os.path.join(tempfile.gettempdir(), "mdview",
                       re.sub(r"[^\w.-]", "_", os.path.basename(path)) + ".html")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(page)
    return out


ASSOC = r"""
[먼저] 렌더가 되는지부터 확인:
    python "{tool}" "{sample}"

[그다음] .md 를 «더블클릭»으로 열리게 하려면 — 관리자 권한으로 한 번만.

⚠ assoc·ftype 은 «cmd 내장 명령»이라 PowerShell에서 그대로 치면
   "assoc : 용어를 인식할 수 없습니다" 가 난다. 셸에 맞는 쪽을 쓸 것.

  ● 관리자 «명령 프롬프트(cmd)» 인 경우
    assoc .md=MarkdownFile
    ftype MarkdownFile="{py}" "{tool}" "%1"

  ● 관리자 «PowerShell» 인 경우 (cmd 를 거쳐야 한다)
    cmd /c 'assoc .md=MarkdownFile'
    cmd /c 'ftype MarkdownFile="{py}" "{tool}" "%1"'

되돌리기:  cmd /c 'assoc .md='

⚠ 시스템 설정을 바꾼다. 더블클릭이 필요 없으면 그냥 위 [먼저] 명령만 쓰면 된다.
"""


def main():
    args = [a for a in sys.argv[1:]]
    if "--assoc" in args:
        # 샘플 경로를 같이 찍어 «먼저 렌더를 확인»하게 한다. --assoc만 돌리고
        # 「됐다」고 넘어가면 정작 렌더가 되는지는 확인하지 않은 것이 된다.
        sample = os.path.join(os.path.dirname(os.path.dirname(
            os.path.dirname(os.path.abspath(__file__)))), "SESSION_LOG.md")
        if not os.path.isfile(sample):
            sample = "파일.md"
        print(ASSOC.format(py=sys.executable, tool=os.path.abspath(__file__),
                           sample=sample))
        return 0
    only_print = "--print" in args
    files = [a for a in args if not a.startswith("--")]
    if not files:
        print(__doc__)
        return 1
    for f in files:
        if not os.path.isfile(f):
            print("없는 파일: %s" % f)
            return 1
        out = build(f)
        if only_print:
            print(out)
        else:
            webbrowser.open("file:///" + out.replace("\\", "/"))
            print("열었다: %s" % out)
    return 0


if __name__ == "__main__":
    # Windows 콘솔은 cp949 라 한글·긴줄표(—)·기호 출력에서 UnicodeEncodeError 로 죽는다.
    # 결과를 다 만들어 놓고 «찍는 순간» 죽으므로, 부르는 쪽에는 도구가 고장난 것처럼 보인다.
    # ⚠ 모듈 최상단이 아니라 여기 두는 이유: 이 파일이 import 되기도 하면 최상단
    #    reconfigure 가 «호출자»의 인코딩을 바꾼다. 스크립트로 실행할 때만 돌게 한다.
    try:
        import sys as _s
        _s.stdout.reconfigure(encoding='utf-8', errors='replace')
        _s.stderr.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        pass
    sys.exit(main())
