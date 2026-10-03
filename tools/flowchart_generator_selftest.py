# -*- coding: utf-8 -*-
"""flowchart_generator.py 자기시험 -- autofit(넓히기 -> 줄바꿈).

⚠ CORE §3: 「게이트도 시험한다 -- 통과할 것·실패할 것을 하나씩 넣어 본다.」
   여기서는 세 입력을 넣는다:
     (a) 이미 들어가는 글 -> 폭이 «기본값 그대로»여야 한다(기존 그림이 리플로우되면 안 된다)
     (b) 조금 넘치는 글   -> 폭이 «넓어져야» 한다(줄바꿈은 아직 아님)
     (c) 아주 긴 한 단어  -> 넓힘 상한에 걸려 «줄바꿈»으로 넘어가야 한다

Usage: python flowchart_generator_selftest.py     (exit 0 = 통과)
"""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import flowchart_generator as FG  # noqa: E402

FAILS = []


def check(cond, msg):
    print(("  ok    " if cond else "  FAIL  ") + msg)
    if not cond:
        FAILS.append(msg)


W_IN, FS, FAM = 6.5, 9, "Arial"
BASE_L, BASE_R = 4.5, 4.0


def fit(texts, base, other):
    return FG._fit_column(texts, base, other, W_IN, FAM, FS)


def main():
    print("(a) 이미 들어가는 글 -> 기본 폭 유지 (기존 그림 보호)")
    w, out = fit(["KELS2013 cohort\n(n = 7,324)"], BASE_L, BASE_R)
    check(abs(w - BASE_L) < 1e-9, "폭이 기본값 4.5 그대로 (%.3f)" % w)
    check(out[0].count("\n") == 1, "줄 수 불변 (줄바꿈이 끼어들지 않았다)")

    print("(b) 조금 넘치는 글 -> 폭이 넓어진다")
    long1 = "Responded to adolescent mental health items in at least 4 of waves 3-8"
    w2, out2 = fit([long1], BASE_L, BASE_R)
    check(w2 > BASE_L, "폭이 넓어졌다 (%.3f > 4.5)" % w2)

    print("(c) 넓혀도 안 되는 글 -> 줄바꿈으로 넘어간다")
    long2 = ("Participants responded to every adolescent mental health item in at "
             "least four of the six annual waves between 2015 and 2020 inclusive")
    w3, out3 = fit([long2], BASE_L, BASE_R)
    check(out3[0].count("\n") >= 1, "줄이 나뉘었다 (%d줄)" % (out3[0].count("\n") + 1))
    cap_share = (w3 / (FG.MARGIN_L + w3 + FG.GUTTER + BASE_R + FG.MARGIN_R))
    check(cap_share <= FG.MAX_SHARE + 1e-6,
          "폭이 상한(%.0f%%) 안 (%.1f%%)" % (FG.MAX_SHARE * 100, cap_share * 100))

    print("(d) 넓히거나 나눈 뒤 모든 줄이 실제로 들어간다")
    for label, w_, out_ in (("b", w2, out2), ("c", w3, out3)):
        room_in = (w_ / (FG.MARGIN_L + w_ + FG.GUTTER + BASE_R + FG.MARGIN_R)) * W_IN
        room_pts = (room_in - FG.PAD_IN) * 72.0
        widest = max(FG._line_pts(ln, FAM, FS) for ln in out_[0].split("\n"))
        check(widest <= room_pts + 1.0,
              "(%s) 가장 긴 줄 %.1fpt <= 여유 %.1fpt" % (label, widest, room_pts))

    print("(e) autofit=False 면 옛 동작 그대로 (폭 고정·줄 그대로)")
    tmp = os.path.join(tempfile.mkdtemp(), "f.png")
    FG.make_flowchart([{"box": long2, "exclude": None}], tmp, autofit=False)
    check(os.path.exists(tmp), "autofit=False 로도 그려진다")

    print("(f) 호출자의 steps 를 변형하지 않는다")
    steps = [{"box": long2, "exclude": None}]
    FG.make_flowchart(steps, os.path.join(tempfile.mkdtemp(), "g.png"))
    check(steps[0]["box"] == long2, "원본 steps 가 그대로다")

    print("(g) 두 갈래 흐름도 — 그려지고, 상자끼리 겹치지 않고, 글이 상자 안에 든다")
    trunk = [{"box": "Year 2 participants\n(n = 645)", "exclude": None},
             {"box": "Seen at Year 2 and Year 3\n(n = 353)",
              "exclude": "Excluded (n = 292)\n• Not seen at Year 3"}]
    arms = [{"box": "MCI at Year 2\n(n = 192)",
             "exclude": "Excluded (n = 21)\n• No facial task (n = 12)\n• Missing covariate (n = 9)",
             "final": "MCI analytic sample\n(n = 171)",
             "outcomes": ["Stable MCI\n(n = 112)", "Reverted\n(n = 52)", "Dementia\n(n = 7)"]},
            {"box": "Normal cognition at Year 2\n(n = 149)",
             "exclude": "Excluded (n = 13)\n• No facial task (n = 9)\n• Missing covariate (n = 4)",
             "final": "Normal analytic sample\n(n = 136)",
             "outcomes": ["Normal\n(n = 99)", "MCI or dementia\n(n = 37)"]}]
    tmp = os.path.join(tempfile.mkdtemp(), "split.png")
    geo = FG.make_split_flowchart(trunk, arms, tmp)
    check(os.path.exists(tmp), "두 갈래 PNG 가 만들어졌다")
    bx = geo["boxes"]
    over = [(a[0], b[0]) for i, a in enumerate(bx) for b in bx[i + 1:]
            if a[1] < b[3] and b[1] < a[3] and a[2] < b[4] and b[2] < a[4]]
    check(not over, "상자 겹침 없음 (%d)" % len(over))
    inside = all(0 <= b[1] and b[3] <= geo["W"] + 1e-6 and 0 <= b[2] and b[4] <= geo["H"] + 1e-6
                 for b in bx)
    check(inside, "모든 상자가 그림 안에 있다")
    fitok = all(max(FG._line_pts(ln, "Arial", 10) for ln in b[5].split("\n")) / 72.0
                <= (b[3] - b[1]) + 1e-6 for b in bx)
    check(fitok, "모든 글이 상자 폭 안에 든다")
    check(sum(1 for b in bx if b[0] == "outcome") == 5, "결과 상자 5 개(3 + 2)")

    print("(h) 갈래가 둘이 아니면 거부한다 (실패해야 하는 입력)")
    try:
        FG.make_split_flowchart(trunk, arms[:1], os.path.join(tempfile.mkdtemp(), "x.png"))
        check(False, "갈래 1 개를 받아들였다")
    except ValueError:
        check(True, "갈래 1 개 -> ValueError")

    print("(i) 제외를 줄기에서 한 번에 하면 갈래 상자 없이 그린다 (2026-10-02)")
    arms0 = [dict(a, box=None, exclude=None) for a in arms]
    g0 = FG.make_split_flowchart(trunk, arms0, os.path.join(tempfile.mkdtemp(), "x0.png"))
    check(not any(b[0] == "arm" for b in g0["boxes"]), "갈래 상자를 그리지 않는다")
    check(sum(1 for b in g0["boxes"] if b[0] == "final") == 2, "최종 상자 둘")
    try:
        FG.make_split_flowchart(trunk, [dict(arms0[0], exclude="x"), arms0[1]],
                                os.path.join(tempfile.mkdtemp(), "x1.png"))
        check(False, "갈래 상자 없이 갈래 제외를 받아들였다")
    except ValueError:
        check(True, "갈래 상자 없이 갈래 제외 -> ValueError")

    print()
    if FAILS:
        print("FAILED %d" % len(FAILS))
        return 1
    print("all passed")
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
