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

    print()
    if FAILS:
        print("FAILED %d" % len(FAILS))
        return 1
    print("all passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
