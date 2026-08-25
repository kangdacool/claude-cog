# -*- coding: utf-8 -*-
"""
deck_render_audit.py — PPTX «렌더 잘림» 검사기

deck_audit.py 가 XML 구조를, audit_font_sizes.py 가 글자 크기를 본다면,
이 스크립트는 **실제로 그려진 그림**을 본다.

왜 필요한가
  PowerPoint 는 표의 행 높이를 «내용에 맞춰 늘린다». 그래서 파일에 저장된
  좌표·크기를 아무리 검사해도 «표가 슬라이드 아래로 넘쳤다»를 잡을 수 없다.
  python-pptx 로 읽은 shape.top + shape.height 는 슬라이드 안에 들어 있는데
  화면에는 마지막 행이 잘려 나간다.

  2026-07-30 선행학습 6종에서 이 검사가 좌표 검사·XML 검사가 전부 통과시킨
  결함 12건(표 넘침 11 + 불릿 넘침 1)을 잡았다. 사람 눈으로도 놓쳤던 것이다.

원리
  PowerPoint COM 으로 PNG 를 내보낸 뒤, 각 장의 **아래 가장자리 띠**(기본 26px)
  에서 배경색이 아닌 «진한 픽셀»을 센다. 배경은 그 띠의 최빈색으로 잡는다.
  어두운 배경(구분장 등)은 건너뛴다.

사용법
  python deck_render_audit.py deck.pptx [deck2.pptx ...]
  python deck_render_audit.py deck.pptx --png-dir out/ --band 26 --hits 60
  python deck_render_audit.py --skip-render --png-dir out/     # 이미 뽑아둔 PNG만 검사
  python deck_render_audit.py deck.pptx --checkpoints cps.json # 하단 띠 + 임의 y좌표 다건 검사

종료 코드: 의심 슬라이드가 있으면 1, 없으면 0

pptx-editing 스킬의 `audit_text_fit.py`(PowerPoint COM으로 TextRange.BoundWidth/Height를 실측)가
더 정밀하지만 **표는 못 본다** — 표 셰이프는 shape 단위 TextFrame이 없어서(텍스트는 셀마다 있다)
그 스크립트의 측정 대상에서 아예 빠진다. 표 행높이 확장으로 인한 잘림은 여전히 이 스크립트가
유일하게 잡는다. 대체 관계가 아니라 상호보완 — 표가 있는 덱이면 둘 다 돌릴 것.

⚠️ 2026-08-14 확장 이유 — 기본 동작(슬라이드 맨 아래 띠만 검사)은 "표가 늘어나서 슬라이드
바닥을 넘었다"만 잡지, **포스터 중간에서 표가 늘어나 바로 아래 캡션과 겹치는 경우**는 못
잡는다 — 한 포스터에서 같은 빌드 안에 이 유형 결함이 3번 독립적으로 발생했는데(행
라벨 줄바꿈, 헤더 셀 줄바꿈, 제목 줄바꿈), 셋 다 슬라이드 맨 아래가 아니라 포스터 중간이라
기존 하단 띠 검사로는 하나도 안 걸렸다. 근본 수정은 `poster_kit.ptable()`이 셀 줄바꿈에서
실제 필요한 행높이를 계산하도록 고친 것(같은 날짜, poster_kit.py)이지만, 이 스크립트도
"하단 말고 아무 y좌표나 검사"를 지원해야 그 수정이 없는 프로젝트나 다른 종류의 겹침도
잡을 수 있다. `--checkpoints`로 `[{"label": str, "y_cm": float}, ...]` JSON을 주면, 슬라이드
크기 기준으로 그 y좌표 각각에서 얇은 띠를 검사한다 — 빌드 스크립트가 표/캡션을 그릴 때마다
이미 계산해 둔 y좌표를 그대로 흘려보내면 된다(예: `checkpoints.append({"label": "Table 3
캡션 시작", "y_cm": y})`). 체크포인트가 없으면 기존처럼 하단 띠만 본다 — 하위호환 유지.
"""
import argparse
import sys
import unicodedata
from collections import Counter
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BAND = 26          # 아래 가장자리 몇 px 를 볼지
HIT_MIN = 60       # 이 개수 이상 진한 픽셀이면 «잘림 의심»
DARK = 150         # 평균 밝기가 이보다 낮으면 진한 픽셀
BG_DARK = 120      # 배경 자체가 이보다 어두우면 구분장으로 보고 건너뛴다
WIDTH = 1600


def render(pptx, png_dir):
    """PowerPoint COM 으로 슬라이드를 PNG 로 내보낸다. 절대경로여야 한다."""
    import win32com.client
    pptx = Path(pptx).resolve()
    dst = Path(png_dir).resolve() / pptx.stem
    dst.mkdir(parents=True, exist_ok=True)
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(str(pptx), WithWindow=False)
    try:
        for i, slide in enumerate(pres.Slides, start=1):
            slide.Export(str(dst / f"{i:02d}.png"), "PNG", WIDTH,
                         int(WIDTH * pres.PageSetup.SlideHeight
                             / pres.PageSetup.SlideWidth))
    finally:
        pres.Close()
        app.Quit()
    return dst


def check_band_at(png, y_frac, band=BAND):
    """이미지 높이 비율 y_frac(0=맨 위, 1=맨 아래) 위치를 중심으로 한 얇은 띠에서 진한 픽셀
    수를 센다. 배경이 어두우면(구분장 등) None. `check()`는 y_frac=1.0(하단)인 특수 케이스."""
    from PIL import Image
    im = Image.open(png).convert("RGB")
    w, h = im.size
    y0 = max(0, min(h - band, round(y_frac * h - band / 2)))
    strip = im.crop((0, y0, w, y0 + band))
    px = list(strip.getdata())
    bg = Counter(px).most_common(1)[0][0]
    if sum(bg) / 3 < BG_DARK:
        return None
    return sum(1 for r, g, b in px if (r + g + b) / 3 < DARK)


def check(png, band=BAND, hits_min=HIT_MIN):
    """아래 가장자리에서 진한 픽셀 수. 배경이 어두우면 None."""
    from PIL import Image
    im = Image.open(png).convert("RGB")
    w, h = im.size
    strip = im.crop((0, h - band, w, h))
    px = list(strip.getdata())  # Pillow 14에서 get_flattened_data 로 바뀔 예정
    bg = Counter(px).most_common(1)[0][0]
    if sum(bg) / 3 < BG_DARK:
        return None
    return sum(1 for r, g, b in px if (r + g + b) / 3 < DARK)


def load_checkpoints(path):
    """{"slide_h_cm": float, "points": [{"label": str, "y_cm": float}, ...]} 읽기.
    slide_h_cm이 있어야 y_cm을 이미지 높이 비율로 바꿀 수 있다 — 빌드 스크립트가 이미
    알고 있는 값(포스터 PH 등)이므로 여기서 추측하지 않는다."""
    import json
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    slide_h = float(data["slide_h_cm"])
    return slide_h, data.get("points", [])


def audit_dir(d, band, hits_min, checkpoints=None):
    """checkpoints: (slide_h_cm, [{"label","y_cm"}, ...]) or None (하단 띠만 검사, 기존 동작)."""
    bad = []
    pngs = sorted(Path(d).glob("*.png"))
    for p in pngs:
        n = check(p, band, hits_min)
        if n is not None and n >= hits_min:
            bad.append(("bottom edge", p.stem, n))
        if checkpoints:
            slide_h, points = checkpoints
            for pt in points:
                y_frac = float(pt["y_cm"]) / slide_h
                n2 = check_band_at(p, y_frac, band)
                if n2 is not None and n2 >= hits_min:
                    bad.append((pt["label"], p.stem, n2))
    return len(pngs), bad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", nargs="*")
    ap.add_argument("--png-dir", default="_render")
    ap.add_argument("--band", type=int, default=BAND)
    ap.add_argument("--hits", type=int, default=HIT_MIN)
    ap.add_argument("--skip-render", action="store_true")
    ap.add_argument("--checkpoints", default=None,
                     help='JSON: {"slide_h_cm": .., "points": [{"label":..,"y_cm":..}, ...]} - '
                          "check arbitrary y-positions (e.g. right after each table), not just "
                          "the slide's bottom edge. Omit for the original bottom-only behavior.")
    a = ap.parse_args()

    dirs = []
    if a.skip_render:
        dirs = [d for d in sorted(Path(a.png_dir).iterdir()) if d.is_dir()]
    else:
        if not a.pptx:
            print(__doc__)
            return 1
        for f in a.pptx:
            dirs.append(render(f, a.png_dir))

    checkpoints = load_checkpoints(a.checkpoints) if a.checkpoints else None
    label = f"아래 가장자리 {a.band}px" + (f" + 체크포인트 {len(checkpoints[1])}개" if checkpoints else "")
    print(f"=== 렌더 잘림 검사 ({label})")
    total = 0
    for d in dirs:
        n, bad = audit_dir(d, a.band, a.hits, checkpoints)
        total += len(bad)
        name = unicodedata.normalize("NFC", d.name)[:30]
        print(f"  {name:32s} 검사 {n:3d}장  의심 {len(bad)}")
        for where, s, k in bad:
            print(f"      slide {s}  [{where}]  진한 픽셀 {k}")
    print(f"\n=== TOTAL 의심 {total}")
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
