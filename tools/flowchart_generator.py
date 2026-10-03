"""
STROBE-style participant selection flowchart generator.

Produces a clean black-and-white flowchart matching the Kang Lab standard:
  - Left column: retained population boxes (vertical flow)
  - Right column: exclusion reason boxes (horizontal arrows branching from vertical connector)
  - Black borders, white fill, centered text, no colors

Usage:
    from flowchart_generator import make_flowchart

    steps = [
        {"box": "All patients ID\n(N=91,152)", "exclude": None},
        {"box": "Patients with baseline US\n(N=88,925)",
         "exclude": "Exclude duplicates (N=23)\nWithout baseline US (N=2,204)"},
        ...
    ]
    make_flowchart(steps, "flowchart.png", dpi=300, width_in=6.5)
"""

import textwrap

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.font_manager import FontProperties
from matplotlib.textpath import TextPath

# Layout constants shared by the fitter and the drawing pass.
MARGIN_L = 0.75          # left margin (data coords)
GUTTER   = 1.55          # gap between the two columns
MARGIN_R = 0.70          # right margin
# Clearance inside a box, inches (both sides together). 0.10 was enough for the
# text to fit but left it visually touching the border — passing "it fits" is not
# the same as passing "it looks right" (CORE §2⑦), and that only showed on render.
PAD_IN   = 0.26
# A column may not take more than this share of the figure width, or the other
# column is squeezed out. Past it we wrap the text instead of widening further.
MAX_SHARE = 0.60


def _line_pts(line, font_family, font_size):
    """Rendered width of one line, in points. Canvas-free (TextPath), so this can
    run before the figure exists — which it must, since the figure geometry is
    what we are solving for."""
    if not line:
        return 0.0
    fp = FontProperties(family=font_family, size=font_size)
    return float(TextPath((0, 0), line, prop=fp).get_extents().width)


def _needed_w(need_in, other_w, width_in):
    """Box width (data coords) whose PHYSICAL width is at least `need_in` inches.

    The figure is a fixed `width_in` inches spanning TOTAL_W data units, so a box
    of width w is (w / TOTAL_W) * width_in inches across. Widening w also widens
    TOTAL_W, so the gain is sub-linear but real. Solving

        w / (K + w) * width_in >= need_in ,  K = MARGIN_L + GUTTER + other_w + MARGIN_R

    gives w >= need_in * K / (width_in - need_in). Returns None when the text
    cannot fit at any width (need_in >= width_in) — the caller then wraps."""
    K = MARGIN_L + GUTTER + other_w + MARGIN_R
    if need_in >= width_in:
        return None
    return need_in * K / (width_in - need_in)


def _fit_column(texts, base_w, other_w, width_in, font_family, font_size):
    """Widen the column to hold its widest line; past MAX_SHARE, wrap instead.

    Returns (width, [possibly wrapped texts]). Text that already fits at
    `base_w` returns `base_w` unchanged, so existing figures never reflow."""
    w_cap = _needed_w(MAX_SHARE * width_in, other_w, width_in) or base_w
    w_cap = max(w_cap, base_w)

    def widest_in(ts):
        return max((_line_pts(ln, font_family, font_size)
                    for t in ts if t for ln in t.split("\n")), default=0.0)

    need_in = widest_in(texts) / 72.0 + PAD_IN
    w = _needed_w(need_in, other_w, width_in)
    if w is not None and w <= w_cap:
        return max(base_w, w), list(texts)

    # Too wide to solve by widening: cap the box and wrap to what fits there.
    room_in = (w_cap / (MARGIN_L + w_cap + GUTTER + other_w + MARGIN_R)) * width_in
    return w_cap, _wrap_to_room(texts, room_in, font_family, font_size)


def _wrap_to_room(texts, room_in, font_family, font_size):
    """Break each line to what fits in `room_in` inches of box interior.

    Split out of _fit_column so the final-geometry check can reuse it: the
    fitting loop can leave text that does not fit, and only a check against the
    FINAL widths catches that (see _enforce_fit)."""
    room_pts = max((room_in - PAD_IN) * 72.0, 1.0)
    out = []
    for t in texts:
        if not t:
            out.append(t)
            continue
        lines = []
        for ln in t.split("\n"):
            if _line_pts(ln, font_family, font_size) <= room_pts:
                lines.append(ln)
                continue
            # Estimate characters per line from this line's own average glyph
            # width, then let textwrap break on word boundaries.
            avg = _line_pts(ln, font_family, font_size) / max(len(ln), 1)
            ncol = max(int(room_pts / avg), 8)
            lines.extend(textwrap.wrap(ln, width=ncol) or [ln])
        out.append("\n".join(lines))
    return out


def _enforce_fit(boxes, excls, left_w, right_w, width_in, font_family, font_size):
    """Last line of defence: wrap anything that still overruns the FINAL geometry.

    Why this is needed even though _fit_column already fits each column.
    Widening either column enlarges TOTAL_W, and MAX_SHARE is a share of the
    figure — so the other column's cap grows too. With both columns hungry the
    caps never bind, the loop keeps inflating both, and it exits (on iteration
    count, not convergence) with the column fitted FIRST measured against a
    narrower partner than it ends up with. That column then renders outside its
    own border, and nothing reports it: the loop believed both columns fit.

    Measured 2026-08-31 (psy STROBE chart, width_in=6.10): the left column
    needed 3.24 in and the right 3.46 in — 6.70 in of text for 6.10 in of
    figure. No cap ever bound; the 52-character left label ran across the frame.

    Checking against the final widths is the only place the truth is known, so
    the fix belongs here rather than inside the loop. Text that fits is returned
    untouched, so figures that render correctly today do not reflow."""
    total_w = MARGIN_L + left_w + GUTTER + right_w + MARGIN_R
    left_room = (left_w / total_w) * width_in
    right_room = (right_w / total_w) * width_in
    return (_wrap_to_room(boxes, left_room, font_family, font_size),
            _wrap_to_room(excls, right_room, font_family, font_size))


def make_flowchart(steps, outpath, dpi=300, width_in=6.5,
                   font_size=9, font_family="Arial",
                   left_w=4.5, right_w=4.0, autofit=True):
    """
    Parameters
    ----------
    steps : list of dict
        Each dict has:
          "box"     : str  — text for the retained-population box (use \\n for lines)
          "exclude" : str or None — text for the exclusion box (None = no exclusion)
    outpath : str — output PNG path.
    dpi : int — resolution.
    width_in : float — figure width in inches (native = display size).
    font_size : int — base font size.
    font_family : str — font family.
    left_w, right_w : float — box widths in data coords (defaults = original layout).
        These are FLOORS: with autofit on, a column is widened past its default
        when a line would otherwise run into the box border. Do NOT shrink
        font_size instead — see CORE §2① (chronic too-small-text bias).
        The columns are re-centred automatically so the figure stays balanced.
    autofit : bool — widen a column to hold its widest line, and once widening
        would take the column past MAX_SHARE of the figure, wrap the text to the
        next line instead. Text that already fits is untouched and the widths stay
        at their defaults, so figures that render correctly today do not reflow.
        Set False to reproduce the pre-2026-08-23 behaviour exactly (text that is
        too long then overflows the border silently, which is what this fixes).
    """

    # ── Layout parameters (data coords) ──
    # Positions are derived from the widths so that the original geometry
    # (LEFT_CX 3.0, RIGHT_CX 8.8, TOTAL_W 11.5) is reproduced exactly at the
    # default widths — widening a box must not silently reflow existing figures.
    steps = [dict(s) for s in steps]          # never mutate the caller's list

    # ── Lab norm check ──────────────────────────────────────────────────
    # Measured from the pipelines that actually reached submission:
    #   a submitted cohort paper  3 boxes   a manuscript in progress  5 boxes
    # Each box there is a POPULATION with 2-3 lines of description ending in
    # the count — not one criterion per box. A chart that enumerates every
    # filter reads as a debug trace, and reviewers see it that way.
    #
    # psy 2026-08-31: this chart was built at 8 boxes, one reason each, after
    # surveying *published* flowcharts instead of opening the lab's own. The
    # researcher had pointed at the lab's ("우리 다른 파이프라인에서는 잘 만들던데")
    # and the survey went the opposite way. Warn at the point of creation so
    # the next pipeline is asked the question before it ships.
    if len(steps) > 6:
        print(f"  [flowchart] 상자 {len(steps)}개 — 투고까지 간 이 연구실 도표는 3~5개다.\n"
              f"              «기준 하나»가 아니라 «인구집단» 단위로 묶고, 제외 사유는\n"
              f"              한 곁상자에 여러 줄로 싣는지 확인할 것. "
              f"([[flowchart-generator]] 의 실측표)")
    _thin = [i for i, s in enumerate(steps)
             if s.get("exclude") and "\n" not in str(s["exclude"])]
    if len(steps) > 4 and len(_thin) >= len(steps) - 1:
        print(f"  [flowchart] 곁상자가 전부 한 줄이다({len(_thin)}개) — 사유를 단계마다\n"
              f"              쪼개고 있지 않은지 볼 것. 투고본은 한 상자에 2~3개를 묶는다.")
    if autofit:
        # Each column's fit depends on the other's width through TOTAL_W, so
        # iterate to a fixed point instead of doing a fixed number of passes.
        #
        # The old code ran left -> right -> left. Ending on `left` meant the last
        # change to left_w was never re-checked against the right column: the
        # right box had been wrapped to the room available BEFORE left grew, so
        # its text then ran outside its own border. Only visible on render.
        # (psy 2026-08-31: a 45-character exclusion line overflowed the frame.)
        #
        # Re-fit from the ORIGINAL strings every round — feeding already-wrapped
        # text back in would accumulate line breaks and give worse break points.
        src_box = [s["box"] for s in steps]
        src_exc = [s.get("exclude") for s in steps]
        boxes, excls = list(src_box), list(src_exc)
        for _ in range(6):
            prev = (left_w, right_w)
            left_w, boxes = _fit_column(src_box, left_w, right_w, width_in,
                                        font_family, font_size)
            right_w, excls = _fit_column(src_exc, right_w, left_w, width_in,
                                         font_family, font_size)
            if (left_w, right_w) == prev:
                break
        # The loop above can exit on iteration count rather than convergence.
        # Re-check both columns against the widths they actually ended up with.
        # Feed the ORIGINAL strings, never the already-wrapped ones: re-wrapping
        # wrapped text breaks each existing line separately and strands two-word
        # lines ("with a", "recorded"). Same reason the loop above re-fits from
        # src_box/src_exc. (psy 2026-08-31: the fix added the day before produced
        # exactly the raggedness it was meant to prevent.)
        boxes, excls = _enforce_fit(src_box, src_exc, left_w, right_w, width_in,
                                    font_family, font_size)
        for s, b, e in zip(steps, boxes, excls):
            s["box"] = b
            if s.get("exclude"):
                s["exclude"] = e

    LEFT_W     = left_w
    RIGHT_W    = right_w
    LEFT_CX    = MARGIN_L + LEFT_W / 2
    RIGHT_CX   = MARGIN_L + LEFT_W + GUTTER + RIGHT_W / 2
    TOTAL_W    = MARGIN_L + LEFT_W + GUTTER + RIGHT_W + MARGIN_R
    LINE_H     = 0.40         # height per text line
    PAD_V      = 0.2          # vertical padding inside box (total)
    GAP_V      = 0.5          # vertical gap between rows (connector space)

    # ── Compute box heights ──
    def nlines(text):
        return text.count("\n") + 1 if text else 0

    def box_h(text):
        return nlines(text) * LINE_H + PAD_V

    lefts = [box_h(s["box"]) for s in steps]
    rights = [box_h(s["exclude"]) if s.get("exclude") else 0.0 for s in steps]

    # ── Row gaps ──
    # An exclusion box is centred on the connector between step i and step i+1
    # (see the drawing pass below), and its height is unbounded — it grows with
    # the number of reasons listed. With a fixed GAP_V, two tall exclusion boxes
    # on consecutive connectors OVERLAP each other, which only shows on render.
    # (psy 2026-08-27: a 6-line eligibility box ran into the next exclusion box.)
    #
    # So grow a gap only when two neighbouring exclusion boxes would actually
    # collide. Charts whose exclusion boxes already clear each other keep
    # GAP_V exactly, so existing figures do not reflow.
    CLEAR_V = 0.18                       # minimum white space between two right boxes
    gaps = [GAP_V] * max(len(steps) - 1, 0)

    # A connector must be at least as long as the exclusion box it carries.
    # Otherwise a six-line box hangs off a short arrow while a one-line box sits
    # on a long one, and the chart reads as unbalanced even though nothing
    # overlaps. (psy 2026-08-31, researcher: "텍스트가 저렇게 길면 화살표 길이도
    # 그만큼 길게 해야 균형이 맞지 않나 — 오히려 배제과정이 한 줄인 쪽이 더 길다.")
    # Charts whose boxes are all shorter than GAP_V keep GAP_V exactly.
    for i in range(len(gaps)):
        if rights[i + 1]:
            gaps[i] = max(gaps[i], rights[i + 1] + CLEAR_V)

    for i in range(len(gaps) - 1):
        # Right boxes hang on connector i (explains step i+1) and connector i+1
        # (explains step i+2). Centre-to-centre distance between them:
        #   gaps[i]/2 + lefts[i+1] + gaps[i+1]/2
        if not (rights[i + 1] and rights[i + 2]):
            continue
        need = (rights[i + 1] + rights[i + 2]) / 2 + CLEAR_V
        have = gaps[i] / 2 + lefts[i + 1] + gaps[i + 1] / 2
        if have < need:
            gaps[i + 1] += 2 * (need - have)

    # Build layout positions (top-down)
    rows = []
    y = 0  # top of figure (will flip at the end)
    for i, step in enumerate(steps):
        rows.append({
            "left_top": y,
            "left_h": lefts[i],
            "left_bot": y + lefts[i],
            "right_h": rights[i],
        })
        y += lefts[i]
        if i < len(steps) - 1:
            y += gaps[i]

    total_h = y
    # Add margin (kept tight — see pad_inches on save)
    margin = 0.10
    # A tall exclusion box centred on the FIRST connector can reach above the top
    # of the first left box (and the last one below the bottom). Widen the margin
    # by however much it actually overhangs, so nothing is cropped at save time.
    over_top = over_bot = 0.0
    for i in range(len(steps) - 1):
        if not rights[i + 1]:
            continue
        mid = rows[i]["left_bot"] + gaps[i] / 2
        over_top = max(over_top, (rights[i + 1] / 2) - mid)
        over_bot = max(over_bot, mid + rights[i + 1] / 2 - y)
    margin_top = margin + max(0.0, over_top)
    margin_bot = margin + max(0.0, over_bot)
    total_h += margin_top + margin_bot

    # Compute figure height.
    #
    # ⚠️ Height is derived from the WIDTH ratio, so widening a column grows
    # TOTAL_W and therefore SHRINKS the vertical scale. With a few rows that is
    # invisible; with many rows the text collides with its own box border and
    # with the row below. (psy 2026-08-31: an 8-stage chart came out the same
    # height as the 4-stage one it replaced, and every second box overflowed.)
    #
    # So also floor the height at what the rows physically need. LINE_H is in
    # data units; MIN_UNIT_IN is the inches each unit must get for a 9pt line to
    # sit inside its box. Charts that already had enough room are untouched,
    # so existing figures do not reflow.
    MIN_UNIT_IN = 0.40
    height_in = width_in * (total_h / TOTAL_W)
    height_in = max(height_in, 3.5, total_h * MIN_UNIT_IN)

    # ── Create figure ──
    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=dpi)
    ax.set_xlim(0, TOTAL_W)
    ax.set_ylim(total_h, 0)  # inverted: 0 at top
    ax.axis("off")
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)

    ## shrinkA/B=0: matplotlib 화살표는 기본으로 양 끝을 2 pt 띄워 그려, 꺾이고 갈라지는 자리마다 틈이 났다
    ## (2026-10-02 연구자: 「화살표가 꺾이거나 파생되는 곳에서 빈칸이 생긴다 · reverted 화살표도 끊어졌다」).
    arrow_kw = dict(arrowstyle="->,head_width=0.15,head_length=0.15",
                    color="black", lw=1.2, shrinkA=0, shrinkB=0)

    def draw_box(cx, top, w, h, text):
        x0 = cx - w / 2
        rect = mpatches.FancyBboxPatch(
            (x0, top), w, h,
            boxstyle="square,pad=0",
            facecolor="white", edgecolor="black", linewidth=1.0,
            zorder=2
        )
        ax.add_patch(rect)
        ax.text(cx, top + h / 2, text,
                ha="center", va="center",
                fontsize=font_size, fontfamily=font_family,
                zorder=3, linespacing=1.4)

    def arrow_v(x, y1, y2):
        ax.annotate("", xy=(x, y2), xytext=(x, y1),
                    arrowprops=arrow_kw, zorder=1)

    def arrow_h(x1, x2, y):
        ax.annotate("", xy=(x2, y), xytext=(x1, y),
                    arrowprops=arrow_kw, zorder=1)

    # ── Draw ──
    # Exclusion semantics: step[i]["exclude"] explains how step[i-1] → step[i].
    # So the horizontal arrow branches from the connector ABOVE step[i].
    offset = margin_top

    for i, step in enumerate(steps):
        r = rows[i]
        lt = r["left_top"] + offset
        lh = r["left_h"]
        lb = lt + lh

        # Draw left box
        draw_box(LEFT_CX, lt, LEFT_W, lh, step["box"])

        # Vertical connector to next box
        if i < len(steps) - 1:
            next_step = steps[i + 1]
            next_r = rows[i + 1]
            next_lt = next_r["left_top"] + offset
            mid_y = (lb + next_lt) / 2

            if next_step["exclude"]:
                rh = box_h(next_step["exclude"])
                # Vertical line: box bottom -> junction
                ax.plot([LEFT_CX, LEFT_CX], [lb, mid_y],
                        color="black", lw=1.2, zorder=1, solid_capstyle="projecting")
                # Junction -> next box (with arrowhead)
                arrow_v(LEFT_CX, mid_y, next_lt)

                # Horizontal arrow from junction to right box
                right_left_edge = RIGHT_CX - RIGHT_W / 2
                arrow_h(LEFT_CX, right_left_edge, mid_y)

                # Right exclusion box centered on junction y
                right_top = mid_y - rh / 2
                draw_box(RIGHT_CX, right_top, RIGHT_W, rh, next_step["exclude"])
            else:
                # Simple vertical arrow
                arrow_v(LEFT_CX, lb, next_lt)

    # ── Save ──
    fig.savefig(outpath, dpi=dpi, bbox_inches="tight",
                facecolor="white", edgecolor="none", pad_inches=0.02)  # tight margins
    plt.close(fig)
    print(f"Saved: {outpath}")


def make_split_flowchart(trunk, arms, outpath, dpi=300, font_size=10,
                         font_family="Arial", emph_final=True):
    """Two-arm STROBE flowchart: a shared spine that splits into two cohorts.

    Added 2026-09-30 (한 코호트 결과 덱). The single-column make_flowchart cannot
    draw a split, so a two-arm chart had been hand-drawn with slide shapes and
    looked nothing like the lab's charts. Layout follows McGirr 2022 Neurology
    Figure 1 and Gifford 2014 Alzheimers Dement Figure 1 (the published designs
    for "normal at baseline" and "MCI at baseline" cohorts side by side), in this
    lab's black-and-white style:

        trunk box ── exclusion (right)
            │
        trunk box
       ┌────┴────┐
    excl ─ arm     arm ─ excl        exclusions sit on the OUTER side of each arm
          │         │
        final     final              analytic sample (thicker border)
       ┌──┼──┐    ┌─┴─┐
       o  o  o    o   o              outcomes at follow-up, n (%)

    Parameters
    ----------
    trunk : list of {"box": str, "exclude": str|None}, same semantics as
        make_flowchart: step[i]["exclude"] explains step[i-1] -> step[i].
    arms : list of exactly two dicts, left then right:
        {"box": str|None, "exclude": str|None, "final": str, "outcomes": [str, ...]}
        "exclude" explains box -> final. "outcomes" may be empty.
        "box": None on BOTH arms -> the split lands directly on the final boxes. Use it when every
        exclusion is done once on the trunk (2026-10-02 한 코호트 결과 덱: arm-side exclusion boxes made the
        chart 8.3 in wide; one trunk exclusion box gives a narrow chart that fits a manuscript page).
    Returns a geometry dict {"boxes": [(kind, x0, y0, x1, y1, text)], "W", "H"}
    in inches, so tests can check fit and overlap without looking at pixels.

    Sizing is in inches and boxes are sized from the rendered text, so the font
    never has to shrink (CORE §2①). Exclusion text is left-aligned, population
    text centred (McGirr style). Use "n = 1,234" with spaces.
    """
    if len(arms) != 2:
        raise ValueError("make_split_flowchart draws exactly two arms (got %d)" % len(arms))
    fs = font_size
    LH = fs * 1.42 / 72.0            # line height, inches
    PV, PH = 0.08, 0.14              # box padding (each side), inches
    GAP_H, M = 0.40, 0.08            # gap between arms, outer margin
    EXC_GAP = 0.22                   # horizontal gap: spine box edge -> exclusion box

    def tw(t):                       # text width, inches
        return max((_line_pts(ln, font_family, fs) for ln in t.split("\n")), default=0.0) / 72.0

    def bw(t, minw=0.9):
        return max(tw(t) + 2 * PH, minw)

    def bh(t):
        return (t.count("\n") + 1) * LH + 2 * PV

    # ── widths ──
    A = []
    for a in arms:
        ow = [bw(o, 0.7) for o in a.get("outcomes", [])]
        ## 결과 상자는 한 갈래 안에서 «같은 폭» — 폭이 다르면 가운데 상자가 갈래 중심에서 비켜나, 갈래 세로선과
        ## 가운데 화살표가 나란히 두 줄로 보였다(2026-10-02 연구자: 「reverted 화살표도 끊어졌어」).
        ow = [max(ow)] * len(ow) if ow else ow
        out_total = sum(ow) + 0.12 * max(len(ow) - 1, 0)
        core = max(bw(a["box"]) if a.get("box") else 0.0, bw(a["final"]))
        half = max(core / 2, out_total / 2)
        ew = bw(a["exclude"]) if a.get("exclude") else 0.0
        A.append(dict(core=core, half=half, ew=ew, ow=ow, out_total=out_total))
    # outer extent: exclusion sits outside the spine boxes
    outL = max(A[0]["half"], A[0]["core"] / 2 + (EXC_GAP + A[0]["ew"] if A[0]["ew"] else 0))
    outR = max(A[1]["half"], A[1]["core"] / 2 + (EXC_GAP + A[1]["ew"] if A[1]["ew"] else 0))
    xL = M + outL
    xR = xL + A[0]["half"] + GAP_H + A[1]["half"]
    W = xR + outR + M
    xc = (xL + xR) / 2
    tcore = max(bw(s["box"]) for s in trunk)
    tew = max((bw(s["exclude"]) for s in trunk if s.get("exclude")), default=0.0)
    if tew:
        W = max(W, xc + tcore / 2 + EXC_GAP + tew + M)

    geo = []
    fig_items = []                   # deferred drawing: (kind, args)

    def box(kind, cx, top, w, text, align="center", lw=1.0, h=None):
        h = bh(text) if h is None else h
        geo.append((kind, cx - w / 2, top, cx + w / 2, top + h, text))
        fig_items.append(("box", (cx - w / 2, top, w, h, text, align, lw)))
        return top + h

    def line(x1, y1, x2, y2, head=False):
        fig_items.append(("line", (x1, y1, x2, y2, head)))

    # ── trunk ──
    y = M
    for i, s in enumerate(trunk):
        bot = box("trunk", xc, y, tcore, s["box"])
        if i < len(trunk) - 1:
            nxt = trunk[i + 1]
            if nxt.get("exclude"):
                eh = bh(nxt["exclude"])
                gap = max(0.34, eh + 0.16)
                mid = bot + gap / 2
                line(xc, bot, xc, bot + gap, head=True)
                ex0 = xc + tcore / 2 + EXC_GAP
                line(xc, mid, ex0, mid, head=True)
                box("exclude", ex0 + tew / 2, mid - eh / 2, tew, nxt["exclude"], align="left")
            else:
                gap = 0.30
                line(xc, bot, xc, bot + gap, head=True)
            y = bot + gap
        else:
            y = bot
    # ── split ──
    ybar = y + 0.16
    top_arm = ybar + 0.20
    line(xc, y, xc, ybar)
    line(xL, ybar, xR, ybar)
    for x in (xL, xR):
        line(x, ybar, x, top_arm, head=True)
    # ── arms: same row tops so the two cohorts read as parallel ──
    no_arm_box = not any(a.get("box") for a in arms)
    if no_arm_box and any(a.get("exclude") for a in arms):
        raise ValueError("an arm exclusion needs an arm box to come from")
    arm_h = 0.0 if no_arm_box else max(bh(a["box"]) for a in arms if a.get("box"))
    ex_h = max((bh(a["exclude"]) for a in arms if a.get("exclude")), default=0.0)
    gap_a = 0.0 if no_arm_box else max(0.34, ex_h + 0.16)
    fin_top = top_arm + arm_h + gap_a
    fin_h = max(bh(a["final"]) for a in arms)
    for k, (a, g, x) in enumerate(zip(arms, A, (xL, xR))):
        if not no_arm_box:
            box("arm", x, top_arm, g["core"], a["box"])
            line(x, top_arm + bh(a["box"]), x, fin_top, head=True)
        if a.get("exclude"):
            mid = top_arm + arm_h + gap_a / 2
            eh = bh(a["exclude"])
            if k == 0:   # left arm: exclusion to the left
                ex1 = x - g["core"] / 2 - EXC_GAP
                line(x, mid, ex1, mid, head=True)
                box("exclude", ex1 - g["ew"] / 2, mid - eh / 2, g["ew"], a["exclude"], align="left")
            else:
                ex0 = x + g["core"] / 2 + EXC_GAP
                line(x, mid, ex0, mid, head=True)
                box("exclude", ex0 + g["ew"] / 2, mid - eh / 2, g["ew"], a["exclude"], align="left")
        box("final", x, fin_top, g["core"], a["final"], lw=1.8 if emph_final else 1.0)
    # ── outcomes ──
    H = fin_top + fin_h
    for a, g, x in zip(arms, A, (xL, xR)):
        outs = a.get("outcomes", [])
        if not outs:
            continue
        yb = fin_top + fin_h + 0.16
        yo = yb + 0.20
        cx, x0 = [], x - g["out_total"] / 2
        for w_ in g["ow"]:
            cx.append(x0 + w_ / 2)
            x0 += w_ + 0.12
        line(x, fin_top + bh(a["final"]), x, yb)
        line(min(cx), yb, max(cx), yb)
        for c, w_, o in zip(cx, g["ow"], outs):
            line(c, yb, c, yo, head=True)
            H = max(H, box("outcome", c, yo, w_, o, h=max(bh(t) for t in outs)))
    H += M

    # ── render ──
    fig, ax = plt.subplots(figsize=(W, H), dpi=dpi)
    ax.set_xlim(0, W); ax.set_ylim(H, 0); ax.axis("off")
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    akw = dict(arrowstyle="->,head_width=0.12,head_length=0.14", color="black", lw=1.1,
               shrinkA=0, shrinkB=0)          # 끝을 띄우지 않는다 — 위 make_flowchart 의 arrow_kw 주석
    for kind, args in fig_items:
        if kind == "box":
            x0, top, w, h, text, align, lw = args
            ax.add_patch(mpatches.Rectangle((x0, top), w, h, facecolor="white",
                                            edgecolor="black", linewidth=lw, zorder=2))
            tx = x0 + PH if align == "left" else x0 + w / 2
            ax.text(tx, top + h / 2, text, ha=align, va="center", fontsize=fs,
                    fontfamily=font_family, zorder=3, linespacing=1.3)
        else:
            x1, y1, x2, y2, head = args
            if head:
                ax.annotate("", xy=(x2, y2), xytext=(x1, y1), arrowprops=akw, zorder=1)
            else:
                ax.plot([x1, x2], [y1, y2], color="black", lw=1.1, zorder=1,
                        solid_capstyle="projecting")   # 꺾이는 모서리를 메운다
    fig.savefig(outpath, dpi=dpi, facecolor="white", edgecolor="none")
    plt.close(fig)
    print(f"Saved: {outpath}")
    return {"boxes": geo, "W": W, "H": H}


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
    # exclude on each step explains how the PREVIOUS step's N became this step's N
    steps = [
        {"box": "All patients ID\n(N=91,152)",
         "exclude": None},
        {"box": "Patients with baseline abdomen US,\n2003 October to 2015 December.\n(N=88,925)",
         "exclude": "Exclude duplicates (N=23)\nPatients without baseline US\n(N=2,204)"},
        {"box": "Patients with baseline abdomen US\nand clinical data\n(N=78,676)",
         "exclude": "Without baseline clinical data\n(N=10,249)"},
        {"box": "Cyst negative at baseline, and with any\nfollow-up abdomen US,\n2003 December to 2024 October.\n(N=41,179)",
         "exclude": "Cyst positive at baseline (N=6,717)\nWithout any follow-up US (N=30,658)"},
    ]
    make_flowchart(steps, "presentation/flowchart_toy_test.png")
