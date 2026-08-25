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
    return w_cap, out


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
    if autofit:
        # Two passes: each column's fit depends on the other's width through
        # TOTAL_W, so solve one, then re-solve the other against the result.
        left_w, boxes = _fit_column([s["box"] for s in steps],
                                    left_w, right_w, width_in, font_family, font_size)
        right_w, excls = _fit_column([s.get("exclude") for s in steps],
                                     right_w, left_w, width_in, font_family, font_size)
        left_w, boxes = _fit_column(boxes, left_w, right_w, width_in,
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

    # Build layout positions (top-down)
    rows = []
    y = 0  # top of figure (will flip at the end)
    for i, step in enumerate(steps):
        lh = box_h(step["box"])
        rh = box_h(step["exclude"]) if step["exclude"] else 0

        row = {
            "left_top": y,
            "left_h": lh,
            "left_bot": y + lh,
            "right_h": rh,
        }

        # If there's an exclusion, the horizontal arrow branches from the
        # vertical connector between this box and the next.
        # We need gap space for the connector + branch.
        rows.append(row)
        y += lh
        if i < len(steps) - 1:
            y += GAP_V

    total_h = y
    # Add margin (kept tight — see pad_inches on save)
    margin = 0.10
    total_h += 2 * margin

    # Compute figure height
    height_in = width_in * (total_h / TOTAL_W)
    height_in = max(height_in, 3.5)

    # ── Create figure ──
    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=dpi)
    ax.set_xlim(0, TOTAL_W)
    ax.set_ylim(total_h, 0)  # inverted: 0 at top
    ax.axis("off")
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)

    arrow_kw = dict(arrowstyle="->,head_width=0.15,head_length=0.15",
                    color="black", lw=1.2)

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
    offset = margin

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
                        color="black", lw=1.2, zorder=1)
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


if __name__ == "__main__":
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
