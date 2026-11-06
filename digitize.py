"""Read the data back out of a rendered 2D line plot (Origin/Excel style: black frame,
outward ticks, coloured series, boxed legend) and write a `line-plot` spec for specs.py.

  python digitize.py plot.png --series chest=000000 armpit=FF0000 ... \
      --xticks 0,1000 --yticks 0,5 [--tol 0.006]

--xticks/--yticks give the values of the first and last major tick, read off the picture.
"""
import argparse

import numpy as np
from PIL import Image


def runs(mask_1d):
    """(start, end) of consecutive True runs."""
    d = np.diff(np.concatenate([[0], mask_1d.astype(np.int8), [0]]))
    return list(zip(np.where(d == 1)[0], np.where(d == -1)[0]))


def frame(dark):
    h, w = dark.shape
    rows = [y for y in range(h) if max((e - s for s, e in runs(dark[y])), default=0) > 0.4 * w]
    cols = [x for x in range(w) if max((e - s for s, e in runs(dark[:, x])), default=0) > 0.4 * h]
    return min(cols), min(rows), max(cols), max(rows)


def ticks(strip, axis_len):
    """Centres and lengths of tick marks in a strip next to an axis (strip: rows=along axis)."""
    out = []
    for s, e in runs(strip.any(axis=1)):
        if e - s <= 0.02 * axis_len:
            out.append(((s + e - 1) / 2, int(strip[s:e].sum(axis=1).max())))
    if not out:
        return []
    longest = max(n for _, n in out)
    return [c for c, n in out if n >= 0.7 * longest]


def rdp(points, tol):
    """Ramer-Douglas-Peucker simplification of an (n, 2) array."""
    if len(points) < 3:
        return points
    a, b = points[0], points[-1]
    ab, ap = b - a, points - a
    d = np.abs(ab[0] * ap[:, 1] - ab[1] * ap[:, 0]) / (np.linalg.norm(ab) or 1)
    i = int(d.argmax())
    if d[i] <= tol:
        return np.array([a, b])
    return np.vstack([rdp(points[: i + 1], tol)[:-1], rdp(points[i:], tol)])


def digitize(path, series, xticks, yticks, tol=0.004, work_width=1400):
    img = Image.open(path).convert("RGB")
    W, H = img.size
    k = work_width / W
    small = np.asarray(img.resize((work_width, round(H * k)), Image.BOX)).astype(int)
    dark = small.max(axis=2) < 90
    l, t, r, b = frame(dark)
    xt = ticks(dark[b + 1: b + 25, l:r + 1].T, r - l)
    yt = ticks(dark[t:b + 1, max(0, l - 25): l], b - t)
    x0p, x1p = l + xt[0], l + xt[-1]
    y0p, y1p = t + yt[-1], t + yt[0]  # bottom tick = first value
    px2x = lambda px: xticks[0] + (px - x0p) * (xticks[1] - xticks[0]) / (x1p - x0p)
    px2y = lambda py: yticks[0] + (py - y0p) * (yticks[1] - yticks[0]) / (y1p - y0p)
    inner = np.zeros(dark.shape, bool)
    inner[t + 3: b - 2, l + 3: r - 2] = True
    # legend: the largest box of dark lines inside the frame
    lg = None
    hor = [(y, s, e) for y in range(t + 3, b - 2) for s, e in runs(dark[y, l + 3: r - 2]) if e - s > 0.12 * (r - l)]
    if hor:
        ys = sorted({y for y, _, _ in hor})
        s, e = hor[0][1] + l + 3, hor[0][2] + l + 3
        lg = (s, ys[0], e, max(y for y, s2, _ in hor if abs(s2 + l + 3 - s) < 5))
        inner[lg[1] - 2: lg[3] + 3, lg[0] - 2: lg[2] + 3] = False
    out = []
    for name, hexc in series:
        rgb = np.array([int(hexc[i:i + 2], 16) for i in (0, 2, 4)])
        m = (np.abs(small - rgb).sum(axis=2) < 90) & inner
        xs, ys = [], []
        for x in range(l + 3, r - 2):
            col = np.where(m[:, x])[0]
            if len(col):
                xs.append(px2x(x))
                ys.append(px2y(col.mean()))
        # simplify in axis-fraction units so tol means "fraction of the plot height/width"
        sx, sy = (xticks[1] - xticks[0]) or 1, (yticks[1] - yticks[0]) or 1
        pts = rdp(np.column_stack([np.array(xs) / sx, np.array(ys) / sy]), tol) * [sx, sy] if xs else np.zeros((0, 2))
        out.append({"name": name, "color": hexc,
                    "xy": " ".join(f"{x:.0f},{y:.2f}" for x, y in pts)})
    to_orig = lambda v: int(round(v / k))
    # text heights below the x axis: first band = tick labels, second = axis title
    bands = runs(dark[b + 25:, l:r + 1].any(axis=1))
    cap = lambda band: (band[1] - band[0]) / k
    fonts = {"ticks": round(cap(bands[0]) / 0.72) if bands else 40,
             "label": round(cap(bands[1]) / 0.72) if len(bands) > 1 else 50}
    fonts["legend"] = fonts["label"]
    # gap between tick labels and axis title, used as the title padding on both axes
    fonts["label_gap"] = round((bands[1][0] - bands[0][1]) / k) if len(bands) > 1 else 20
    # line thickness = vertical stroke height where a trace runs flat (centre moves < 1 px between neighbours)
    thick = []
    for _, hexc in series:
        rgb = np.array([int(hexc[i:i + 2], 16) for i in (0, 2, 4)])
        m = (np.abs(small - rgb).sum(axis=2) < 90) & inner
        cols = {x: runs(m[:, x]) for x in range(l + 3, r - 2)}
        for x in range(l + 4, r - 3):
            here, left, right = cols[x], cols[x - 1], cols[x + 1]
            if len(here) == len(left) == len(right) == 1:
                c = lambda rr: (rr[0][0] + rr[0][1]) / 2
                if abs(c(left) - c(right)) < 1:
                    thick.append(here[0][1] - here[0][0])
    lw = round(float(np.median(thick)) / k) if thick else 12
    frame_lw = round(float(np.median([e - s for s, e in runs(dark[t - 5: t + 8, (l + r) // 2])])) / k)
    spec = {
        "renderer": "line-plot",
        "frame": [to_orig(v) for v in (l, t, r, b)],
        "x": {"lim": [round(float(px2x(l)), 1), round(float(px2x(r)), 1)], "major": xt_step(xticks, xt)},
        "y": {"lim": [round(float(px2y(b)), 2), round(float(px2y(t)), 2)], "major": xt_step(yticks, yt)},
        "legend": [to_orig(v) for v in lg] if lg else None,
        "font": fonts, "lw": lw, "frame_lw": frame_lw,
        "series": out,
    }
    return spec


def xt_step(vals, found):
    return float(round((vals[1] - vals[0]) / max(len(found) - 1, 1), 6))


if __name__ == "__main__":
    import sys
    import yaml
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("png")
    ap.add_argument("--series", nargs="+", required=True, help='name=RRGGBB, e.g. "back (top)=0000FF"')
    ap.add_argument("--xticks", required=True, help="first,last major tick value on the x axis")
    ap.add_argument("--yticks", required=True, help="first,last major tick value on the y axis")
    ap.add_argument("--xlabel", default="")
    ap.add_argument("--ylabel", default="")
    ap.add_argument("--tol", type=float, default=0.004, help="simplification; larger = fewer points")
    ap.add_argument("-o", "--out", help="write the spec here instead of stdout")
    a = ap.parse_args()
    found = digitize(a.png, [s.rsplit("=", 1) for s in a.series], [float(v) for v in a.xticks.split(",")],
                     [float(v) for v in a.yticks.split(",")], a.tol)
    w, h = Image.open(a.png).size
    spec = {"renderer": "line-plot", "px": [w, h], "xlabel": a.xlabel, "ylabel": a.ylabel,
            **{k: v for k, v in found.items() if k != "renderer"}}
    text = yaml.safe_dump(spec, sort_keys=False, allow_unicode=True, width=10 ** 9, default_flow_style=None)
    if a.out:
        open(a.out, "w", encoding="utf-8").write(text)
    else:
        print(text)
