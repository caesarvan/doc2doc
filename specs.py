"""Redraw pictures from data written in the Markdown (a YAML spec) instead of from pixels.

Each renderer takes the parsed spec and the original pixel size and returns a PIL image.
Positions in specs are pixels of the original picture, so a redraw lands where the
original content was.
"""
import io

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
from PIL import Image  # noqa: E402


def _font(names, size):
    """First TrueType font that exists on this machine (Windows, macOS and Linux names)."""
    from PIL import ImageFont
    for n in names:
        try:
            return ImageFont.truetype(n, size)
        except OSError:
            continue
    return ImageFont.load_default()


MONO = ["consola.ttf", "Menlo.ttc", "DejaVuSansMono.ttf", "LiberationMono-Regular.ttf", "cour.ttf"]
SANS = ["arial.ttf", "Arial.ttf", "Helvetica.ttc", "LiberationSans-Regular.ttf", "DejaVuSans.ttf"]
SANS_BOLD = ["arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"]


def _figure(size, dpi=100):
    w, h = size
    return plt.figure(figsize=(w / dpi, h / dpi), dpi=dpi)


def _frac(box, size):
    """[left, top, right, bottom] in pixels -> matplotlib [x, y, w, h] figure fractions."""
    w, h = size
    l, t, r, b = box
    return [l / w, 1 - b / h, (r - l) / w, (b - t) / h]


def _to_image(fig, size):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=fig.dpi)
    plt.close(fig)
    return Image.open(buf).convert("RGBA").resize(size)


def confusion_matrix(spec, size):
    counts = np.array(spec["counts"], float)
    totals = counts.sum(axis=0)
    labels = spec["labels"]
    fs = spec.get("font", {})
    fig = _figure(size)
    ax = fig.add_axes(_frac(spec["matrix_box"], size))
    im = ax.imshow(counts, cmap=spec.get("cmap", "Greens"), vmin=0, vmax=spec.get("vmax", counts.max()),
                   aspect="auto")
    n = len(labels)
    ax.set_xticks(range(n), labels, fontsize=fs.get("ticks", 13))
    ax.set_yticks(range(n), labels, fontsize=fs.get("ticks", 13))
    ax.tick_params(top=False, labeltop=True, bottom=True, labelbottom=False)
    ax.set_xlabel(spec["x_label"], fontsize=fs.get("labels", 16))
    ax.set_ylabel(spec["y_label"], fontsize=fs.get("labels", 16))
    ax.set_title(spec["title"], fontsize=fs.get("title", 16))
    thresh = spec.get("vmax", counts.max()) / 2
    for i in range(n):
        for j in range(n):
            v = counts[i, j]
            pct = 100 * v / totals[j] if totals[j] else 0
            ax.text(j, i, f"{pct:.1f}%\n{int(v)}", ha="center", va="center", fontsize=fs.get("cells", 12),
                    color="white" if v > thresh else "black")
    cax = fig.add_axes(_frac(spec["colorbar_box"], size))
    fig.colorbar(im, cax=cax).ax.tick_params(labelsize=fs.get("colorbar", 9))
    return _to_image(fig, size)


RENDERERS = {"confusion-matrix": confusion_matrix}


def render(spec, size):
    return RENDERERS[spec["renderer"]](spec, size)


def _smooth(xy, steps=8):
    """Centripetal Catmull-Rom through simplified points, so a few points still draw a smooth trace."""
    if len(xy) < 3:
        return xy
    p = np.vstack([xy[0], xy, xy[-1]])
    out = []
    for i in range(1, len(p) - 2):
        p0, p1, p2, p3 = p[i - 1], p[i], p[i + 1], p[i + 2]
        t = np.linspace(0, 1, steps, endpoint=False)[:, None]
        out.append(0.5 * (2 * p1 + (-p0 + p2) * t + (2 * p0 - 5 * p1 + 4 * p2 - p3) * t ** 2
                          + (-p0 + 3 * p1 - 3 * p2 + p3) * t ** 3))
    out.append(xy[-1:])
    return np.vstack(out)


def line_plot(spec, size):
    """Origin-style 2D line plot: boxed frame, outward ticks, boxed legend; sizes in source pixels."""
    from matplotlib.ticker import FormatStrFormatter, MultipleLocator
    W, H = size
    dpi = 100
    pt = lambda px: px * 72 / dpi
    fig = _figure(size, dpi)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes(_frac(spec["frame"], size))
    lw = pt(spec.get("lw", 8))
    for s in spec["series"]:
        xy = _smooth(np.array([[float(v) for v in p.split(",")] for p in s["xy"].split()]))
        ax.plot(xy[:, 0], xy[:, 1], color="#" + s["color"], lw=lw, label=s["name"],
                solid_joinstyle="round", solid_capstyle="round")
    ax.set_xlim(*spec["x"]["lim"])
    ax.set_ylim(*spec["y"]["lim"])
    font = spec.get("font", {})
    fam = spec.get("family", "Arial")
    for axis, key in ((ax.xaxis, "x"), (ax.yaxis, "y")):
        major = spec[key]["major"]
        axis.set_major_locator(MultipleLocator(major))
        axis.set_minor_locator(MultipleLocator(major / 2))
        decimals = 0 if float(major).is_integer() else len(str(major).split(".")[1])
        axis.set_major_formatter(FormatStrFormatter(f"%.{decimals}f"))
    flw = pt(spec.get("frame_lw", 8))
    ax.tick_params(which="major", direction="out", length=pt(font.get("ticks", 40) * 0.45), width=flw,
                   labelsize=pt(font.get("ticks", 40)), pad=pt(font.get("ticks", 40) * 0.3))
    ax.tick_params(which="minor", direction="out", length=pt(font.get("ticks", 40) * 0.25), width=flw)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_family(fam)
    for sp in ax.spines.values():
        sp.set_linewidth(flw)
    gap = pt(font.get("label_gap", 20))
    ax.set_xlabel(spec.get("xlabel", ""), fontsize=pt(font.get("label", 50)), family=fam, labelpad=gap)
    ax.set_ylabel(spec.get("ylabel", ""), fontsize=pt(font.get("label", 50)), family=fam, labelpad=gap)
    if spec.get("legend"):
        l, t, r, b = spec["legend"]
        leg = ax.legend(loc="upper left", bbox_to_anchor=(l / W, 1 - t / H), bbox_transform=fig.transFigure,
                        prop={"family": fam, "size": pt(font.get("legend", 40))}, frameon=True, fancybox=False,
                        framealpha=1, edgecolor="black", borderpad=0.2, handlelength=2.2, labelspacing=0.15,
                        borderaxespad=0)
        leg.get_frame().set_linewidth(pt(spec.get("frame_lw", 8)) * 0.5)
    return _to_image(fig, size)


RENDERERS["line-plot"] = line_plot


def solid(spec, size):
    return Image.new("RGBA", size, "#" + spec.get("color", "FFFFFF"))


RENDERERS["solid"] = solid


def code(spec, size):
    """Notebook-style code cell: light background, monospace lines, simple Python colouring."""
    import re as _re
    from PIL import ImageDraw, ImageFont
    w, h = size
    img = Image.new("RGBA", size, "#" + spec.get("bg", "F7F7F7"))
    d = ImageDraw.Draw(img)
    fs = spec.get("font_px", 15)
    font = _font([spec["font"]] + MONO if spec.get("font") else MONO, fs)
    x0, y0, lh = spec.get("left", 50), spec.get("top", 8), spec.get("line_px", fs + 7)
    if spec.get("run_button", True):
        r = fs * 0.75
        d.ellipse([12, y0 + 2, 12 + 2 * r, y0 + 2 + 2 * r], fill="#333333")
        d.polygon([(12 + r * 0.8, y0 + 2 + r * 0.55), (12 + r * 0.8, y0 + 2 + r * 1.45), (12 + r * 1.55, y0 + 2 + r)],
                  fill="white")
    colours = [(r"#.*", "#008000"), (r"'[^']*'", "#A31515"), (r"\b\d+\b", "#098658")]
    for n, line in enumerate(spec["lines"]):
        y = y0 + n * lh
        spans = []
        for pat, c in colours:
            spans += [(m.start(), m.end(), c) for m in _re.finditer(pat, line)]
        spans.sort()
        pos, x = 0, x0
        for s, e, c in spans + [(len(line), len(line), None)]:
            if s < pos:
                continue
            for text, colour in ((line[pos:s], "#000000"), (line[s:e], c)):
                if text and colour:
                    d.text((x, y), text, font=font, fill=colour)
                    x += d.textlength(text, font=font)
            pos = e
    return img


RENDERERS["code"] = code


def table(spec, size):
    """Grid table: rows of cell text; a cell that is exactly "■" is drawn filled (schedules)."""
    from PIL import ImageDraw, ImageFont
    w, h = size
    rows = spec["rows"]
    ncol = max(len(r) for r in rows)
    cw = spec.get("col_widths") or [1 / ncol] * ncol
    rh = spec.get("row_heights") or [1 / len(rows)] * len(rows)
    xs = np.concatenate([[0], np.cumsum(cw)]) / sum(cw) * (w - 1)
    ys = np.concatenate([[0], np.cumsum(rh)]) / sum(rh) * (h - 1)
    img = Image.new("RGBA", size, "#" + spec.get("bg", "FFFFFF"))
    d = ImageDraw.Draw(img)
    fs = spec.get("font_px", max(10, int(min(rh) / sum(rh) * h * 0.45)))
    font = _font([spec["font"]] + SANS if spec.get("font") else SANS, fs)
    bold = _font(SANS_BOLD, fs)
    header = spec.get("header_rows", 1)
    for i, row in enumerate(rows):
        for j in range(ncol):
            text = str(row[j]) if j < len(row) and row[j] is not None else ""
            box = [xs[j], ys[i], xs[j + 1], ys[i + 1]]
            if text == "■":
                d.rectangle(box, fill="black")
            elif text:
                f = bold if i < header else font
                d.multiline_text((box[0] + 4, box[1] + 2), text.replace("<br>", "\n"), font=f, fill="black")
    lw = spec.get("line_px", 1)
    for x in xs:
        d.line([(x, 0), (x, h)], fill="black", width=lw)
    for y in ys:
        d.line([(0, y), (w, y)], fill="black", width=lw)
    return img


RENDERERS["table"] = table


def similarity(a, b):
    """Mean SSIM over RGB channels (Gaussian windows); both images must have the same size."""
    import cv2
    x = np.asarray(a.convert("RGB"), np.float64)
    y = np.asarray(b.convert("RGB"), np.float64)
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    blur = lambda m: cv2.GaussianBlur(m, (7, 7), 1.5)
    mx, my = blur(x), blur(y)
    vx, vy, cxy = blur(x * x) - mx * mx, blur(y * y) - my * my, blur(x * y) - mx * my
    ssim = ((2 * mx * my + c1) * (2 * cxy + c2)) / ((mx * mx + my * my + c1) * (vx + vy + c2))
    return float(ssim.mean())
