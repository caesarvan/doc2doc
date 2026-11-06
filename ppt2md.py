"""PPTX -> compact Markdown ("pptx2md-mini v1"): layout as one-line comments, text as text,
pictures as image prompts or as data (chart specs) -- no embedded pixels.

  python ppt2md.py dump    deck.pptx WORK              # pictures -> WORK/pics, WORK/pictures.json, WORK/skeleton.md
  python ppt2md.py to-md   deck.pptx assets.yaml out.md
  python ppt2md.py check   deck.pptx assets.yaml       # every picture covered? charts re-render close to original?
  python ppt2md.py preview deck.pptx OUTDIR            # slide PNGs via LibreOffice (to look at SmartArt etc.)
  python ppt2md.py facts   WORK                        # WORK/facts.json: colours, photo/graphic guess, OCR text

assets.yaml maps picture id -> {prompt: ...} | {chart: <yaml text>} | {chart_file: path} | {same: other_id}.
"""
import argparse
import colorsys
import io
import json
import re
import sys
from collections import Counter
from pathlib import Path

import yaml
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))

CM = 360000
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
ALIGN = {"ctr": "c", "r": "r", "just": "j", "dist": "j"}
WARN = []


def num(v):
    return f"{v:.2f}".rstrip("0").rstrip(".") if v else "0"


def local(el):
    return etree.QName(el).localname


# ---------------------------------------------------------------- theme & inheritance

class Theme:
    def __init__(self, prs):
        master = prs.slide_masters[0]
        part = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        root = etree.fromstring(part.blob)
        self.colors = {}
        for el in root.find(f".//{A}clrScheme"):
            c = el[0]
            self.colors[local(el)] = c.get("lastClr") or c.get("val")
        for k, v in {"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}.items():
            self.colors[k] = self.colors[v]
        fonts = {}
        for kind in ("major", "minor"):
            f = root.find(f".//{A}{kind}Font")
            fonts[kind] = (f.find(f"{A}latin").get("typeface") if f is not None else "Calibri",
                           (f.find(f"{A}ea").get("typeface") or None) if f is not None and f.find(f"{A}ea") is not None else None)
        self.fonts = fonts
        self.default_style = prs.part._element.find(qn("p:defaultTextStyle"))
        self.tx_styles = master._element.find(qn("p:txStyles"))

    def font(self, typeface, ea=False):
        if not typeface:
            return None
        if typeface.startswith("+"):
            kind = "major" if typeface.startswith("+mj") else "minor"
            return self.fonts[kind][1 if ea else 0]
        return typeface

    def color(self, el):
        """Hex colour of a fill/colour container (<a:solidFill>, <a:fillRef>, <a:srgbClr> ...)."""
        if el is None:
            return None
        c = el if local(el).endswith("Clr") else next((ch for ch in el if local(ch).endswith("Clr")), None)
        if c is None:
            return None
        tag = local(c)
        if tag == "srgbClr":
            hexv = c.get("val")
        elif tag == "schemeClr":
            hexv = self.colors.get(c.get("val"), "000000")
        elif tag == "sysClr":
            hexv = c.get("lastClr", "000000")
        elif tag == "prstClr":
            hexv = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
                    "blue": "0000FF", "yellow": "FFFF00"}.get(c.get("val"), "000000")
        elif tag == "scrgbClr":
            hexv = "".join(f"{round(int(c.get(k)) / 100000 * 255):02X}" for k in ("r", "g", "b"))
        else:
            return None
        mod = {local(m): int(m.get("val")) / 100000 for m in c if m.get("val", "").lstrip("-").isdigit()}
        if any(k in mod for k in ("lumMod", "lumOff", "tint", "shade")):
            r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            if "tint" in mod:
                l = l + (1 - l) * (1 - mod["tint"])
            if "shade" in mod:
                l = l * mod["shade"]
            l = min(1, max(0, l * mod.get("lumMod", 1) + mod.get("lumOff", 0)))
            hexv = "".join(f"{round(v * 255):02X}" for v in colorsys.hls_to_rgb(h, l, s))
        return hexv.upper()


def lvl_props(lst, level, theme):
    """Text defaults from one list-style element (lstStyle / titleStyle / bodyStyle / defaultTextStyle)."""
    out = {}
    if lst is None:
        return out
    ppr = lst.find(f"{A}lvl{level + 1}pPr")
    if ppr is None:
        return out
    if ppr.get("algn"):
        out["al"] = ALIGN.get(ppr.get("algn"), "l")
    if ppr.find(f"{A}buNone") is not None:
        out["bu"] = None
    elif ppr.find(f"{A}buChar") is not None:
        out["bu"] = ppr.find(f"{A}buChar").get("char")
    elif ppr.find(f"{A}buAutoNum") is not None:
        out["bu"] = "1."
    d = ppr.find(f"{A}defRPr")
    if d is not None:
        out.update(rpr_props(d, theme))
    return out


def rpr_props(rpr, theme):
    out = {}
    if rpr.get("sz"):
        out["sz"] = int(rpr.get("sz")) / 100
    for k in ("b", "i"):
        if rpr.get(k) is not None:
            out[k] = rpr.get(k) in ("1", "true")
    if rpr.get("u") is not None:
        out["u"] = rpr.get("u") != "none"
    for tag, key in (("latin", "f"), ("ea", "fe")):
        el = rpr.find(f"{A}{tag}")
        if el is not None and el.get("typeface"):
            v = theme.font(el.get("typeface"), ea=(tag == "ea"))
            if v:
                out[key] = v
    c = theme.color(rpr.find(f"{A}solidFill"))
    if c:
        out["c"] = c
    return out


def ph_chain(sh):
    """[shape, layout placeholder, master placeholder] that this placeholder inherits from."""
    chain, cur = [sh], sh
    for _ in range(2):
        try:
            cur = cur._base_placeholder
        except Exception:
            cur = None
        if cur is None:
            break
        chain.append(cur)
    return chain


def text_defaults(sh, level, theme):
    """Effective paragraph/run defaults for a shape at a list level (approximates PowerPoint's chain)."""
    st = {"sz": 18.0, "f": theme.fonts["minor"][0], "fe": theme.fonts["minor"][1], "c": "000000",
          "b": False, "i": False, "u": False, "al": "l", "bu": None}
    st.update({k: v for k, v in lvl_props(theme.default_style, level, theme).items()})
    if sh.is_placeholder:
        ptype = str(sh.placeholder_format.type)
        if "TITLE" in ptype:
            style, major = "titleStyle", True
        elif any(t in ptype for t in ("BODY", "OBJECT", "SUBTITLE")):
            style, major = "bodyStyle", False
        else:
            style, major = "otherStyle", False
        if major:
            st["f"], st["fe"] = theme.fonts["major"]
        if theme.tx_styles is not None:
            st.update(lvl_props(theme.tx_styles.find(f"{P}{style}"), level, theme))
        for base in reversed(ph_chain(sh)[1:]):
            st.update(lvl_props(base._element.find(f".//{A}lstStyle"), level, theme))
    st.update(lvl_props(sh._element.find(f".//{A}lstStyle"), level, theme))
    return st


def body_props(sh):
    """Vertical anchor, wrap and insets, following placeholder inheritance."""
    out = {}
    chain = ph_chain(sh) if sh.is_placeholder else [sh]
    for s in reversed(chain):
        bp = s._element.find(f".//{A}bodyPr")
        if bp is None:
            continue
        for k in ("anchor", "wrap", "lIns", "tIns", "rIns", "bIns"):
            if bp.get(k) is not None:
                out[k] = bp.get(k)
    return out


# ---------------------------------------------------------------- element writers

def fill_of(sp_pr, theme):
    if sp_pr is None:
        return None
    if sp_pr.find(f"{A}noFill") is not None:
        return "none"
    if sp_pr.find(f"{A}solidFill") is not None:
        return theme.color(sp_pr.find(f"{A}solidFill"))
    grad = sp_pr.find(f"{A}gradFill")
    if grad is not None:  # gradients are kept as their first stop colour
        WARN.append("gradient fill simplified to its first colour")
        return theme.color(grad.find(f".//{A}gs"))
    return None


def line_of(sp_pr, theme):
    ln = sp_pr.find(f"{A}ln") if sp_pr is not None else None
    if ln is None:
        return {}
    if ln.find(f"{A}noFill") is not None:
        return {"line": "none"}
    out = {}
    c = theme.color(ln.find(f"{A}solidFill"))
    if c:
        out["line"] = c
    if ln.get("w"):
        out["lw"] = num(int(ln.get("w")) / 12700)
    dash = ln.find(f"{A}prstDash")
    if dash is not None and dash.get("val") != "solid":
        out["dash"] = dash.get("val")
    for end, key in (("headEnd", "head"), ("tailEnd", "tail")):
        e = ln.find(f"{A}{end}")
        if e is not None and e.get("type") not in (None, "none"):
            out[key] = e.get("type")
    return out


def shape_colours(sh, theme):
    sp_pr = sh._element.find(qn("p:spPr"))
    style = sh._element.find(qn("p:style"))
    attrs = {}
    f = fill_of(sp_pr, theme)
    if f is None and style is not None and style.find(f"{A}fillRef") is not None \
            and style.find(f"{A}fillRef").get("idx") != "0":
        f = theme.color(style.find(f"{A}fillRef"))
    if f:
        attrs["fill"] = f
    attrs.update(line_of(sp_pr, theme))
    has_ln = sp_pr is not None and sp_pr.find(f"{A}ln") is not None and sp_pr.find(f"{A}ln/{A}solidFill") is not None
    if "line" not in attrs and not has_ln and style is not None and style.find(f"{A}lnRef") is not None \
            and style.find(f"{A}lnRef").get("idx") != "0":
        c = theme.color(style.find(f"{A}lnRef"))
        if c:
            attrs["line"] = c
    return attrs


def style_diff(st, base):
    out = []
    if st["sz"] != base["sz"]:
        out.append(f"sz={num(st['sz'])}")
    for k in ("f", "fe"):
        if st.get(k) and st.get(k) != base.get(k):
            out.append(f'{k}="{st[k]}"')
    for k in ("b", "i", "u"):
        if bool(st[k]) != bool(base.get(k)):
            out.append(k if st[k] else f"{k}=0")
    if st["c"] != base["c"]:
        out.append(f"c={st['c']}")
    return out


def text_block(sh, theme, deck):
    """(element attrs, paragraph lines) for a shape's text frame."""
    paras = []
    for p in sh.text_frame.paragraphs:
        lvl = p.level
        base = text_defaults(sh, lvl, theme)
        ppr = p._p.find(f"{A}pPr")
        pst = dict(base)
        if ppr is not None:
            if ppr.get("algn"):
                pst["al"] = ALIGN.get(ppr.get("algn"), "l")
            if ppr.find(f"{A}buNone") is not None:
                pst["bu"] = None
            elif ppr.find(f"{A}buChar") is not None:
                pst["bu"] = ppr.find(f"{A}buChar").get("char")
            elif ppr.find(f"{A}buAutoNum") is not None:
                pst["bu"] = "1."
        runs = []
        for r in p._p:
            if local(r) in ("r", "fld"):
                rst = dict(base)
                rpr = r.find(f"{A}rPr")
                if rpr is not None:
                    rst.update(rpr_props(rpr, theme))
                t = r.find(f"{A}t")
                runs.append((t.text if t is not None and t.text else "", rst))
            elif local(r) == "br":
                runs.append(("\v", dict(base)))
        paras.append((lvl, pst, runs))
    styled = [s for _, _, runs in paras for t, s in runs if t.strip()]
    elem = dict(styled[0]) if styled else dict(paras[0][1]) if paras else dict(deck)
    attrs = style_diff(elem, deck)
    first_al = paras[0][1]["al"] if paras else "l"
    if first_al != "l":
        attrs.append(f"al={first_al}")
    bp = body_props(sh)
    if bp.get("anchor") in ("ctr", "b"):
        attrs.append("va=" + {"ctr": "m", "b": "b"}[bp["anchor"]])
    if bp.get("wrap") == "none":
        attrs.append("wrap=0")
    dflt = [91440, 45720, 91440, 45720]
    ins = [int(bp[k]) if bp.get(k) is not None else d for k, d in zip(("lIns", "tIns", "rIns", "bIns"), dflt)]
    if ins != dflt:
        attrs.append("ins=" + ",".join(num(v / CM) for v in ins))
    lines = []
    for lvl, pst, runs in paras:
        parts = []
        for t, s in runs:
            t = t.replace("\v", "\\n")
            diff = style_diff(s, elem)
            parts.append(f"[{t}]{{{' '.join(diff)}}}" if diff and t.strip() else t)
        line = "".join(parts)
        suffix = []
        if pst["al"] != first_al:
            suffix.append(f"al={pst['al']}")
        if lvl:
            suffix.append(f"lvl={lvl}")
        if pst["bu"] and line.strip():
            suffix.append(f'bu="{pst["bu"]}"')
        if suffix and line:
            line += " {" + " ".join(suffix) + "}"
        lines.append(line if line.strip() else "\\")
    while lines and lines[-1] == "\\":
        lines.pop()
    return attrs, lines


def group_xform(grp, outer):
    xfrm = grp._element.find(f".//{A}xfrm")
    off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
    choff, chext = xfrm.find(f"{A}chOff"), xfrm.find(f"{A}chExt")
    if xfrm.get("rot") or xfrm.get("flipH") or xfrm.get("flipV"):
        WARN.append(f"group '{grp.name}' rotation/flip ignored (children flattened)")
    ox, oy, cx, cy = (int(e.get(k)) for e, k in ((off, "x"), (off, "y"), (choff, "x"), (choff, "y")))
    sx = int(ext.get("cx")) / max(int(chext.get("cx")), 1)
    sy = int(ext.get("cy")) / max(int(chext.get("cy")), 1)

    def inner(x, y, w, h):
        return outer(ox + (x - cx) * sx, oy + (y - cy) * sy, w * sx, h * sy)
    return inner


def picture_part(sh):
    blip = sh._element.find(f".//{A}blip")
    if blip is None or blip.get(R_EMBED) is None:
        return None
    return sh.part.related_part(blip.get(R_EMBED))


def md_cell(text):
    return text.replace("|", "\\|").replace("\n", "<br>").replace("\v", "<br>") or " "


def table_element(sh, theme, deck, head):
    tbl = sh.table
    tblpr = sh._element.find(f".//{A}tblPr")
    sid = sh._element.find(f".//{A}tableStyleId")
    attrs = ["cols=" + ",".join(num(c.width / CM) for c in tbl.columns),
             "rows=" + ",".join(num(r.height / CM) for r in tbl.rows)]
    if sid is not None and sid.text:
        attrs.append(f"style={sid.text.strip()}")
    flags = [k for k in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol")
             if tblpr is not None and tblpr.get(k) == "1"]
    if flags:
        attrs.append("flags=" + ",".join(flags))
    first_run = next((r for row in tbl.rows for c in row.cells for p in c.text_frame.paragraphs for r in p.runs
                      if r.text.strip()), None)
    st = dict(deck)
    if first_run is not None:
        st.update(rpr_props(first_run._r.find(f"{A}rPr"), theme) if first_run._r.find(f"{A}rPr") is not None else {})
    attrs += style_diff(st, deck)
    lines = [f"<!-- {head} table {' '.join(attrs)} -->"]
    for i, row in enumerate(tbl.rows):
        cells = []
        for c in row.cells:
            tc = c._tc
            if tc.get("hMerge") == "1" or tc.get("vMerge") == "1":
                cells.append("^")
                continue
            extra = []
            if tc.get("gridSpan"):
                extra.append(f"span={tc.get('gridSpan')}")
            if tc.get("rowSpan"):
                extra.append(f"rspan={tc.get('rowSpan')}")
            tcpr = tc.find(f"{A}tcPr")
            fl = theme.color(tcpr.find(f"{A}solidFill")) if tcpr is not None else None
            if fl:
                extra.append(f"fill={fl}")
            txt = md_cell(c.text_frame.text)
            cells.append(txt + (" {" + " ".join(extra) + "}" if extra else ""))
        lines.append("| " + " | ".join(cells) + " |")
        if i == 0:
            lines.append("|" + "---|" * len(cells))
    return lines


def chart_element(sh, head, used):
    ch = sh.chart
    pid = Path(ch.part.partname).stem
    ctype = str(ch.chart_type).split(".")[-1].split(" ")[0]
    lines = [f"<!-- {head} chart id={pid} -->"]
    if pid in used:
        return lines
    used.add(pid)
    spec = {"renderer": "pptx-chart", "type": ctype}
    try:
        spec["categories"] = [str(c) for c in ch.plots[0].categories]
    except Exception:
        spec["categories"] = []
    spec["series"] = [{"name": str(s.name or ""), "values": [None if v is None else float(v) for v in s.values]}
                      for p in ch.plots for s in p.series]
    spec["legend"] = bool(ch.has_legend)
    if ch.has_title and ch.chart_title.has_text_frame:
        spec["title"] = ch.chart_title.text_frame.text
    if "XY" in ctype or "BUBBLE" in ctype:
        WARN.append(f"{pid}: scatter/bubble chart stored with y values only")
    body = yaml.safe_dump(spec, sort_keys=False, allow_unicode=True, width=10 ** 9, default_flow_style=None)
    lines.append(f"```chart id={pid}\n{body.rstrip()}\n```")
    return lines


def asset_block(pid, part_blob, assets):
    a = assets.get(pid, {})
    px = ""
    if part_blob is not None:
        try:
            img = Image.open(io.BytesIO(part_blob))
            px = f" px={img.width}x{img.height}"
        except Exception:
            pass
    if "chart" in a:
        return f"```chart id={pid}\n{a['chart'].rstrip()}\n```"
    if "same" in a:
        return f"```prompt id={pid} same={a['same']}\n```"
    return f"```prompt id={pid}{px}\n{a.get('prompt', 'TODO').strip()}\n```"


def elements(shapes, theme, deck, xf, assets, used, pics, where):
    out = []
    for sh in shapes:
        tag = local(sh._element)
        if tag == "grpSp":
            out += elements(sh.shapes, theme, deck, group_xform(sh, xf), assets, used, pics, where)
            continue
        x, y, w, h = xf(sh.left or 0, sh.top or 0, sh.width or 0, sh.height or 0)
        head = f"@{num(x / CM)},{num(y / CM)} {num(w / CM)}x{num(h / CM)}"
        xfrm = sh._element.find(f".//{A}xfrm")
        extra = []
        if xfrm is not None:
            if xfrm.get("rot"):
                extra.append(f"rot={num(int(xfrm.get('rot')) / 60000)}")
            extra += [k for k in ("flipH", "flipV") if xfrm.get(k) == "1"]
        if tag == "graphicFrame" and getattr(sh, "has_table", False) and sh.has_table:
            out += table_element(sh, theme, deck, head)
            continue
        if tag == "graphicFrame" and getattr(sh, "has_chart", False) and sh.has_chart:
            out += chart_element(sh, head, used)
            continue
        part = picture_part(sh) if tag in ("pic", "graphicFrame") else None
        if tag == "pic" or tag == "graphicFrame":
            if part is not None:
                pid = Path(part.partname).stem
                blob = part.blob
            else:  # SmartArt / OLE without preview: needs a prompt written from a rendered slide
                pid = f"frame{where}_{sh.shape_id}"
                blob = None
                WARN.append(f"{pid}: {sh.name} has no picture (SmartArt/OLE?) - describe it from the slide preview")
            video = sh.shape_type == MSO_SHAPE_TYPE.MEDIA or sh._element.find(f".//{A}videoFile") is not None
            extra.insert(0, f"id={pid}")
            if tag == "pic":
                crop = [sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom]
                if any(abs(c) > 0.001 for c in crop):
                    extra.append("crop=" + ",".join(num(c) for c in crop))
            out.append(f"<!-- {head} {'video' if video else 'img'} {' '.join(extra)} -->")
            pics.setdefault(pid, {"blob": blob, "where": [], "box": (w / CM, h / CM)})["where"].append(where)
            if pid not in used:
                used.add(pid)
                out.append(asset_block(pid, blob, assets))
            continue
        if tag == "cxnSp":
            la = line_of(sh._element.find(qn("p:spPr")), theme)
            if "line" not in la:
                st = sh._element.find(qn("p:style"))
                c = theme.color(st.find(f"{A}lnRef")) if st is not None else None
                if c:
                    la["line"] = c
            extra += [f"{k}={v}" for k, v in la.items()]
            out.append(f"<!-- {head} line {' '.join(extra)} -->")
            continue
        if tag != "sp":
            WARN.append(f"{where}: skipped <{tag}> '{sh.name}'")
            continue
        if sh.is_placeholder and "SLIDE_NUMBER" in str(sh.placeholder_format.type):
            attrs, _ = text_block(sh, theme, deck) if sh.has_text_frame else ([], [])
            out.append(f"<!-- {head} num {' '.join(attrs + extra)} -->".replace("  ", " "))
            continue
        kind = "text"
        geom = sh._element.find(f".//{A}prstGeom")
        prst = geom.get("prst") if geom is not None else ("custom" if sh._element.find(f".//{A}custGeom") is not None else "rect")
        if prst == "custom":
            WARN.append(f"{where}: freeform '{sh.name}' kept as its bounding rectangle")
            prst = "rect"
        fl = shape_colours(sh, theme)
        if prst != "rect" or fl.get("fill", "none") != "none" or fl.get("line", "none") != "none":
            kind = prst
            extra += [f"{k}={v}" for k, v in fl.items() if not (k == "fill" and v == "none")]
            adj = [g.get("fmla").split()[-1] for g in geom.iter(f"{A}gd")] if geom is not None else []
            if adj:
                extra.append("adj=" + ",".join(adj))
        has_text = sh.has_text_frame and sh.text_frame.text.strip()
        attrs, lines = text_block(sh, theme, deck) if has_text else ([], [])
        if kind == "text" and not lines:
            continue
        out.append(f"<!-- {head} {kind} {' '.join(extra + attrs)} -->".replace("  ", " "))
        out += lines
    return out


def background(owner, theme):
    bg = owner._element.find(f".//{P}bg")
    if bg is None:
        return []
    pr = bg.find(f"{P}bgPr")
    c = None
    if pr is not None:
        c = fill_of(pr, theme)
        if pr.find(f"{A}blipFill") is not None:
            WARN.append("picture background not kept")
    elif bg.find(f"{P}bgRef") is not None:
        c = theme.color(bg.find(f"{P}bgRef"))
    return [f"<!-- bg={c} -->"] if c and c != "none" else []


def load_assets(path):
    if not path:
        return {}
    assets = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    for a in assets.values():
        if isinstance(a, dict) and "chart_file" in a:
            a["chart"] = (Path(path).parent / a.pop("chart_file")).read_text(encoding="utf-8")
    return assets


def convert(src, assets):
    """-> (markdown text, pictures {id: {blob, where, box}})"""
    prs = Presentation(src)
    theme = Theme(prs)
    fonts, sizes = Counter(), Counter()
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    base = text_defaults(sh, p.level, theme)
                    for r in p.runs:
                        st = dict(base)
                        if r._r.find(f"{A}rPr") is not None:
                            st.update(rpr_props(r._r.find(f"{A}rPr"), theme))
                        fonts[(st["f"], st.get("fe"))] += len(r.text)
                        sizes[st["sz"]] += len(r.text)
    (font, font_ea), size = (fonts.most_common(1) or [((theme.fonts["minor"][0], None), 0)])[0][0], \
        (sizes.most_common(1) or [(18.0, 0)])[0][0]
    deck = {"sz": size, "f": font, "fe": font_ea, "c": "000000", "b": False, "i": False, "u": False}
    used, pics = set(), {}
    ident = lambda x, y, w, h: (x, y, w, h)
    title = assets.get("_title") or (prs.core_properties.title or Path(src).stem)
    fe = f' fe="{font_ea}"' if font_ea else ""
    lines = [f"# {title}",
             f'<!-- deck {num(prs.slide_width / CM)}x{num(prs.slide_height / CM)} f="{font}"{fe} sz={num(size)}'
             f" · pptx2md-mini v1 · 单位 cm -->", "", "## 母版"]
    master = prs.slide_masters[0]
    lines += background(master, theme)
    lines += elements([s for s in master.shapes if not s.is_placeholder], theme, deck, ident, assets, used, pics, "M")
    layouts = []
    for slide in prs.slides:
        if slide.slide_layout not in layouts:
            layouts.append(slide.slide_layout)
    for lay in layouts:
        body = background(lay, theme) + elements([s for s in lay.shapes if not s.is_placeholder], theme, deck,
                                                 ident, assets, used, pics, f"L{layouts.index(lay) + 1}")
        hide = lay._element.get("showMasterSp") == "0"
        lines += ["", f'## 版式 {lay.name}' + (" <!-- nomaster -->" if hide else "")] + body
    for i, slide in enumerate(prs.slides, 1):
        tags = [f'layout="{slide.slide_layout.name}"']
        if slide._element.get("showMasterSp") == "0":
            tags.append("nomaster")
        if slide._element.get("show") == "0":
            tags.append("hidden")
        lines += ["", f"## {i} <!-- {' '.join(tags)} -->"]
        lines += background(slide, theme)
        lines += elements(slide.shapes, theme, deck, ident, assets, used, pics, str(i))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            lines += ["<!-- notes -->"] + slide.notes_slide.notes_text_frame.text.strip().splitlines()
    return "\n".join(lines) + "\n", pics


def slide_texts(src):
    prs = Presentation(src)
    out = {}
    for i, s in enumerate(prs.slides, 1):
        out[str(i)] = " | ".join(sh.text_frame.text.strip().replace("\n", " ") for sh in s.shapes
                                 if sh.has_text_frame and sh.text_frame.text.strip())[:400]
    return out


def dump(src, work):
    work = Path(work)
    (work / "pics").mkdir(parents=True, exist_ok=True)
    md, pics = convert(src, {})
    (work / "skeleton.md").write_text(md, encoding="utf-8")
    texts = slide_texts(src)
    index = []
    for pid, info in pics.items():
        entry = {"id": pid, "slides": info["where"], "box_cm": [round(v, 2) for v in info["box"]],
                 "aspect": round(info["box"][0] / max(info["box"][1], 0.01), 3),
                 "slide_text": texts.get(str(info["where"][0]), "")}
        if info["blob"] is not None:
            img = Image.open(io.BytesIO(info["blob"]))
            ext = (img.format or "png").lower().replace("jpeg", "jpg")
            fn = work / "pics" / f"{pid}.{ext}"
            fn.write_bytes(info["blob"])
            entry.update(file=str(fn), px=list(img.size))
            if ext in ("emf", "wmf"):
                entry["note"] = "vector metafile - convert with LibreOffice or describe from slide preview"
        else:
            entry["file"] = None
        index.append(entry)
    (work / "pictures.json").write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(index)} pictures -> {work / 'pics'}; skeleton {len(md.encode()) / 1024:.1f} KB")
    for w in sorted(set(WARN)):
        print("WARN", w)


def facts(work):
    """Machine-readable facts per dumped picture, for models that cannot look at images:
    size, aspect, photo-or-graphic guess, dominant colours, and OCR text when rapidocr is installed."""
    work = Path(work)
    index = json.loads((work / "pictures.json").read_text(encoding="utf-8"))
    try:
        from rapidocr_onnxruntime import RapidOCR
        ocr = RapidOCR()
    except Exception:
        ocr = None
        print("rapidocr not installed: no OCR text (pip install rapidocr-onnxruntime)")
    for e in index:
        if not e.get("file") or e["file"].lower().endswith((".emf", ".wmf")):
            continue
        img = Image.open(e["file"]).convert("RGB")
        thumb = img.copy()
        thumb.thumbnail((200, 200))
        colours = len(thumb.getcolors(maxcolors=200 * 200) or [])
        q = thumb.quantize(colors=6)
        pal = q.getpalette()[:18]
        counts = sorted(q.getcolors(), reverse=True)
        e["kind_guess"] = "photo/render" if colours > 2500 else "graphic/diagram/chart" if colours > 40 else "flat/solid"
        e["dominant_colours"] = ["#%02X%02X%02X" % tuple(pal[i * 3: i * 3 + 3]) for _, i in counts[:5]]
        if ocr is not None:
            res, _ = ocr(str(e["file"]))
            e["ocr_text"] = [r[1] for r in (res or []) if r[2] > 0.5]
    (work / "facts.json").write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"facts for {len(index)} pictures -> {work / 'facts.json'}")


def to_md(src, assets_path, out):
    assets = load_assets(assets_path)
    md, pics = convert(src, assets)
    Path(out).write_text(md, encoding="utf-8")
    todo = [p for p in pics if p not in assets]
    print(f"{out}: {len(md.encode()) / 1024:.1f} KB, {len(pics)} pictures, {len(todo)} without prompt/chart")
    for w in sorted(set(WARN)):
        print("WARN", w)
    if todo:
        print("TODO:", " ".join(todo))


def check(src, assets_path):
    import specs
    assets = load_assets(assets_path)
    _, pics = convert(src, assets)
    bad = 0
    for pid, info in pics.items():
        a = assets.get(pid)
        if not a:
            print(f"{pid}: MISSING (no prompt/chart/same)")
            bad += 1
        elif "same" in a and a["same"] not in assets:
            print(f"{pid}: same={a['same']} points to nothing")
            bad += 1
        elif "prompt" in a and len(a["prompt"].split()) < 25:
            print(f"{pid}: prompt is short ({len(a['prompt'].split())} words) - add subject, layout, colours, text")
        elif "chart" in a:
            spec = yaml.safe_load(a["chart"])
            if spec.get("renderer") == "pptx-chart":
                continue
            size = tuple(spec["px"])
            out = specs.render(spec, size)
            if info["blob"] is not None:
                orig = Image.open(io.BytesIO(info["blob"]))
                k = min(1.0, 1200 / max(size))
                sm = (max(1, round(size[0] * k)), max(1, round(size[1] * k)))
                bg = lambda im: Image.alpha_composite(Image.new("RGBA", im.size, "white"), im.convert("RGBA")).convert("RGB")
                s = specs.similarity(bg(orig.resize(sm)), bg(out.resize(sm)))
                print(f"{pid}: chart {spec['renderer']} SSIM {s:.3f}" + ("  <- check" if s < 0.8 else ""))
    print("OK" if not bad else f"{bad} problem(s)")


def preview(src, outdir):
    """Render slides to PNG with LibreOffice + PyMuPDF."""
    import shutil
    import subprocess
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if not soffice and Path(cand).exists():
            soffice = cand
    if not soffice:
        raise SystemExit("LibreOffice (soffice) not found")
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)], check=True,
                       capture_output=True, timeout=600)
        import pymupdf as fitz
        pdf = fitz.open(next(Path(tmp).glob("*.pdf")))
        for i, page in enumerate(pdf, 1):
            page.get_pixmap(matrix=fitz.Matrix(1600 / page.rect.width, 1600 / page.rect.width)).save(
                str(out / f"slide{i:02d}.png"))
    print(f"{len(pdf)} slides -> {out}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("dump"); p.add_argument("src"); p.add_argument("work")
    p = sub.add_parser("to-md"); p.add_argument("src"); p.add_argument("assets"); p.add_argument("out")
    p = sub.add_parser("check"); p.add_argument("src"); p.add_argument("assets")
    p = sub.add_parser("preview"); p.add_argument("src"); p.add_argument("outdir")
    p = sub.add_parser("facts"); p.add_argument("work")
    a = ap.parse_args()
    {"dump": lambda: dump(a.src, a.work), "to-md": lambda: to_md(a.src, a.assets, a.out),
     "check": lambda: check(a.src, a.assets), "preview": lambda: preview(a.src, a.outdir),
     "facts": lambda: facts(a.work)}[a.cmd]()
