# 技能手册 B：紧凑 Markdown → PPT（Codex）

> 本手册写给 B 机上的 AI agent（Codex）。照着做，可以从一台什么都没装的电脑开始，把 A 机产出的 pptx2md-mini v1 Markdown 重建成 `.pptx`。所有图片都由你自己生成：不调用任何外部生图 API，不读取任何 API key，最终文件里不留任何占位框。

## 召唤语（用户把这段发给 Codex）

```text
请完整阅读 B_md2ppt_Codex.md，严格按其中的步骤，把 <md 文件路径> 重建成同名的 .pptx。
图片必须由你自己生成：有原生生图能力就直接用，没有就手写 SVG。不得调用外部 API，不得留下占位框。
从「1. 环境配置」开始逐步执行，每完成一步，用一两句话报告结果；最后按「6. 交付与报告」输出报告。
```

## 0. 目标和底线

- **产出：** `<md 文件名>.pptx`。版面、文字、颜色、字号按 md 精确重建。数据图（chart）由代码按数据重画；prompt 图由你生成。
- **只用你自己的能力：**
  - 允许：原生图像生成，或手写 SVG / 用本地代码绘图。
  - 禁止：调用任何外部生图服务或 API，查找或读取任何密钥。
- **不能有占位框：** `md2ppt.py build` 默认是严格模式，缺任何一张图都会拒绝生成。`--draft` 只用来预览版面，绝不能作为交付物。
- **md 是唯一的信息来源：** 不要修改 md。没有原始 PPT 可以参考。

## 1. 环境配置（从零开始）

以下命令在 PowerShell（Windows）或 bash/zsh（macOS/Linux）中执行。下文用 `PY` 代表虚拟环境里的 Python：Windows 是 `.venv\Scripts\python.exe`，macOS/Linux 是 `.venv/bin/python`。

1. 安装 uv。已经有 `uv --version` 输出的跳过这一步。
   - Windows：`powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
   - macOS/Linux：`curl -LsSf https://astral.sh/uv/install.sh | sh`
2. 新建工作目录，并创建 Python 3.12 虚拟环境：
   ```text
   mkdir pptx2md
   cd pptx2md
   uv venv .venv --python 3.12
   uv pip install --python .venv python-pptx==1.0.2 lxml pillow pyyaml numpy matplotlib opencv-python-headless pymupdf
   ```
3. 强烈建议：安装 LibreOffice，用于第 5 步的逐页渲染检查。
   - Windows：`winget install --id TheDocumentFoundation.LibreOffice -e`
   - macOS：`brew install --cask libreoffice`
   - Ubuntu/Debian：`sudo apt-get install -y libreoffice-impress`
4. 字体：md 文件头和元素里用到的字体（`f=` / `fe=`，例如 Arial、Times New Roman、Arial Unicode MS、微软雅黑），本机缺的尽量装上，否则 PowerPoint 会自动替换字体，文字宽度也会变。Linux 至少安装 `fonts-liberation`、`fonts-dejavu` 和 `fonts-noto-cjk`。
5. 在 `pptx2md` 目录下，按「附录 B」原样创建两个文件：`md2ppt.py`、`specs.py`。逐字复制，不要改写。
6. 把要还原的 md 复制到 `pptx2md` 目录。自检：`PY md2ppt.py --help` 能打印出用法。

## 2. 列出要生成的图片

```text
PY md2ppt.py prompts deck.md > prompts.jsonl
PY md2ppt.py build deck.md draft.pptx --images images --draft    # 只用来先看版面
```

- `prompts.jsonl` 每行一张图，包含：
  - `id`：保存时的文件名。
  - `aspect`：图框的宽高比（宽 ÷ 高），生成的图必须与它一致。
  - `box_cm`：图框尺寸。
  - `orig_px`：原图像素，仅供参考。
  - `prompt`：要生成的内容。
- 用 `same=` 引用的图不会出现在列表里，它们直接复用被引用的那张。
- `chart` 图不会出现在列表里，构建时由代码自动重画。

## 3. 逐张生成图片，存到 `images/<id>.png`（或 `.jpg` / `.svg`）

**选择方式（按顺序）：**

1. 如果你有原生图像生成能力（例如内置的 image generation 工具），就用它：在 prompt 末尾加一句 `Aspect ratio <宽>:<高>.`（按 `aspect` 换算成近似整数比），生成后存为 `images/<id>.png`。如果生成结果比例不对，用本地 PIL 居中裁成 `aspect`。
2. 如果没有原生生图能力，或者某张图被拒绝或生成失败，就自己手写 SVG，存为 `images/<id>.svg`，要求如下：
   - 根元素：`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 H">`，其中 H = 1000 ÷ aspect，四舍五入取整。
   - 只用基础元素：`rect` `circle` `ellipse` `line` `polyline` `polygon` `path` `text` `g` `linearGradient` `radialGradient` `defs`。不用 filter、foreignObject、外部图片、外部字体或 CSS 文件；字体写 `font-family="Arial, sans-serif"`。
   - prompt 里双引号中的每一段文字，都用 `<text>` 原样写出，并确保不超出画面。
   - 结构必须完整：流程图要画出全部节点和箭头；表格画出全部格线；界面截图画出全部区块；图表画出坐标轴和曲线的大致走势。
   - 照片和 3D 渲染图画成插画：用渐变、高光和阴影塑造主体，主体位置、颜色和构图与 prompt 一致。
   - 先手写一张最复杂的图，用第 5 步的方法渲染出来看效果，满意后再批量画其余的图。

**统一要求：**

- 每个 id 一个文件，文件名就是 id。
- 不要在图片里加水印、边框，或额外写上 prompt 文字。
- 逐张检查：主体、构图、颜色、文字与 prompt 一致，比例与 `aspect` 一致。不合格的重做。

## 4. 构建 PPT（严格模式）

```text
PY md2ppt.py check deck.md --images images
PY md2ppt.py build deck.md deck.pptx --images images
```

- `check` 必须显示 `N/N pictures present`；有 `MISSING` 就回到第 3 步补齐。比例偏差超过 8% 的会被提示，需要裁剪或重做。
- `build` 成功时输出 `deck.pptx: N slides`。如果报“缺少图片”，回到第 3 步，不要改用 `--draft`。

## 5. 逐页检查

```text
PY md2ppt.py preview deck.pptx previews     # 需要 LibreOffice，输出 previews/slideNN.png
```

逐页打开 `previews/slideNN.png`，对照 md 中的同一页检查：

1. **文字：** 文字齐全，没有明显溢出或重叠。如果溢出是因为缺少字体，先装字体，再重新构建。
2. **形状：** 颜色、位置、箭头方向、叠放顺序与元素行一致。母版和版式元素（例如顶部和底部的色条、Logo）每页都在，标了 `nomaster` 的页面除外。
3. **图片：** 没有灰色框或任何 “missing” 字样。每张 AI 图或 SVG 图的内容与 prompt 一致，且没有变形。
4. **数据图：** chart 图（曲线、矩阵、表格、代码）清晰；原生表格和原生图表能正常显示。

发现问题：重做对应的图片，或补装字体后，重新执行第 4 步和第 5 步。没有 LibreOffice 时，在报告里说明，并请用户用 PowerPoint 打开 `deck.pptx` 目视检查。

## 6. 交付与报告

- 交付：`deck.pptx`。可以附上 `images/` 目录，方便以后复用。
- 报告（直接回复给用户）需要包含：
  - 页数。
  - 图片统计：原生生成 / 手写 SVG / chart 重画 / same 复用，各多少张。
  - 被重做的图片及原因。
  - 缺少的字体。
  - 第 5 步发现的问题，以及哪些已经修复、哪些仍然存在。

## 格式契约：pptx2md-mini v1

两台电脑之间唯一传递的就是这一个 `.md` 文件。A 机按本契约写出，B 机按本契约读回。两份手册里的这一节完全相同；如有改动，必须两边同步修改。

### 文件结构

````text
# <演示文稿标题>
<!-- deck <宽>x<高> f="<默认字体>" [fe="<默认东亚字体>"] sz=<默认字号pt> · pptx2md-mini v1 · 单位 cm -->

## 母版
<!-- bg=RRGGBB -->                      ← 可选：背景色
<元素行 + 内容 ...>                        ← 画在每一页的最底层

## 版式 <版式名> [<!-- nomaster -->]       ← 版式自带的图形；nomaster 表示这个版式不显示母版图形
<元素行 + 内容 ...>

## <页码> <!-- layout="<版式名>" [nomaster] [hidden] -->
<!-- bg=RRGGBB -->                      ← 可选：本页背景（优先级 本页 > 版式 > 母版）
<元素行 + 内容 ...>
<!-- notes -->                          ← 可选：之后各行是演讲者备注
<备注文字>
````

- 叠放顺序：母版元素 → 版式元素 → 本页元素，同一层内按出现顺序，后出现的在上面。
- 页面标题行带 `nomaster`：本页不画母版和版式图形。版式带 `nomaster`：使用该版式的页面不画母版图形。
- `hidden`：隐藏页。

### 元素行

````text
<!-- @x,y wxh kind 属性... -->
````

- `x,y` 是左上角坐标，`w,h` 是宽和高，单位都是 cm，原点在页面左上角。坐标可以为负或超出页面，超出部分按页面裁掉。
- **kind 的取值：**

| kind | 含义 | 后面跟的内容 |
|---|---|---|
| `text` | 文本框（无填充、无边框） | 文字行 |
| `num` | 页码（内容就是页序号） | 无 |
| `line` | 直线：从 (x,y) 到 (x+w,y+h)；`flipH` 交换两端的 x，`flipV` 交换两端的 y | 无 |
| `img` / `video` | 图片 / 视频（视频用封面图代替） | 该 id 第一次出现时跟一个 `prompt` 或 `chart` 代码块 |
| `table` | 原生表格 | Markdown 表格 |
| `chart` | 原生图表 | 该 id 第一次出现时跟 `chart` 代码块（`renderer: pptx-chart`） |
| 其他任意名称 | DrawingML 预设形状，如 `rect` `roundRect` `ellipse` `homePlate` `rightArrow` `downArrow` `mathPlus` `rightBrace` `triangle` … | 可选文字行（形状内文字） |

- **形状属性：**
  - `fill=RRGGBB`：填充色；没写就是无填充。
  - `line=RRGGBB|none`：边框色；预设形状没写就是无边框。`lw=`：线宽（pt）。`dash=`：DrawingML 虚线名。
  - `head=` / `tail=`：线条起点 / 终点的箭头（arrow、triangle、stealth、oval、diamond）。
  - `rot=`：顺时针旋转角度。`flipH` / `flipV`：水平 / 垂直翻转。
  - `adj=v1,v2`：形状调节值，原始值，除以 100000 得到比例。
- **文字属性（元素级默认值）：**
  - `sz=`：字号（pt）。`f="字体"`；`fe="东亚字体"`。
  - `b` / `i` / `u`：粗体、斜体、下划线；`=0` 表示取消。`c=RRGGBB`：文字颜色。
  - `al=l|c|r|j`：水平对齐。`va=m|b`：垂直居中 / 底部，默认顶部。
  - `wrap=0`：不自动换行。`ins=l,t,r,b`：内边距（cm），默认 0.25,0.13,0.25,0.13。
  - 没写的属性取文件头的默认值；文字颜色默认 `000000`。
- **图片属性：**
  - `id=<资源 id>`。
  - `crop=l,t,r,b`：四边裁掉的比例。只对 `chart` 图生效；`prompt` 图描述的就是框里看到的样子，不再裁剪。
  - `rot`：同样只对 `chart` 图生效。

### 文字行

- 元素行之后，直到下一个元素行、`##` 标题、代码块或 `<!-- notes -->` 之前，每一行是一个段落。
- 单独一个 `\` 表示空段落。段落内的 `\n`（反斜杠加 n 两个字符）表示软换行。
- 行尾的 ` {al=c lvl=1 bu="•"}` 是段落属性，前面必须有一个空格：
  - `al`：段落对齐，覆盖元素级设置。
  - `lvl`：缩进级别。
  - `bu`：项目符号字符；`bu="1."` 表示自动编号。
- `[文字]{sz=36 c=FF0000 b}` 表示一段文字使用不同样式，属性语法与元素行相同。

### 表格

````text
<!-- @1,4 11x4 table cols=3.67,3.67,3.67 rows=1.33,1.33,1.33 style={5C22544A-7EE6-4342-B048-85BDC9FD1C3A} flags=firstRow,bandRow sz=14 -->
| 指标 | Q2 | Q3 |
|---|---|---|
| 收入 {fill=FFE0A0} | 10 | 12 |
````

- `cols` / `rows`：各列宽 / 各行高（cm）。`style`：PowerPoint 内置表格样式的 GUID。`flags`：表格选项。
- Markdown 的第一行就是表格的第一行，分隔线只是 Markdown 语法，不算一行。
- 单元格末尾的 ` {span=2 rspan=2 fill=RRGGBB}`：横向合并、纵向合并、底色。被合并掉的单元格写 `^`。单元格内换行写 `<br>`，竖线写 `\|`。

### 图片与数据图（代码块）

````text
```prompt id=image7 px=748x413
English prompt for an image model: medium, subject, layout, colours, exact quoted text ...
```
```prompt id=image38 same=image24
```
```chart id=image23
renderer: confusion-matrix
px: [1051, 833]
...
```
````

- `prompt`：交给图像模型（或由 B 机 agent 自己绘制）的英文描述。`px` 是原图像素尺寸，仅供参考；生成时以图框的宽高比为准。
- `same=Y`：复用图片 Y，不再单独生成。
- `chart`：用代码按数据重画，必须带 `renderer` 和 `px`（`pptx-chart` 除外）。可用的 renderer：

| renderer | 用途 | 主要字段 |
|---|---|---|
| `solid` | 纯色图（遮挡块等） | `color` |
| `line-plot` | 折线图（由 `digitize.py` 生成） | `frame` `x{lim,major}` `y{lim,major}` `legend` `font` `lw` `frame_lw` `xlabel` `ylabel` `series[{name,color,xy:"x,y x,y …"}]` |
| `confusion-matrix` | 混淆矩阵 / 热力图 | `title` `labels` `x_label` `y_label` `cmap` `vmax` `counts`（行 = 预测，列 = 真实） `matrix_box` `colorbar_box` `font` |
| `table` | 表格类图片（含扫描件） | `rows[[...]]` `col_widths` `row_heights` `header_rows` `font_px`；单元格内容正好是 `■` 时画成实心格 |
| `code` | 代码截图 | `lines` `font_px` `line_px` `left` `top` `bg` `run_button` |
| `pptx-chart` | PowerPoint 原生图表（自动生成） | `type`（python-pptx 的 XL_CHART_TYPE 名称） `categories` `series[{name,values}]` `legend` `title` |

- 所有像素坐标（`frame`、`legend`、`matrix_box` 等）都以原图左上角为原点，格式为 `[左, 上, 右, 下]`。

## 已知局限

- prompt 图由你生成，只能做到与原图“相似”。图里的小字以 prompt 中引号内的文字为准。
- 视频只还原成封面图。动画、切换效果和超链接不在 md 里，无法还原。
- 构建时使用 python-pptx 默认模板加空白版式，母版和版式图形直接画在每一页上（外观相同，但不是 PowerPoint 里的母版对象）。

## 附录 B：代码文件（逐字创建）

### md2ppt.py

````python
"""Compact Markdown ("pptx2md-mini v1") -> .pptx.

  python md2ppt.py prompts deck.md                 # JSON lines: every picture the agent must create
  python md2ppt.py check   deck.md --images DIR    # which pictures are still missing
  python md2ppt.py build   deck.md out.pptx --images DIR [--draft]
  python md2ppt.py preview out.pptx OUTDIR         # slide PNGs via LibreOffice, for visual checks

Pictures: ```chart``` blocks are redrawn from their data (specs.py, or native PowerPoint charts);
```prompt``` blocks must exist as DIR/<id>.png|jpg|jpeg|webp|svg, made by the agent itself.
--draft puts a grey box where a picture is missing (layout preview only; never ship it).
"""
import argparse
import io
import json
import re
import sys
from pathlib import Path

import yaml
from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

sys.path.insert(0, str(Path(__file__).parent))
import specs  # noqa: E402

CM = 360000
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
ELEM = re.compile(r"^<!-- @(-?[\d.]+),(-?[\d.]+) ([\d.]+)x([\d.]+) (\S+)(.*?) -->$")
FENCE = re.compile(r"^```(prompt|chart) (.*)$")
SPAN = re.compile(r"\[([^\]]*)\]\{([^}]*)\}")
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".webp", ".svg")


def parse_attrs(text):
    out = {}
    for k, v, flag in re.findall(r'(\w+)=("[^"]*"|\S+)|(\w+)', text or ""):
        if flag:
            out[flag] = "1"
        else:
            out[k] = v.strip('"')
    return out


def parse_md(path):
    """-> deck, sections {"master": sec, "layout:<name>": sec}, slides [sec], assets {id: (kind, attrs, body)}
    where sec = {"attrs": {...}, "bg": hex|None, "elems": [...], "notes": [...]}."""
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    deck, sections, slides, assets = {}, {}, [], {}
    sec, elem, in_notes, i = None, None, False, 0
    while i < len(lines):
        line = lines[i]
        m = FENCE.match(line)
        if line.startswith("<!-- deck "):
            deck = parse_attrs(line[len("<!-- deck "):-3])
            w, h = line.split()[2].split("x")
            deck["w"], deck["h"] = float(w), float(h)
        elif line.startswith("## "):
            head, _, comment = line[3:].partition("<!--")
            sec = {"attrs": parse_attrs(comment.replace("-->", "")), "bg": None, "elems": [], "notes": []}
            name = head.strip()
            if name == "母版":
                sections["master"] = sec
            elif name.startswith("版式"):
                sections["layout:" + name[2:].strip()] = sec
            else:
                slides.append(sec)
            elem, in_notes = None, False
        elif m:
            attrs = parse_attrs(m.group(2))
            body = []
            i += 1
            while i < len(lines) and lines[i] != "```":
                body.append(lines[i])
                i += 1
            assets[attrs["id"]] = (m.group(1), attrs, "\n".join(body))
        elif sec is not None and line.startswith("<!-- bg="):
            sec["bg"] = line[len("<!-- bg="):-3].strip()
        elif sec is not None and line.startswith("<!-- notes -->"):
            elem, in_notes = None, True
        elif sec is not None and (em := ELEM.match(line)):
            x, y, w, h = (float(v) for v in em.group(1, 2, 3, 4))
            elem = {"box": (x, y, w, h), "kind": em.group(5), "attrs": parse_attrs(em.group(6)), "text": []}
            sec["elems"].append(elem)
            in_notes = False
        elif in_notes and sec is not None:
            sec["notes"].append(line)
        elif elem is not None and line.strip():
            elem["text"].append(line)
        i += 1
    return deck, sections, slides, assets


def all_elems(sections, slides):
    for sec in list(sections.values()) + slides:
        yield from sec["elems"]


# ---------------------------------------------------------------- pictures

def resolve(ident, assets):
    seen = set()
    while ident in assets and assets[ident][1].get("same") and ident not in seen:
        seen.add(ident)
        ident = assets[ident][1]["same"]
    return ident


def image_file(ident, images_dir):
    if not images_dir:
        return None
    for ext in IMAGE_EXT:
        f = Path(images_dir) / f"{ident}{ext}"
        if f.exists():
            return f
    return None


def svg_to_png(svg_path, width_px):
    import pymupdf as fitz
    doc = fitz.open(str(svg_path))
    page = doc[0]
    k = width_px / max(page.rect.width, 1)
    return page.get_pixmap(matrix=fitz.Matrix(k, k), alpha=True).tobytes("png")


def draft_box(text, aspect):
    w = 900
    h = max(60, round(w / max(aspect, 0.05)))
    img = Image.new("RGB", (w, h), (242, 242, 242))
    d = ImageDraw.Draw(img)
    d.rectangle([2, 2, w - 3, h - 3], outline=(160, 160, 160), width=3)
    try:
        font = ImageFont.truetype("arial.ttf", 22)
    except OSError:
        font = ImageFont.load_default()
    d.text((14, 10), text[:300], fill=(90, 90, 90), font=font)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def picture(ident, assets, images_dir, box, draft, missing):
    """-> (png/jpg bytes or None, 'chart'|'native'|'image')"""
    rid = resolve(ident, assets)
    kind, attrs, body = assets.get(rid, ("prompt", {}, ""))
    if kind == "chart":
        spec = yaml.safe_load(body)
        if spec.get("renderer") == "pptx-chart":
            return spec, "native"
        buf = io.BytesIO()
        specs.render(spec, tuple(spec["px"])).save(buf, "PNG")
        return buf.getvalue(), "chart"
    f = image_file(rid, images_dir) or image_file(ident, images_dir)
    if f is None:
        missing.add(rid)
        return (draft_box(f"[missing {rid}] {body}", box[2] / max(box[3], 0.01)) if draft else None), "image"
    if f.suffix.lower() == ".svg":
        return svg_to_png(f, max(600, round(box[2] / 2.54 * 220))), "image"
    if f.suffix.lower() == ".webp":
        buf = io.BytesIO()
        Image.open(f).save(buf, "PNG")
        return buf.getvalue(), "image"
    return f.read_bytes(), "image"


# ---------------------------------------------------------------- shapes & text

def rgb(hexv):
    return RGBColor.from_string(hexv)


def emu(*vals):
    return [Emu(round(v * CM)) for v in vals]


def apply_line(shape, a):
    if a.get("line") == "none":
        shape.line.fill.background()
    elif a.get("line"):
        shape.line.color.rgb = rgb(a["line"])
    if a.get("lw"):
        shape.line.width = Pt(float(a["lw"]))
    ln = shape.line._get_or_add_ln()
    if a.get("dash"):
        etree.SubElement(ln, qn("a:prstDash")).set("val", a["dash"])
    for key, tag in (("head", "a:headEnd"), ("tail", "a:tailEnd")):
        if a.get(key):
            etree.SubElement(ln, qn(tag)).set("type", a[key])


def set_font(run, st):
    f = run.font
    f.name, f.size = st["f"], Pt(st["sz"])
    f.bold, f.italic, f.underline = st["b"], st["i"], st["u"]
    f.color.rgb = rgb(st["c"])
    if st.get("fe"):
        rpr = run._r.get_or_add_rPr()
        ea = rpr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(rpr, qn("a:ea"))
        ea.set("typeface", st["fe"])


def base_style(a, deck):
    return {"f": a.get("f", deck.get("f", "Calibri")), "fe": a.get("fe", deck.get("fe")),
            "sz": float(a.get("sz", deck.get("sz", 18))), "c": a.get("c", "000000"),
            "b": a.get("b") == "1", "i": a.get("i") == "1", "u": a.get("u") == "1"}


def fill_paragraph(p, line, base):
    pos, chunks = 0, []
    for m in SPAN.finditer(line):
        chunks.append((line[pos:m.start()], {}))
        chunks.append((m.group(1), parse_attrs(m.group(2))))
        pos = m.end()
    chunks.append((line[pos:], {}))
    for text, over in chunks:
        if not text:
            continue
        st = dict(base)
        for k, v in over.items():
            st[k] = float(v) if k == "sz" else (v != "0" if k in ("b", "i", "u") else v)
        for n, piece in enumerate(text.split("\\n")):
            if n:
                p._p.add_br()
            if piece:
                r = p.add_run()
                r.text = piece
                set_font(r, st)


ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY}


def set_text(tf, lines, a, deck):
    base = base_style(a, deck)
    tf.word_wrap = a.get("wrap") != "0"
    tf.auto_size = None
    if a.get("va"):
        tf.vertical_anchor = {"m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM, "t": MSO_ANCHOR.TOP}[a["va"]]
    if a.get("ins"):
        l, t, r, b = (Emu(round(float(v) * CM)) for v in a["ins"].split(","))
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = l, t, r, b
    for n, line in enumerate(lines):
        p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
        pa = {}
        if (m := re.search(r" \{([^{}]*)\}$", line)) and "=" in m.group(1):
            pa, line = parse_attrs(m.group(1)), line[: m.start()]
        p.alignment = ALIGN[pa.get("al", a.get("al", "l"))]
        lvl = int(pa.get("lvl", 0))
        if lvl:
            p.level = min(lvl, 8)
        ppr = p._p.get_or_add_pPr()
        if pa.get("bu"):
            ppr.set("marL", str(round((lvl + 1) * 0.5 * CM)))
            ppr.set("indent", str(round(-0.5 * CM)))
            if pa["bu"] == "1.":
                etree.SubElement(ppr, qn("a:buAutoNum")).set("type", "arabicPeriod")
            else:
                etree.SubElement(ppr, qn("a:buChar")).set("char", pa["bu"])
        else:
            etree.SubElement(ppr, qn("a:buNone"))
        if line != "\\":
            fill_paragraph(p, line, base)


def add_table(slide, e, deck):
    x, y, w, h = e["box"]
    a = e["attrs"]
    rows = []
    for line in e["text"]:
        if re.match(r"^\|(\s*-+\s*\|)+$", line.replace(":", "")):
            continue
        cells = [c.strip().replace("\\|", "|") for c in re.split(r"(?<!\\)\|", line.strip())[1:-1]]
        rows.append(cells)
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    gf = slide.shapes.add_table(len(rows), ncol, *emu(x, y, w, h))
    tbl = gf.table
    if a.get("cols"):
        for col, v in zip(tbl.columns, a["cols"].split(",")):
            col.width = Emu(round(float(v) * CM))
    if a.get("rows"):
        for row, v in zip(tbl.rows, a["rows"].split(",")):
            row.height = Emu(round(float(v) * CM))
    tblpr = gf._element.find(f".//{A}tblPr")
    if a.get("style"):
        sid = tblpr.find(qn("a:tableStyleId"))
        if sid is None:
            sid = etree.SubElement(tblpr, qn("a:tableStyleId"))
        sid.text = a["style"]
    flags = a.get("flags", "").split(",") if "flags" in a else None
    if flags is not None:
        for k in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol"):
            tblpr.set(k, "1" if k in flags else "0")
    base = base_style(a, deck)
    for i, row in enumerate(rows):
        for j in range(ncol):
            raw = row[j] if j < len(row) else ""
            if raw == "^":
                continue
            ca = {}
            if (m := re.search(r" \{([^{}]*)\}$", raw)):
                ca, raw = parse_attrs(m.group(1)), raw[: m.start()]
            cell = tbl.cell(i, j)
            span, rspan = int(ca.get("span", 1)), int(ca.get("rspan", 1))
            if span > 1 or rspan > 1:
                cell.merge(tbl.cell(min(i + rspan, len(rows)) - 1, min(j + span, ncol) - 1))
            if ca.get("fill"):
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(ca["fill"])
            tf = cell.text_frame
            for n, part in enumerate(raw.strip().split("<br>")):
                p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
                if part.strip():
                    fill_paragraph(p, part.strip(), base)


def add_native_chart(slide, e, spec):
    x, y, w, h = e["box"]
    ctype = getattr(XL_CHART_TYPE, spec.get("type", "COLUMN_CLUSTERED"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    data = CategoryChartData()
    cats = spec.get("categories") or [str(i + 1) for i in range(len(spec["series"][0]["values"]))]
    data.categories = cats
    for s in spec["series"]:
        data.add_series(s.get("name") or "", s["values"])
    gf = slide.shapes.add_chart(ctype, *emu(x, y, w, h), data)
    ch = gf.chart
    ch.has_legend = bool(spec.get("legend", True))
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    if spec.get("title"):
        ch.has_title = True
        ch.chart_title.text_frame.text = spec["title"]


def add_element(slide, e, deck, assets, images_dir, slide_no, draft, missing):
    x, y, w, h = e["box"]
    a, kind = e["attrs"], e["kind"]
    if kind in ("img", "video", "chart"):
        data, src = picture(a["id"], assets, images_dir, e["box"], draft, missing)
        if data is None:
            return
        if src == "native":
            add_native_chart(slide, e, data)
            return
        pic = slide.shapes.add_picture(io.BytesIO(data), *emu(x, y, w, h))
        if src == "chart":  # generated pictures already show exactly the visible part
            if a.get("crop"):
                pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = (float(v) for v in a["crop"].split(","))
            if a.get("rot"):
                pic.rotation = float(a["rot"])
        return
    if kind == "table":
        add_table(slide, e, deck)
        return
    if kind == "line":
        x1, y1, x2, y2 = x, y, x + w, y + h
        if a.get("flipH"):
            x1, x2 = x2, x1
        if a.get("flipV"):
            y1, y2 = y2, y1
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, *emu(x1, y1, x2, y2))
        apply_line(c, a)
        return
    if kind in ("text", "num"):
        shape = slide.shapes.add_textbox(*emu(x, y, w, h))
    else:
        try:
            prst = MSO_SHAPE.RECTANGLE if kind == "rect" else MSO_SHAPE.from_xml(kind)
        except Exception:
            print(f"WARN slide {slide_no}: unknown shape '{kind}', drawn as rectangle")
            prst = MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(prst, *emu(x, y, w, h))
        shape.shadow.inherit = False
        if a.get("fill") and a["fill"] != "none":
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb(a["fill"])
        else:
            shape.fill.background()
        if "line" not in a:
            shape.line.fill.background()
        apply_line(shape, a)
        if a.get("adj"):
            for n, v in enumerate(a["adj"].split(",")):
                try:
                    shape.adjustments[n] = int(v) / 100000
                except (IndexError, ValueError):
                    pass
    if a.get("rot"):
        shape.rotation = float(a["rot"])
    xfrm = shape._element.find(f".//{A}xfrm")
    for k in ("flipH", "flipV"):
        if a.get(k):
            xfrm.set(k, "1")
    lines = [str(slide_no)] if kind == "num" else e["text"]
    if lines:
        set_text(shape.text_frame, lines, a, deck)


def set_bg(slide, colour):
    if colour:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(colour)


def build(md_path, out, images_dir=None, draft=False):
    deck, sections, slides, assets = parse_md(md_path)
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(deck["w"], deck["h"])
    blank = prs.slide_layouts[6]
    master = sections.get("master", {"elems": [], "bg": None, "attrs": {}})
    missing = set()
    for n, sec in enumerate(slides, 1):
        slide = prs.slides.add_slide(blank)
        layout = sections.get("layout:" + sec["attrs"].get("layout", ""), {"elems": [], "bg": None, "attrs": {}})
        set_bg(slide, sec["bg"] or layout["bg"] or master["bg"])
        hide_slide = sec["attrs"].get("nomaster")
        layer = []
        if not hide_slide and not layout["attrs"].get("nomaster"):
            layer += master["elems"]
        if not hide_slide:
            layer += layout["elems"]
        for e in layer + sec["elems"]:
            add_element(slide, e, deck, assets, images_dir, n, draft, missing)
        if sec["notes"]:
            slide.notes_slide.notes_text_frame.text = "\n".join(sec["notes"]).strip()
        if sec["attrs"].get("hidden"):
            slide._element.set("show", "0")
    if missing and not draft:
        raise SystemExit("缺少图片，未生成 PPT：" + " ".join(sorted(missing)) +
                         f"\n请把它们存为 {images_dir or 'images'}/<id>.png|jpg|svg 后重试")
    prs.save(out)
    print(f"{out}: {len(slides)} slides" + (f" (draft, {len(missing)} missing pictures)" if missing else ""))


def prompt_list(md_path):
    deck, sections, slides, assets = parse_md(md_path)
    boxes = {}
    for e in all_elems(sections, slides):
        if e["kind"] in ("img", "video"):
            boxes.setdefault(resolve(e["attrs"]["id"], assets), e["box"])
    out = []
    for ident, (kind, attrs, body) in assets.items():
        if kind == "prompt" and not attrs.get("same"):
            b = boxes.get(ident, (0, 0, 1, 1))
            out.append({"id": ident, "aspect": round(b[2] / max(b[3], 0.01), 3),
                        "box_cm": [b[2], b[3]], "orig_px": attrs.get("px"), "prompt": " ".join(body.split())})
    return out


def check(md_path, images_dir):
    items = prompt_list(md_path)
    have = [i for i in items if image_file(i["id"], images_dir)]
    print(f"{len(have)}/{len(items)} pictures present in {images_dir}")
    for i in items:
        f = image_file(i["id"], images_dir)
        if not f:
            print(f"  MISSING {i['id']} (aspect {i['aspect']})")
        elif f.suffix.lower() != ".svg":
            im = Image.open(f)
            ratio = im.width / im.height
            if abs(ratio - i["aspect"]) / i["aspect"] > 0.08:
                print(f"  {i['id']}: aspect {ratio:.2f} vs box {i['aspect']:.2f} (will be stretched)")


def preview(src, outdir):
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
    p = sub.add_parser("prompts"); p.add_argument("md")
    p = sub.add_parser("check"); p.add_argument("md"); p.add_argument("--images", default="images")
    p = sub.add_parser("build"); p.add_argument("md"); p.add_argument("out")
    p.add_argument("--images", default="images"); p.add_argument("--draft", action="store_true")
    p = sub.add_parser("preview"); p.add_argument("src"); p.add_argument("outdir")
    a = ap.parse_args()
    if a.cmd == "prompts":
        for item in prompt_list(a.md):
            print(json.dumps(item, ensure_ascii=False))
    elif a.cmd == "check":
        check(a.md, a.images)
    elif a.cmd == "build":
        build(a.md, a.out, a.images, a.draft)
    else:
        preview(a.src, a.outdir)
````

### specs.py

````python
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
````
