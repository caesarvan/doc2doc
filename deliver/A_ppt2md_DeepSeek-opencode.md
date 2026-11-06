# 技能手册 A：PPT → 紧凑 Markdown（DeepSeek V4.1 + opencode）

> 本手册写给 A 机上的 AI agent。照着做，可以从一台什么都没装的电脑开始，把任意 `.pptx` 转成一个很小的纯文本 Markdown。这个 md 交给 B 机（Codex），由 B 机按《技能手册 B》重建出 PPT。

## 召唤语（用户把这段发给 opencode）

```text
请完整阅读 A_ppt2md_DeepSeek-opencode.md，严格按其中的步骤，把 <PPT 文件路径> 转成 pptx2md-mini v1 格式的 Markdown。
从「1. 环境配置」开始逐步执行，不要跳步；每完成一步，用一两句话报告结果。
最终只交付一个 <PPT 文件名>.md，并按「6. 交付与报告」输出报告。
```

## 0. 目标和底线

- **产出：** 一个 `.md` 文件，格式见本文「格式契约」一节。版面（坐标、形状、颜色、字体字号）和全部文字原样保留。
- **图片不嵌像素：** 每张图片要么写成一段英文生图 prompt，要么写成可以用代码重画的数据（chart）。
- **禁止：**
  - base64、图片文件、图片链接、矢量描摹（SVG 路径）等任何形式的像素编码。
  - 修改或遗漏原文文字。
- **越小越好：**
  - 重复的图片用 `same=`。
  - prompt 控制在 40–120 个英文词。
  - 曲线图只存简化后的点。
  - 参考：11 页、41 张图的科研汇报约 55 KB，其中曲线数据约 30 KB。
- **按数据重画优先于 prompt：** 凡是能按数据重画的（纯色块、折线图、混淆矩阵、表格、代码截图），一律写成 chart。数据重画的结果与原图基本一致；AI 生图只能做到“像”。

## 1. 环境配置（从零开始）

以下命令在 PowerShell（Windows）或 bash/zsh（macOS/Linux）中执行。下文用 `PY` 代表虚拟环境里的 Python：Windows 是 `.venv\Scripts\python.exe`，macOS/Linux 是 `.venv/bin/python`。

1. 安装 uv（Python 环境管理器）。已经有 `uv --version` 输出的跳过这一步。
   - Windows：`powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
   - macOS/Linux：`curl -LsSf https://astral.sh/uv/install.sh | sh`
   - 装完重开终端，确认 `uv --version` 有输出。
2. 新建工作目录，并创建 Python 3.12 虚拟环境：
   ```text
   mkdir pptx2md
   cd pptx2md
   uv venv .venv --python 3.12
   uv pip install --python .venv python-pptx==1.0.2 lxml pillow pyyaml numpy matplotlib opencv-python-headless pymupdf
   ```
3. 可选，强烈建议：安装 LibreOffice，用来把幻灯片渲染成图片，方便看懂 SmartArt、EMF 和整页版面。
   - Windows：`winget install --id TheDocumentFoundation.LibreOffice -e`
   - macOS：`brew install --cask libreoffice`
   - Ubuntu/Debian：`sudo apt-get install -y libreoffice-impress`
4. 可选：如果你（模型）不能直接看图片，安装 OCR，用来读取图片里的文字：`uv pip install --python .venv rapidocr-onnxruntime`
5. 在 `pptx2md` 目录下，按「附录 A」原样创建三个文件：`ppt2md.py`、`digitize.py`、`specs.py`。逐字复制，不要改写。
6. 自检：`PY ppt2md.py --help` 和 `PY digitize.py --help` 都能打印出用法。

## 2. 拆解 PPT

```text
PY ppt2md.py dump "<PPT 路径>" work
PY ppt2md.py preview "<PPT 路径>" work/slides     # 需要 LibreOffice；没有就跳过
PY ppt2md.py facts work                           # 不能看图时必做；能看图也建议做
```

- **dump 的产出：**
  - `work/pics/<id>.*`：每张原图。
  - `work/pictures.json`：每张图的 id、出现在哪几页、图框尺寸 `box_cm`、宽高比 `aspect`、原图像素 `px`、所在页文字 `slide_text`。
  - `work/skeleton.md`：所有图片都还是 `TODO` 的初稿。
  - 终端里的 `WARN` 行：记下来，最后写进报告。
- **preview 的产出：** `work/slides/slideNN.png`，即整页渲染图。
- **facts 的产出：** `work/facts.json`，每张图的照片/图形判断、主色和 OCR 文字。

## 3. 逐张理解图片，填写 `work/assets.yaml`

打开 `work/pictures.json`，为每个 id 写一条记录。能看图就逐张打开 `work/pics/<id>.*`，并对照它所在页的 `work/slides/slideNN.png`。不能看图就依据 `facts.json` 的 OCR 文字、主色、类型判断和 `slide_text` 来写，并在报告里把这些图标为“需人工复核”。

### 3.1 先判断每张图属于哪一类（从上到下，命中即停）

| 图片是什么 | 写法 |
|---|---|
| 与另一张图内容相同（同一文件或同一照片） | `idX: {same: idY}` |
| 纯色块（常用来遮挡母版元素） | `chart`，用 `solid` |
| 有坐标轴的折线图，颜色区分系列，有边框和图例（Origin、Excel、matplotlib 都算） | 用 `digitize.py` 提取数据，再用 `chart_file` 引用（3.3 节） |
| 混淆矩阵 / 带数字的热力图 | `chart`，用 `confusion-matrix`，逐格抄写数字 |
| 表格截图或扫描件 | `chart`，用 `table`，逐格抄写；涂黑的格子写 `■` |
| 代码截图 | `chart`，用 `code`，逐行抄写代码 |
| 照片、3D 渲染、示意图、流程图、Logo、界面截图、3D 图表等其他所有图 | `prompt`（3.2 节） |
| 原生表格 / 原生图表 | 不用写，`ppt2md.py` 会自动输出 |
| SmartArt / 无预览的 OLE 对象（id 以 `frame` 开头） | 看 `work/slides` 里的整页图，写 `prompt` |

### 3.2 prompt 写法（决定 B 机生成的图像不像）

- 用英文，40–120 个词，写成一段。
- 第一句写清载体和主体，例如 `Photo, top view, ...`、`Flat vector infographic on white ...`、`3D CAD render ...`、`Phone screenshot, dark theme ...`。
- 按从左到右、从上到下的顺序描述构图；写明主要颜色（重要的给出十六进制色值）；说明背景。
- 画面里出现的每一段文字，都用英文双引号原样写出，例如 `labelled "Training data"`。
- 只描述图框里看得见的部分：图片有裁剪（元素行有 `crop=`）时，只写裁剪后的区域。
- 写出画面比例，例如 `wide 2:1` 或 `portrait`，与 `pictures.json` 的 `aspect` 一致。
- 不要写“同上一页”“如原图”这类依赖上下文的话。B 机看不到原图，只能读到这段文字。
- 好的例子：`Flat infographic on white, left to right: cyan signal arcs from a yellow dot, cyan text "Wireless transmission"; silver laptop "Data collection"; green arrow "Training data" and blue curved arrow "Real-time data" into a vertical cyan bar; neural network of cyan nodes with thin purple links titled pink "Training model" and purple "1D CNN"; bold blue "Real-time prediction" below. Wide 2:1.`
- 差的例子：`A diagram about machine learning.`（太笼统，没有构图、颜色和文字）

### 3.3 曲线图提取数据

1. 看图，确定以下信息：
   - 每条曲线的名称和颜色（RRGGBB；Origin 默认色依次是 000000、FF0000、0000FF、008080、FF00FF、808000 …）。
   - x 轴和 y 轴第一个、最后一个主刻度的数值。
   - 两个轴的标题。
2. 运行：
   ```text
   PY digitize.py work/pics/image11.png --series "chest=000000" "armpit=FF0000" "back (top)=0000FF" --xticks 0,1000 --yticks 0,5 --xlabel "Data Number" --ylabel "Voltage (V)" -o work/image11.yaml
   ```
3. 在 assets.yaml 里写：`image11: {chart_file: image11.yaml}`。路径相对于 assets.yaml 所在目录。
4. md 太大时，把 `--tol` 调大（默认 0.004，可以试 0.006–0.01）：点数会变少，细节略有损失。

### 3.4 assets.yaml 示例

```yaml
_title: E-Vest — March 2021          # 可选：md 的一级标题
image6:  {chart: "renderer: solid\npx: [100, 98]\ncolor: FFFFFF\n"}
image11: {chart_file: image11.yaml}
image23:
  chart: |
    renderer: confusion-matrix
    px: [1051, 833]
    title: "Accuracy: 97.778%"
    labels: [steady state, walk, run]
    x_label: True label
    y_label: Predicted label
    cmap: Greens
    vmax: 20
    counts:
      - [20, 0, 1]
      - [0, 20, 0]
      - [0, 0, 19]
    matrix_box: [152, 60, 893, 801]     # 原图中矩阵区域 [左,上,右,下] 像素
    colorbar_box: [967, 60, 1004, 801]
    font: {ticks: 13, labels: 16, title: 16, cells: 12, colorbar: 9}
image36:
  chart: |
    renderer: code
    px: [1002, 180]
    font_px: 13
    line_px: 21
    left: 50
    top: 9
    lines:
    - "# build the model and train it"
    - "model = tf.keras.Sequential()"
image41:
  chart: |
    renderer: table
    px: [730, 577]
    header_rows: 2
    col_widths: [5, 1, 1, 1, 1]
    rows:
    - ["Quarters", "Q1", "Q2", "Q3", "Q4"]
    - ["Smart brace development", "■", "■", "■", ""]
image3:
  prompt: >-
    Scientific process schematic on white, isometric 3D icons, bold black sans-serif labels. ...
image38: {same: image24}
```

## 4. 生成 md 并检查

```text
PY ppt2md.py check "<PPT 路径>" work/assets.yaml
PY ppt2md.py to-md "<PPT 路径>" work/assets.yaml "<PPT 文件名>.md"
```

- `check` 必须输出 `OK`：每张图都要有 prompt、chart 或 same。
- `check` 里 chart 的 SSIM（与原图的相似度）应当 ≥ 0.85；低于 0.8 会标 `<- check`，需要修正数据或参数（坐标框、字号、刻度）后重跑。
- `check` 提示 “prompt is short” 的，把 prompt 补到 40 词以上。
- `to-md` 的最后一行必须是 `0 without prompt/chart`；输出的 md 里不能再出现 `TODO`。
- 记下 md 的体积（KB）。

## 5. 自查清单（逐项确认）

1. 打开 md 抽查 3 页，与原 PPT（或 `work/slides`）对照：文字一字不差，元素数量合理，母版和版式区块存在。
2. md 里没有 base64、`data:`、`.png` 链接、`<svg` 路径或 `TODO`。
3. 所有 `same=` 都指向存在的 id。
4. 只交付 md 一个文件；`work/` 目录留在本机。

## 6. 交付与报告

- 交付：`<PPT 文件名>.md`。
- 报告（直接回复给用户）需要包含：
  - md 的体积，以及页数、元素数、图片数。
  - 图片分类统计：prompt / chart（各 renderer 的数量）/ same。
  - 每个 chart 的 SSIM。
  - 所有 `WARN` 行，例如被简化的渐变、组合的旋转、自由曲线、SmartArt。
  - 需人工复核的图片 id 及原因。

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

## 已知局限（如实写进报告）

- **形状和样式：**
  - 自由曲线（custGeom）按外框记成矩形。
  - 渐变填充只保留第一个颜色。
  - 组合的旋转和翻转会被忽略（子元素按位置展开）。
  - 主题里的线宽样式（lnRef）不保留。
- **段落格式：** 段落间距、行距和超链接不保留。
- **动画和切换效果：** 不保留。
- **图片背景：** 不保留，只保留纯色背景。
- **图表：** XY 散点图和气泡图只保留 y 值。
- **prompt 图：** 由 B 机重新生成，内容相似，但不会与原图逐像素一致。

## 附录 A：代码文件（逐字创建）

### ppt2md.py

````python
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
````

### digitize.py

````python
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
