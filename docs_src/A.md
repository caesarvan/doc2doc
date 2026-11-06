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

{{INCLUDE:format_spec.md}}

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

{{FILE:ppt2md.py}}

### digitize.py

{{FILE:digitize.py}}

### specs.py

{{FILE:specs.py}}
