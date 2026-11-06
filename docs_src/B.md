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

{{INCLUDE:format_spec.md}}

## 已知局限

- prompt 图由你生成，只能做到与原图“相似”。图里的小字以 prompt 中引号内的文字为准。
- 视频只还原成封面图。动画、切换效果和超链接不在 md 里，无法还原。
- 构建时使用 python-pptx 默认模板加空白版式，母版和版式图形直接画在每一页上（外观相同，但不是 PowerPoint 里的母版对象）。

## 附录 B：代码文件（逐字创建）

### md2ppt.py

{{FILE:md2ppt.py}}

### specs.py

{{FILE:specs.py}}
