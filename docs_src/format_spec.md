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
