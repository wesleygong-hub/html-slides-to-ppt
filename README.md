# HTML Slides to PPT

一个面向 Windows 的本地 HTML 幻灯片转 PowerPoint 工具。

默认输出“对象可编辑版 PPTX”：程序会把 HTML 中的文字重建为 PowerPoint 文本框，把纯色色块、边框和圆形等重建为 PowerPoint 原生形状，并把照片、SVG、Canvas、渐变等复杂视觉保留为各自独立的图片对象。因此常见页面不再是一张整页背景图，可以分别选择、移动、缩放和编辑对象，同时尽量保持浏览器中的布局排版。Windows 安装了桌面版 Microsoft PowerPoint 时，程序还会把常见 HTML 入场动效映射为原生 PowerPoint 动画。

## Windows 快速使用

1. 安装 [Node.js 20+](https://nodejs.org/)。
2. 首次使用，双击 `setup-windows.cmd`。
3. 以后直接双击 `start-windows.cmd`。
4. 选择 `.html` 文件和输出 `.pptx` 位置，等待转换完成。

也可以把 HTML 文件拖到 `start-windows.cmd` 上。默认文件名以 `-objects.pptx` 结尾。

程序优先使用系统自带的 Microsoft Edge，其次是 Google Chrome，最后尝试 Playwright Chromium。如果都无法启动，可双击 `install-chromium-fallback.cmd` 安装专用 Chromium。

## 三种输出模式

- `objects`（默认）：文字可编辑；纯色色块、边框、线条和圆形为 PowerPoint 原生形状；照片、SVG、Canvas、图表和渐变为分别可选择的图片对象。
- `editable`：复杂视觉合并成一张高保真背景，文字为独立文本框。适合只改文字、优先追求视觉一致性的情况。
- `image`：每页作为一张整页图片，视觉最稳定，但页面内容不可编辑。

### 视觉一致性优先

如果交付要求是“保留浏览器最终画面、不因对方电脑缺字体而变形”，使用整页高清图片模式：

```powershell
node convert.mjs -i demo.html -o demo-fidelity.pptx --mode image --scale 2 --animations off --keep-images
```

该模式不保留可编辑文字或页面内动画。`objects` / `editable` 使用 PowerPoint 自身的排版引擎，无法承诺任意 HTML 在完全可编辑的同时像素级一致；需要接收方安装相同字体，字体字重、字距和复杂裁切仍应逐页核对。建议重要交付同时生成可编辑版与静态保真版。

`--keep-images` 在所有模式下保存**完整浏览器最终画面**，不是移除文字或视觉对象后的底图，可与 PowerPoint 导出的逐页图片对照。

## 命令行使用

```powershell
node convert.mjs --input "C:\decks\demo.html" --output "C:\decks\demo-objects.pptx"
```

常用选项：

```text
--mode objects              对象拆分模式（默认）
--mode editable             背景保真 + 文字可编辑
--mode image                整页图片模式
--selector "section.slide"  自定义每页的 CSS 选择器
--scale 2                   生成 2×/4K 高清图片对象（默认）
--scale 1                   使用原始像素倍率，减小文件和内存占用
--scale 3                   生成更高精度图片对象，文件会进一步增大
--wait 1000                 每次切页后等待 1 秒
--fit contain               保持比例并完整放入 16:9 页面
--fit cover                 保持比例并铺满，可能裁边
--fit stretch               强制拉伸铺满
--page-base layout          对象模式底图放入共享版式（默认）
--page-base slide           对象模式底图保留为每页可选图层
--animations auto           自动写入原生 PPT 动画；不可用时保留静态版（默认）
--animations off            不采集或写入动画
--animations required       动画写入失败时终止转换
--keep-images               同时保留逐页 PNG
--browser msedge            固定使用 Microsoft Edge
```

运行 `node convert.mjs --help` 查看完整参数。

## 支持的 HTML 结构

程序按顺序自动识别：

1. Reveal.js；
2. `section.slide`；
3. `[data-slide]`；
4. `.slides > .slide`；
5. `.slide`；
6. 普通 `section`。

最简单的结构：

```html
<section class="slide">第一页</section>
<section class="slide">第二页</section>
```

如果你的结构不同，通过 `--selector` 指定每一页元素。例如：

```powershell
node convert.mjs -i demo.html --selector ".deck-page"
```

程序也兼容常见的 JavaScript 切页方式，包括 Reveal.js、`deck.showSlide(i)`、`deck.show(i)`、全局 `showSlide(i)`，以及通过 `active`、`present` 或 `visible` 类名显示页面的自定义 deck。

页面外的导航按钮和交互浮层会在捕获时隐藏；透明且不响应鼠标的装饰覆盖层会保留。对自定义控件，也可显式加上 `data-html2ppt-ignore`，仅从导出中排除，不修改源 HTML：

```html
<button data-html2ppt-ignore>下一页</button>
```

## 转换建议

- HTML 最好使用固定 16:9 画布，例如 `1920×1080` 或 `1280×720`。
- 网络字体和远程图片会受网络影响；交付前建议把字体和图片放到 HTML 同目录。
- 可编辑模式按本机实际安装字体解析 HTML 字体栈，字体替换会在命令行提示；保留已安装的 `Noto Serif SC` / `Noto Sans SC`，不会因名称中的 `Serif` 被错误映射为宋体。网页能加载字体不代表 PowerPoint 已安装该字体，本工具不自动嵌入字体。
- 固定画布通过父级 `transform: scale()` 缩放时，字号、字距、边框和圆角与坐标使用相同倍率；同行的连续强调文字会合并为保留各段颜色与字重的富文本框。
- 默认采集 CSS Transition、CSS Animation 和 Web Animations 的时长、延迟及关键帧，并在生成对象后写入原生 PowerPoint 时间轴。
- 动画方向默认按 CSS 关键帧、`transform-origin`、`clip-path` 和页面几何关系确定。具有共同基线且逐级升高的 3 个以上阶梯或柱形，会从基线向上展开；无法取得更多证据的 `grow/scaleY` 也安全回退为从底部向上。
- 原生动画后处理需要 Windows 桌面版 Microsoft PowerPoint。默认 `auto` 模式在 PowerPoint 不可用时仍会输出静态 PPT；交付前必须确保动画存在时可使用 `--animations required`。
- `--animations off` 会恢复为只保留最终静态状态；动态数据尚未稳定时，仍可用 `--wait` 增加等待时间。
- `file://` 页面若访问了浏览器禁止的跨域资源，应改为本地相对路径或内嵌资源。
- 默认使用 `--scale 2`：1920×1080 页面按 3840×2160 捕获，SVG、Canvas、网页图表、照片和页面底图的像素尺寸同步翻倍，放大查看时更清晰。
- 对清晰度要求特别高且可以接受更大文件时，可使用 `--scale 3` 或 `--scale 4`；超长 deck 或文件大小优先时使用 `--scale 1`。

## 指定动画展示方向

转换器通常会自动判断方向。复杂布局可在已有入场动画的元素或其祖先上增加：

```html
<div class="steps" data-html2ppt-reveal-from="bottom">
  <div class="step"></div>
  <div class="step"></div>
  <div class="step"></div>
</div>
```

`data-html2ppt-reveal-from` 支持 `auto`、`top`、`right`、`bottom`、`left`。这里的边表示内容最先出现的边，例如 `bottom` 会从底边向上展开。最近元素或祖先上的值优先；子元素使用 `auto` 可取消祖先指定并恢复自动推断。该属性只调整已经存在的有限入场动画，不会为静态元素创建动画；显式边缘用于原本不支持方向的入场效果时，会转换为 Wipe。

PowerPoint Wipe 的方向映射为 `top → 1`、`right → 2`、`bottom → 3`、`left → 4`。非法属性在默认 `auto` 模式下会警告并回退自动推断，在 `--animations required` 下会终止转换并指出页码和元素。

可运行 `npm test` 执行跨平台语义单测；Windows 已安装桌面版 PowerPoint 时，运行 `npm run test:powerpoint` 会转换最小阶梯样例，并回读方向、时长、延迟与触发方式，再比较动画版和静态版的 PowerPoint 整页渲染。

## 对象模式的可编辑范围

- 纯色矩形、圆角矩形、圆形、边框和线条：PowerPoint 原生形状，可改颜色、大小、位置和边框。
- 圆角矩形会按 HTML 的实际 `border-radius` 换算为 PowerPoint 圆角参数，不使用 PowerPoint 偏大的默认圆角。
- CSS `::before` / `::after` 中的文本、列表圆点会参与文字提取；空内容装饰与父级背景一起捕获，保留圆角与 `overflow:hidden` 裁切关系，避免径向渐变被拆成越界方块。此类装饰不再保证单独可编辑。
- CSS 边框绘制的三角形、箭头尖端和方向指示符会自动识别朝向，并转换为可单独编辑的 PowerPoint 原生三角形。
- 透明 SVG、Canvas 和 PNG 在截取为独立图片对象时，会继承最近的纯色父级底色，避免透明区域错误带入页面灰底。
- 程序会自动区分“纯媒体裁剪容器”和“包含正文、卡片、图表的复合内容容器”。`overflow:hidden`、圆角或遮罩不会再导致整块内容误合并；复合区域中的卡片、SVG、Canvas 和色块会按 DOM 结构自主拆分。
- 文字：PowerPoint 文本框，可修改内容、字体、字号、颜色和位置。
- 照片、SVG、Canvas、图表、阴影、渐变和复杂 CSS 裁切：独立图片对象，可单独移动、缩放、旋转和裁切；图片内部的像素或图表数据不能直接编辑。
- 只有不包含独立文字、卡片或其他视觉块的纯媒体遮罩容器，才会为了保持裁切关系合并为一个图片对象。
- 页面级伪元素等无法可靠映射到单一 DOM 节点的装饰，可能保留在最底层背景中。
- 对象模式的页面底图默认按内容去重并放入 PowerPoint 共享版式；普通编辑视图中不会再出现遮挡选择的整页底图。需要单独编辑每页底图时，可使用 `--page-base slide` 恢复为页面图层。
- 常见淡入、上浮、平移、缩放、旋转、擦除、路径绘制和柱形增长会映射为 PowerPoint 的 Fade、Fly、Zoom、Spin 或 Wipe 等近似效果，并保留 HTML 的自动触发、时长和分段延迟；上浮、下浮等纵向浮动入口统一使用更克制的 Fade。
- SVG 或 Canvas 内部无法成为独立 PPT 对象的子节点动画，会汇总并选择一个最具代表性的效果应用到对应图片对象，路径绘制和图表增长优先于普通淡入。这样可避免同一张完整图片先淡入、再擦除或因多个子节点而反复入场。
- PowerPoint 与浏览器的字体排版引擎不同，极复杂页面可能出现轻微字距或字号差异。
- 视频、音频、无限循环装饰动画和无法识别的复杂交互仍会保留为转换时的静态画面。

## 布局回归验证

```powershell
npm test
npm run test:layout
npm run test:powerpoint
```

`test:layout` 使用 `tests/fixtures/scaled-layout.html`，在 1280×720、1920×1080 两种视口检查缩放、连续文字、伪元素裁切、透明背景及页面外控件隔离；Windows 使用 Edge，其他平台需要 Playwright Chromium。`test:powerpoint` 需要 Windows 桌面版 PowerPoint。

可用 `tests/inspect-render-ppt.ps1 -PptxPath <文件> -OutputDirectory <目录>` 只读导出全部页 PNG，并输出页数、对象、字体字号和动画数量，供交付检查。
