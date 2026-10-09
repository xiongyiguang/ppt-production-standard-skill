---
name: ppt-production-standard
description: "为 PPTX/POTX 制作、改版与审查提供模板、文字、对象和渲染验收规范；作为制作技能的规范层配合使用，不单独承担售前内容策划或 HTML 制作。"
---

# PPT Production Standard

本技能是 PowerPoint 规范层。由当前制作任务统筹，文件操作使用可用的 `pptx` 或演示文稿技能；规范执行不反向启动内容编排，避免同一任务重入。

## 强制制作与验收门禁

下次调用本技能即执行以下规则。新建或整体优化逐页执行；局部修改仅执行授权页，审查模式只报告。不得从旧生成脚本恢复用户删页或覆盖其最新保存稿。

- 逐页复查表达、信息完整性和版式，必须提升专业度与有效信息密度；不能以缩字、堆字或添饰品充数。业务流程、分层架构及重复模块页是重点复核对象，按[内容设计](references/content-led-design.md)执行。
- 正文页标题明确包含“当前板块 - 本页主题”；标题下概述默认按实际阅读尺寸形成两行，交代核心判断与具体安排。封面、目录过渡、结束页除外；复杂内容最多三行。详见完整规范3.3。
- 首页和目录过渡必须遵循当前权威模板：已登记中国移动V2首页居中、彩讯2026首页左对齐；目录板块之间插入对应过渡页，突出当前板块。不得把一次项目的四板块数量固化为所有方案。模板版本与示例边界见[模板登记](references/template-profiles.md)。
- 分层架构必须同时表达具体模块、上下层调用或数据关系、外部接口和相关横向治理，不能只列空泛层名；缺乏依据的组件标为待确认，不能为增加信息量虚构技术栈、资源或承诺。详见完整规范3.6。
- 新建或整体优化必须选取重点页加载imagegen技能并生成版式参考，通常覆盖业务流程、需求/职责或架构中的2—3页；小稿可只选一页。借鉴构图、阅读路径和主次后原生重构，减少表格感与装饰图标。局部改字不触发整稿制图，局部重构只研究该页。详见[内容设计](references/content-led-design.md)。
- 每个授权正文页必须先判断主要关系、再选主体结构，并记录选型理由；采用表格须说明核对对象、共同字段或对应维度及逐项阅读需求。“名称＋说明”、缺少数值或制作方便均不足以支持表格选型。流程、层级、协同等关系仅被铺成行列而未表达时判为待改；等大卡片和无边线网格按同一标准检查。判据、例外及逐页记录见[内容设计](references/content-led-design.md)。
- 复核全稿表格与重复结构：保留确需比较、映射和核对的页面；重复结构例外须注明页码、共同维度与逐页适用理由。重点页研究不替代其他页选型；只有换色、删边线、换列数不能算结构改善。模块识别需要图标时采用轻量开放布局，图标数量服从当前要求。
- 整稿全部界面截图必须采用同一线条母样；局部任务只统一授权范围内截图：默认1.25磅、accent1主题主色至透明的线性渐变。用户已有明确截图母样时先提取并统一传播；照片、Logo、图标不套用截图线条。母样与验收见完整规范9.4.4和`assets/ui-screenshot-border.xml`。
- 上述强制项必须分别留下实际检查结果，不能用结构审计通过代替内容、层级、构图、截图可读性和PowerPoint视觉验收。

## macOS 执行说明

- Python 结构审计、文字扫描和 SVG 工具可在 macOS 使用；先通过 `load_workspace_dependencies` 获取应用捆绑的 Python/依赖，不假定系统 `python` 或 `python3` 已配置。
- `render_pptx_powerpoint.ps1`、`measure_text_layout.ps1` 是 Windows PowerPoint COM 工具，不在 Mac 上执行，也不以安装 PowerShell 视作已获得 COM 支持。
- PPTX 制作与预览配合当前可用的 `presentations` 技能。第三方渲染只能作为辅助证据；要求的 PowerPoint 实际打开、自动折行、行距、组合操作和全页视觉验收须在实际安装的 PowerPoint 中验证：可用 Mac 原生版本，也可用 Windows 虚拟机内版本。先核实安装位置、共享路径和自动化能力，不假定 macOS 已安装 Office。缺少应用时明确待验收，不降低原标准。
- 模板、Logo、主题与固定几何保留原样。先核对 Mac 字体可用性；微软雅黑等字体缺失时说明影响并确认替代，不擅自更换或从旧 Windows 复制系统字体。

## 模板与范围

1. 从用户当前指定的权威模板开始；否则按已明确品牌选捆绑模板，不按清单顺序猜品牌。
2. 彩讯新项目使用 2026 主模板；明确为中国移动创新研究院的项目使用创新研究院专用模板，其他中国移动项目使用中国移动V2通用模板，不仅凭“研究院”字样推断机构。旧彩讯 V2 仅用于历史兼容或明确要求。其他品牌使用其已确认模板，不强套彩讯品牌。
3. 实测并记录尺寸、母版/版式、标题占位、Logo 与固定几何及文件哈希。保存输入和模板，通常输出新文件；不修改权威模板母版、主题、版式和固定装饰。
4. 仅做审查时不改文件；局部修订服从用户要求的范围，不自动扩展成全文内容重写或品牌迁移。

## 规范按需读取

- 新建或整体改版内容页：按 [内容驱动的设计判断](references/content-led-design.md) 明确本页判断、关系、重点与条件后自由构图，并执行其中的缩略图与实际阅读尺寸设计复核。
- 每页必须作结构判断，但无需每页套用某种命名方法。售前常见的职责交接、需求映射、业务对象、架构调用及截图讲解，按[售前关系画法](references/presales-relationship-layouts.md)只读相关小节；其他关系按[12类关系表达方法](references/relationship-patterns.md)选读，不把方法当固定模板。
- 新建/整体规范化：读 [完整规范](references/ppt-production-standard.md) 的 1–3 节，再按正在处理的文字、形状、图表等对象读取对应章节；这些是专业要求，不要求先整读所有模式。
- 新建、修订、只读审查、审计和渲染：读 [production-modes](references/production-modes.md) 对应模式及 Required QA。
- 从截图、旧页、PDF 重建或修复素材：读 [reference-slide-reconstruction](references/reference-slide-reconstruction.md)，并执行原生对象和素材溯源检查。
- 原生箭头/连接器：读完整规范 7.7。默认开放式端点（OOXML `arrow`），禁止用实心三角端点（`triangle`）代替；先核对生成工具的实际端点，再统一全稿母样，在100%及50%下检查同色、同视觉线宽和连续性。
- 新建或修改原生文字、形状时读 [editing-quality](references/editing-quality.md)：统一文字属性、检查实际多行行距与组合限制；交付前运行其中的只读检查，不以结构审计通过代替编辑性验收。
- 用户要求局部视觉润色：读 [local-visual-polish](references/local-visual-polish.md)；确需图片生成才加载 imagegen。

- 用户明确要求将PPT主体文字、表格或结构图通过imagegen制成替换图片，或连续发送页面截图要求同类制图时，读 [imagegen-content-replacement](references/imagegen-content-replacement.md)。此模式可交付独立图片，不要求回填PPT；不能把仅做局部插画的原生文字规则误用于已授权的主体位图制作。

## Non-negotiable rules

- 新增或生成的插图、配图、图标默认全部采用2D平面风格，不使用3D、等距立体、拟物材质或厚重立体效果；只有用户明确要求时才采用3D。通过构图、字号层级、色彩、留白和关系表达丰富页面。权威模板已有的3D品牌图形、真实照片和产品截图保持原貌，不因本规则擅自重绘或删除。素材选择、imagegen提示词与交付视觉检查均执行此边界。
- Use integer font sizes. Use 14 pt for ordinary body text. Never go below 12 pt; use 12 pt sparingly.
- Use Microsoft YaHei for Chinese and Arial for Latin text and numbers, or use theme fonts that resolve to those typefaces in the 彩讯科技 theme.
- Put text directly inside rectangles, rounded rectangles, cards, and labels. Do not overlay a separate text box on a shape that already serves as the text container.
- Use straight-corner rectangles for the outer frame of a page section, large module, grouped area, or overall content container.
- Default to straight-corner rectangles for both outer and inner containers, including cards, title bars, content backgrounds, labels and process nodes. Use rounded rectangles sparingly only for a clear semantic need or explicit user request; they are not a default beautification style.
- Check both outer and inner shapes against the straight-corner default; an outer straight frame alone does not establish compliance. Verify a reason for rounded exceptions without inventing a numeric quota. Preserve the original fixed geometry of authoritative templates; do not batch-convert their existing rounded shapes.
- Never change the bundled templates' master, layout, theme, handout-master, or notes-master parts. If a master object uses a fixed color, legacy font, fractional size, or other inherited setting, preserve it and report it as a template exception; do not “repair” it.
- Use standalone text boxes only for genuinely independent titles, labels, captions, footers, and connector annotations.
- Default shape text to vertical middle alignment. Use 1.0 line spacing for one visual line and 1.2 for two or more visual lines. Set paragraph spacing before and after to zero.
- Control text position with shape margins. Do not simulate margins with spaces, tabs, repeated returns, or displaced overlay text boxes.
- 品牌配色与生成图片验收统一见完整规范第8.4A节；字体层级见第4.2节。提示词中的色值、字重和尺寸要求不等于实测通过。
- For 彩讯科技 decks, use the 彩讯科技 color scheme; for 中国移动 decks, use the 中国移动 color scheme. In both cases, use `schemeClr`-equivalent theme references. A fixed RGB is noncompliant even when its value equals the theme color.
- Derive tints, shades, transparency, gradients, borders, icons, charts, and tables from theme slots so they remain theme-linked.
- Allow original colors only for logos, photos, screenshots, customer/co-brand assets, or another explicitly required standard.
- Prefer native theme-linked PowerPoint shapes or recolorable vector icons for generic icons. Fixed-color PNG/SVG artwork is an explicit media exception and must not be described as theme-linked.
- When content does not fit at 14 pt, edit, enlarge, reorganize, or split the slide. Do not solve density by shrinking below 12 pt.

## Resources

- [references/ppt-production-standard.md](references/ppt-production-standard.md): complete V1.9 production and acceptance standard. For native arrows/connectors, apply section 7.7 before editing and during rendered QA; use one deck-wide connector master and inspect shaft/head continuity at 100% and 50%.
- [references/reference-slide-reconstruction.md](references/reference-slide-reconstruction.md): conditional workflow for screenshots, flattened slides, SVG sources, icons, and generated visual assets.
- [assets/彩讯股份2026版-20260725-封面页、内容页、过渡页、结尾页模板.pptx](assets/彩讯股份2026版-20260725-封面页、内容页、过渡页、结尾页模板.pptx): primary 彩讯 2026 template; ordinary content layout is `2_内容1` (`slideLayout3`).
- [assets/新建PPT模板-彩讯科技主题色规范修正版V2.pptx](assets/新建PPT模板-彩讯科技主题色规范修正版V2.pptx): legacy 彩讯 V2 baseline; use only for legacy or explicitly requested work; ordinary content layout is `内容页1` (`slideLayout3`).
- [assets/新建PPT模板-中国移动主题色规范修正版V2.pptx](assets/新建PPT模板-中国移动主题色规范修正版V2.pptx): 当前中国移动默认模板；`内容页1` 对应 `slideLayout1`，首页居中。旧通用文件只作历史兼容，不作为默认值。
- [模板登记](references/template-profiles.md)：两份用户指定模板的来源、哈希、实测版式、首页对齐及目录过渡边界。
- `assets/ui-screenshot-border.xml`：界面截图默认1.25磅主题主色至透明渐变线条母样。
- [中国移动创新研究院专用模板](<assets/新建PPT模板-中国移动主题色规范修正版 - 创新研究院.pptx>): 明确为中国移动创新研究院的项目使用；`内容页1` 对应 `slideLayout1`，主题色方案为“中国移动”。审计时将 `--require-template` 指向此文件，保留其自身母版和固定装饰，不与通用模板混用。
- `scripts/audit_pptx.py`: read-only structural checker for `.pptx` and `.potx` files.
- `scripts/verify_svg.py`: validates editable SVG source pages and optionally requires embedded image assets.
- `scripts/embed_svg_icons.py`: embeds mapped icon files into SVG source pages.
- `scripts/extract_icon_components.py`: extracts separated components from a generated white-background icon sheet; use only with manual semantic-count review.
- `scripts/verify_pptx_text.py`: scans text-bearing PPTX parts for encoding damage and optional placeholder residue.
- `scripts/render_pptx_powerpoint.ps1`: Windows PowerPoint COM exporter; run inside Windows (including an available VM), never directly on macOS. Resolve the actual installed Office environment before choosing automation or UI export.
