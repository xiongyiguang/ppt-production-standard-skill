# PowerPoint 操作模式与验收

命令以 Skill 根为基准。仅读取当前新建、修订、重建、审计模式。

## Select the workflow

### Create a new presentation

1. Confirm the audience, scenario, source material, requested length, and which authoritative template applies.
2. Start from a copy of the selected template file itself. Do not create a blank deck and merely imitate its colors.
3. Preserve the selected template's master, layouts, theme, slide size, logo, fixed page furniture, and all master decorations byte-for-byte. Do not open Slide Master view to “normalize” or restyle them.
4. Bind every ordinary content slide to the content layout measured from the resolved template. The primary 彩讯 2026 template uses `2_内容1`; the 中国移动 and legacy 彩讯 V2 templates use `内容页1`. Use title, section, blank, or ending layouts only for their corresponding non-content slide roles.
5. Treat the bundled template's single sample slide as a structure and theme reference, not as a mandatory page design. Its existing straight outer containers and rounded inner shapes are preserved template examples, not permission to default new inner shapes to rounded corners. Preserve that sample slide's original shape geometry, including its 14 rounded rectangles. If the final deck does not need the sample page, remove or replace the whole page rather than restyling the template example.
6. Build a slide-by-slide content map before editing the file. Review every page for concrete content, semantic hierarchy and effective information density. Include brand-matched cover, chapter transition pages and section-aware body titles. Select key pages for actual imagegen layout exploration under content-led-design.md, then reconstruct native objects with fewer decorative icons.
7. Apply the complete reference standard, then run the audit and rendered QA.

### Normalize or revise an existing presentation

1. Resolve one authoritative template as the destination visual system, even when the source deck uses another master.
2. Audit before editing:

   ```bash
   python scripts/audit_pptx.py "input.pptx"
   ```

3. Separate content problems from formatting problems. Preserve wording, facts, slide order, and information structure unless the user authorizes changes to those content elements.
4. Treat master migration and rebinding ordinary content slides to the selected template's measured content layout as required formatting corrections already authorized by this standard. Do not leave newly revised content pages bound to the old deck's layout.
5. Correct the requested scope across all repeated instances, not only the first visible example.
6. Save as `原文件名-规范修正版.pptx` by default.
7. Re-run the audit, render the affected slides, compare them with the original, fix problems, and verify again.

### Optional local visual polish

After content and structure are substantially settled, use this mode when the user requests or approves local visual enhancement, illustration accents, or restrained icon additions. Do not automatically add artwork to every deck or every page. Read [references/local-visual-polish.md](../references/local-visual-polish.md) completely before selecting or producing assets. Preserve the latest user-edited deck and its brand; this mode does not authorize rewriting approved content or changing the template. Load the available imagegen skill only when generating or editing raster assets is actually needed.

### Review without editing

Inspect the file and report evidence by slide, object type, and property. Do not modify or create a revised deck unless requested.

### Reconstruct from slide images or repair visual assets

1. Follow the selected template and all rules in this skill; the reference image supplies content and visual evidence, not a replacement master or full-slide background.
2. Rebuild titles, body text, cards, connectors, tables, and charts as native PowerPoint objects by default. Keep each icon or illustration as an independent asset.
3. Use the asset-source order in the reconstruction reference. Never regenerate logos, official brand marks, screenshots, or product UI with image generation.
4. If ordinary icons or illustrations genuinely require generation, load and follow the available `imagegen` skill before calling its tool; inspect the source page first and preserve an explicit semantic inventory.
5. Treat a whole-slide SVG inserted into PowerPoint as one graphic object, not as native editable text and shapes. Do not claim native editability unless the PowerPoint objects were actually reconstructed and inspected.
6. Keep SVG pages and icon mappings as optional source artifacts. Run the SVG/text verifiers and the normal template audit, then render every reconstructed slide in Microsoft PowerPoint and compare it with the source image.
7. For case pages, do not preserve the old page's title, subtitle, long explanatory paragraphs, metrics, labels, or conclusions inside a screenshot. Recreate those elements as native PowerPoint text and keep only factual visual evidence such as product UI,现场照片、原始图表或必要示意图。
8. Use visual assets in this order: project originals/system exports; original media, vector or chart objects extracted from PPTX; original bitmap/vector objects extracted from PDF; high-resolution PowerPoint/PDF rendering followed by precise cropping; existing low-resolution screenshots only after the previous sources are exhausted and the user accepts the limitation.
9. Do not use ImageGen, AI upscaling, sharpening, or repainting to fabricate factual product UI, case screenshots, metrics, logos, or evidence. If the source can only be rendered, start PDF rendering at 300 DPI and use 450–600 DPI for small UI text; crop from the rendered source, never from an already compressed preview.

## Content-led overview and visual structure

For content-slide authoring and redesign, follow reference sections 3.3–3.5 before placing shapes. Write a useful overview under the title, two actual rendered lines by default and at most three for complex content; document genuine exceptions without filler. Choose diagrams only when they clarify real relationships; retain tables, checklists, or prose when those are clearer. Use the per-slide record and table eligibility criteria in [content-led-design](content-led-design.md): identify the main relationship and source, chosen primary structure and reason, retained conditions, then actual rendered findings and status. A table needs explicit objects, common comparison or mapping dimensions, and a cell-by-cell reading benefit. Review primary-table slide numbers, structure distribution and repeated ranges; exceptions need concrete pages, common dimensions and per-page applicability. Most pages using tables or equal-row grids triggers a selection review, not a table quota. Key-page imagegen research cannot clear unreviewed pages. Relational content flattened into name/description rows remains a design defect even when drawn as native shapes. Preserve user-edited titles, order, conditions, responsibilities, and exceptions when changing the visual expression.

Structural compliance is not visual-design acceptance. Separately inspect hierarchy, focal emphasis, spacing, diagram clarity, and cross-page rhythm; an editable, theme-compliant deck can still be monotonous.

## Deterministic audit

Use the bundled checker before and after edits:

```bash
python scripts/audit_pptx.py "deck.pptx" --json
python scripts/audit_pptx.py "richinfo-deck.pptx" --require-template "assets/彩讯股份2026版-20260725-封面页、内容页、过渡页、结尾页模板.pptx" --require-layout "2_内容1" --content-slides "2-8" --require-color-scheme "彩讯科技" --strict
python scripts/audit_pptx.py "china-mobile-deck.pptx" --require-template "assets/新建PPT模板-中国移动主题色规范修正版V2.pptx" --require-layout "内容页1" --content-slides "2-8" --require-color-scheme "中国移动" --strict
```

Replace the illustrative `--content-slides "2-8"` range with the actual ordinary content-slide numbers from the content map. Treat the checker as structural evidence, not a substitute for visual review. With `--require-template`, it byte-compares protected master, layout, theme, handout-master, and notes-master parts against the selected baseline and fails if any changed. Unchanged template slides and inherited master settings are baseline exceptions; strict font, size, and fixed-color checks apply to new or changed slide objects. Rectangle counts are descriptive only: visually verify that outer and inner containers default to straight corners and that sparse rounded exceptions have a semantic or user-specified reason. Automatic wrapping and semantic exceptions still require rendered inspection.

## Required QA

1. Confirm the output opens as a valid presentation and preserves the expected slide count, text, objects, and geometry.
2. Extract text to check missing, duplicated, or leftover content.
3. Run the bundled audit with `--require-template`. Confirm protected template parts are unchanged, the resolved content layout exists, and every ordinary content slide is bound to it.
4. Render every changed slide in Microsoft PowerPoint. Use the actually installed PowerPoint: a native Mac app or a Windows VM. Run `.ps1` COM tools inside Windows only; verify the VM command bridge and shared paths first, or use the application UI. Preserve other open user presentations. Export into a new directory so stale previews cannot be mistaken for current output.
5. Inspect overflow, clipping, overlap, spacing, alignment, contrast, image distortion, and unexpected theme-color changes. Separately execute the thumbnail and reading-size design review in `content-led-design.md`; record concrete visual findings and fixes. Structural success cannot clear an unresolved design defect.
6. Use an independent visual-review pass when available, following the general `pptx` skill.
7. If a defect is found, correct it and verify the affected slides again before declaring the deck ready. A clean result does not require an artificial edit cycle.
8. Confirm every outer section or module frame uses straight corners. Inner cards, labels, title bars, backgrounds and process nodes also default to straight corners. Rounded exceptions require a clear semantic reason or user request; checking only the outer frame is insufficient. Preserve authoritative template geometry and avoid arbitrary numeric quotas.
9. For screenshot/image reconstruction, confirm there is no full-slide screenshot substitute, each media asset is independently selectable, the source-to-output text is complete, and the delivered editability statement matches what was actually tested.
10. Report the selected template, output path, preserved inputs, audit results, rendered scope, editability scope, and any remaining manual checks.
11. For every retained case/product image, confirm its provenance, source page/slide or media part, intrinsic pixel dimensions, crop method, displayed size, and whether its effective resolution is sufficient. Reject images enlarged from a low-resolution screenshot when a clearer original or source-page render exists.
12. Compare the entire repeated page skeleton across all content pages: left brand mark, title placeholder, subtitle baseline, right corporate logo, content safe area, footer/conclusion band, and page number. Their position and size must come from the authoritative template, not from visual approximation or a previous deck.
13. Apply reference sections 3.3–3.5 and the per-slide decision/acceptance record in [content-led-design](content-led-design.md). Check that tables serve actual comparison or mapping and that diagrams visibly carry their main relationships. Report concrete unresolved objects as pending redesign, missing render evidence as pending acceptance, and document repeated-structure exceptions by page. Structural audit cannot certify semantic selection or clear these statuses. After expanding an overview, reflow and render the content below it; do not squeeze shapes or reduce text to preserve the old content height.
14. A ZIP/XML or library-open pass does not establish PowerPoint compatibility. Require actual opening without a repair prompt and inspect the exported pages, especially native tables and edited relationships. If PowerPoint repairs the file, inspect the repair log and correct the cause; do not accept a repaired deck with removed objects as a successful output.

15. Execute the mandatory gates in SKILL.md: section-aware titles, cover alignment by authoritative brand template, chapter transitions, two-line body overviews, semantic name/description hierarchy, concrete architecture modules and relations, recorded imagegen layout research, reduced decorative icons and the consistent UI-screenshot gradient outline. Review only the authorized scope for local edits; report pending items instead of claiming the whole deck passed.
