# PowerPoint打开与检查

## 适用环境与默认方式

此配置用于macOS通过Parallels调用Windows内的PowerPoint。原生Mac PowerPoint和原生Windows不套用`prlctl`命令；仍使用独立副本保护用户原稿。先确认虚拟机、Office和共享路径实际可用。

在上述Parallels环境，渲染和文字测量必须先按已验证方式执行：

1. 读取最新已保存稿，记录SHA256。用户正在编辑时不触碰未保存内容，不要求关闭原稿；从已保存文件生成独立副本。
2. 通过`prlctl exec <VM> --current-user`使用当前Windows登录用户执行；不索要或写入密码。共享目录用于读取源和写回证据，实际打开文件先复制到Windows本地临时目录。
3. 使用`PowerPoint.Application.Presentations.Open(copy, -1, 0, -1)`。四个参数依次是文件、只读、是否按无标题副本打开、是否显示窗口。此处为只读、保留文件身份、显示窗口；仅将副本带到前台，不切换或保存用户原稿。
4. `scripts/render_pptx_powerpoint.ps1`与`scripts/measure_text_layout.ps1`已直接采用本地临时副本及可见窗口。检查前核对脚本是否为最新安装版，不复制旧任务脚本的后台参数。
5. 输出到新证据目录/文件，记录输入哈希、PowerPoint版本、打开方式及实际检查范围。导出成功后实际查看相关页面；文字测量不代替视觉复核。
6. 只关闭本次脚本打开的测试副本并释放COM引用；不调用`Application.Quit()`、不终止已有PowerPoint进程、不保存测试变化。源文件检查前后哈希一致。不要把占位符不能组合导致的编辑测试失败说成打开失败。

## 调用示例

以实际虚拟机名称、Windows共享路径及当前技能脚本位置替换变量。中文路径或参数复杂时，优先将PowerShell命令编码为UTF-16LE Base64后传给`-EncodedCommand`，避免多层引号被命令桥剥离。

```text
prlctl exec <VM> --current-user powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand <UTF16LE_BASE64>
```

`-ExecutionPolicy Bypass`仅对本次调用的已审阅本机技能脚本生效，不更改Windows持久策略，不用于未知来源脚本。运行源自可信开发源并已同步的脚本。

编码前的PowerShell命令：

```powershell
& '<技能的Windows共享路径>\scripts\render_pptx_powerpoint.ps1' -InputPptx '<最新保存PPT的Windows路径>' -OutputDir '<新的PNG证据目录>'
& '<技能的Windows共享路径>\scripts\measure_text_layout.ps1' -InputPptx '<同一PPT的Windows路径>' -OutputJson '<新的layout.json路径>'
```

## 有哪些不同选择

| 维度 | 本环境默认 | 其他选择及说明 |
|---|---|---|
| 执行用户 | 当前Windows登录用户 | 未指定用户的命令桥执行可能连接到不同自动化上下文，不能认定与用户桌面等效 |
| 窗口 | 显示窗口，`WithWindow=-1` | 后台不显示窗口，`WithWindow=0`；不作为本环境首选 |
| 打开位置 | Windows本地独立临时副本 | 共享目录直接打开；源保持在项目，共享路径不等于本地路径 |
| 原稿保护 | 副本只读，关闭副本不保存 | 修改或保存用户当前稿不属于显示检查 |
| 用户交互 | COM导出和测量，实际查看画面 | GUI打开是可用降级方式，但需遵守当前UI工具约束并核实打开结果 |

2026-10-11实测：此前无用户指定、后台打开方式报告OutOfMemoryException；改为当前Windows用户、可见窗口、本地独立副本后，PowerPoint16.0成功打开29页文件并导出、测量。多个条件一同改变，不能据此断言唯一根因是用户身份、窗口模式或物理内存，也不能承诺对所有环境消除该错误。不将客户文件或专属截图写入本技能。

## 打开失败时

- 先保存完整错误，记录用户身份、窗口参数、原文件与测试副本路径、PowerPoint版本，核对是否实际执行了默认方式。首次失败时只修正已发现的差异，不反复盲试相同参数。
- 对同一稳定文件在默认配置下仍失败，检查已有窗口或模态提示、共享路径/本地复制和原稿可读性；必要时按UI工具能力尝试GUI打开。没有打开不能声称实际显示已验收。
- 错误文字“内存不足”只记录为应用错误，不自动判定物理内存不足或PPT损坏；如需对照，选一份独立、已知可打开的小测试稿，避免修改用户原稿。
- 确实需要重启时，先让用户保存所有演示文稿，再按当前授权范围仅退出并重开Windows里的PowerPoint。不得依据检查任务自动结束用户进程或重启Windows、Parallels、Mac。只有PowerPoint恢复无效且有进一步证据时才讨论更大范围操作。
- 若无法完成检查，保留可用文件与证据，准确说明“结构/便携预览完成，PowerPoint实际打开或某些页面显示未完成”。
