# Office EX 开发交接 — M1

更新日期：2026-11-24。仓库：`xzcode/office-ex`，主分支：`master`。**M1.1–M1.10 已完成并通过基础设施阶段验收。** 受测源码为 `7d0522dcd2a8f6db676a6878429da2ca31a839f6`，tree 为 `25ea2dd3dfadad3c161798f36fef7896eb64fcc8`；[CI 45375454826](https://github.com/xzcode/office-ex/actions/runs/45375454826) 四个任务全部成功。后续仅补充文档、矩阵和证据，不把文档提交冒充测试源版本。

## 必须保留的决定

用户选择自主开发路线：自研 UI、三个领域模型、编辑语义、布局、公式、场景和格式适配，允许通用底层库。生产依赖没有完整 Office 引擎；不要恢复早期已废弃的 ONLYOFFICE／AGPL 方案、upstream lock 或下载／构建脚本。早期方案仅保留在 Git 历史里，未导入上游 Office 编辑器源码。

原仓库 MIT 许可证保持不变。Electron、fflate、xmldom 等通用依赖按固定版本和各自 notice 管理；见 `docs/DEPENDENCIES.md`、`docs/DEPENDENCIES.lock.json` 和 `THIRD_PARTY_NOTICES.md`。不要把 Electron 的根 MIT 描述成所有 Chromium／Node 组件都只有这一份许可证。

用户明确要求 DOC/DOCX、XLS/XLSX、PPT/PPTX 六格式。旧版目标为 Office 97–2003，XLS 为 BIFF8。M2-L、M3-L、M4-L 是首个正式版必经门槛；M1 的 CFB 容器和格式检查不会替代旧格式编辑与双向转换。

## 进入仓库先读

1. `AGENTS.md`、`docs/LICENSING.md`。
2. `docs/specs/README.md` 与本次相关规范；规范优先于旧路线图。
3. `docs/M1.md`、`docs/decisions/0001-m1-boundaries.md`。
4. `docs/specs/features.json` 的完整需求状态与 `m1Acceptance`，以及 `docs/verification/m1/README.md`、`report.json` 和原生输出。

## 已落地的 M1 代码

- npm workspaces、严格 TypeScript、固定直接依赖和单一 lockfile；公共包入口、依赖图与纯领域编译检查。
- `contracts/core/writer/sheets/slides/render`：三种模型、运行时校验、EMU／Unicode、样式／资源、命令、增量 undo/redo、dirty 和保存票据、派生任务、几何／字体接口。
- `ooxml`：有界 ZIP/XML/OPC；保留原始部件和摘要。原 M0 DOCX/XLSX XML 写入原型仍作为回归证据保留。
- `cfb`：原创 MS-CFB v3/v4 读写、FAT/DIFAT/MiniFAT、目录／storage／stream 和安全检查。DOC/XLS/PPT 语义适配尚未实现。
- `format-common`：有界二进制读取、适配器／能力／诊断／保存计划契约。
- `format-service`：在 sandbox worker 中组合容器检测和内部草稿序列化；不让低层容器互相依赖，不让 Electron main 解析文档。
- `apps/desktop`：本地协议、受限 preload/IPC、窗口／菜单、文件 token、身份检查、原子提交、替换前备份、恢复和设置。
- `ui`：多标签、首页／最近文件、中文／英文、浅／深／系统主题、三个草稿预览、Office 结构检查、取消、脏文档关闭提示。
- CI 和真正的 Electron smoke harness：基础检查、macOS ARM64、Windows x64、Linux x64 均已通过，平台输出和截图已归档。

## 当前能力与边界

`.oex` 是内部草稿 v1，保存 format-independent model；包络为 `magic: "office-ex-draft"`、`version: 1`、`document`。当前拒绝有资源元数据但缺少资源字节封装的草稿，避免图片等内容静默丢失。它不替代六种 Office 文件格式。

Office 文件目前提供容器／标记级结构检查、诊断和字节原样另存，不提供语义导入编辑。PDF 仅识别标记。草稿文字预览不分页；电子表格预览只演示 A1:C3 和 typed cell 事务，公式仅保留源式；幻灯片预览只演示基础矩形和命中。不要把这些界面描述成完整 Writer/Sheets/Slides。

Canvas 的整体文字宽度与绘制来自同一上下文，但前缀宽度推算的字簇位置标记为 `estimated`。混合方向、复杂 shaping、分页、Office 字体一致性和打印需在 M2–M5 验收。

文件提交使用同目录独占临时文件、重读验证、fsync、目标身份复核、替换前持久备份和原子 rename／新目标无覆盖提交。Node 没有跨平台 inode 条件 rename；最后一次身份检查与 rename 之间的外部写入竞态不能声称完全消除。目录同步不可用时返回较弱 durability 并显示提示；网络／不适配文件系统受到限制。

## 运行与验证

Node 24.13+ 的 24.x；CI 固定 Node 24.22.3、Python 3.12.14。首次安装：

```sh
npm ci --ignore-scripts
npm run setup:desktop
npm run start
```

验证入口：`npm run check`、`npm run test:interop`、`npm run audit:licenses`、`npm run build`、`npm run smoke:desktop`。独立 Python 工具先安装 `tests/requirements.txt`。

`tsx` 使用 `node --import tsx --test` 运行，避免 CLI 额外 IPC socket 在受限执行环境失败。Linux smoke 需要图形会话／Xvfb 和正常 sandbox；不要关闭 sandbox 获得通过。

最终源码的 99 项基础回归全部通过，独立 ZIP／CFB 互读、原 M0 XML 语义检查、架构／许可证检查及原生构建均通过。macOS 15.8.2 ARM64、Windows Server 2025 x64、Ubuntu 24.04.5 x64 各通过 15 项真实 Electron smoke。宿主回归子集：macOS／Linux 各 43 通过；Windows 42 通过、1 个设置软链接用例显式跳过。不要把这个跳过项写成通过，或把 Server CI 写成 Windows 10/11 安装认证。

依赖审查确认 66 个锁定 npm 包、23 份当前安装 notice 文件；原生构建 stage 了 Electron 和 Chromium notices。2026-11-24 npm advisory 查询为 0 已知漏洞。`.gitattributes` 固定 lock 的 LF，并保留上游 notice 的原始混合换行，防止 Windows checkout 改变审核 hash。Linux 原生 CI 安装系统中文／emoji 字体，保留 sandbox；截图等待实际可见窗口和绘制帧，最终三平台均无 capture retry。

## 接下来如何继续

2026-11-24 新增 [I18N 多语言规范](specs/09-i18n.md)（v1.1），要求首个正式版支持 22 个 UI locale；I18N-01–05/QA-07 目前均为 planned。先开展 M2-I18N：抽出统一语言包和类型化消息键、系统/手动热切换、现有 zh-CN/en 设置迁移、地区与文档语言独立、RTL 逻辑布局和 Unicode/字体接口。M1 中英文基础不等于新规范完成，历史验收不修改。M3/M4 继续地区公式、B 批语言与 PDF，多语言 M5 门槛与六格式门槛并行，不可相互替代。

下一阶段进入 M2 Writer：建立真正的段落／run 编辑、选区／IME、布局盒和分页，再实现 DOCX 语义映射与意图编辑保存。并行设计 MS-DOC 的 M2-L 子阶段，使用共享 CFB 与 Writer 模型。后续继续 M3/BIFF8 XLS、M4/PPT/PPTX/PDF；正式发布前完成六格式往返、转换、参考 Office 和各平台安装验证。

每次修改继续以受测源码、日期、具体平台和样本更新证据。当前 `features.json` 的 `m1Acceptance` 已通过；较宽的文件、安全、UI 和格式需求整体仍为 `implemented`，其 `remainingScope` 是后续工作或已明确的保证边界，不应直接整体改成 `verified`。

不要新建 Release、宣称已生成安装包、降低 P0 验收，或把 `.oex` 作为正式版兼容目标的替代。修复与验证应围绕具体剩余风险；不重复无关打包或扩大到未授权项目。

M2-I18N 首批实现见 docs/M2-I18N.md；在独立分支验证。四个候选目录不等于 22 语言正式验收。继续先确认 CI，再补 A 批七个语言、伪语言/复数、Writer 内容语言与真实输入/布局；保留 M1 验收证据。

基础检查已通过（104/104），受测源码 0862f02858dd576563bb34f22e45b2860ead6c10，CI https://github.com/xzcode/office-ex/actions/runs/45379428809；原生任务状态详见 docs/M2-I18N.md。后续接续分支 m2-i18n-20261124，不得丢失未合并工作。

M2-I18N A 批第二轮：初始四语言源码 0862f02 三平台 CI 45379428809 全通过；七种新增候选资源、结构化复数/select、伪语言测试函数和全部十一语言切换检查已接入。新提交需独立 CI，不沿用旧证据。详见 docs/M2-I18N.md；M2 后续优先 Writer，B 批及正式多语言验收仍按 v1.1。

2026-11-24：A 批源码 9d5d59b000fc35d27f4e9babb6be508d7124304f 的 CI 45380700385 四项任务全部通过，M2-I18N 初始基础可并入 master。进入 M2 Writer。剩余 11 个 UI locale、真实文档语言、复杂文字布局与 M5 审校门槛继续保留。

## M2 Writer 第一批

参见 docs/M2-WRITER.md：跨 run 段落事务、直接样式、分合段、显式分页符，独立最小 Transitional DOCX 语义读写。当前不改变桌面只读 Office 检查流程，未知属性/部件阻止语义适配。先看独立分支 m2-writer-20261124 的 CI 结果，再接入分页布局、编辑界面、安全原格式保存、DOC 与国际化文档属性。不要将窄 profile 或独立 Python 互读等同于真实 Office 完整兼容。

M2 Writer 首批验收通过：源码 7692c0c7da8e561f3d15da71edf1ab8f2b314a06，CI 45424554932，121/121 回归及独立 DOCX 互读；留在 m2-writer-20261124，下一批在此分支继续分页、富文本输入和桌面安全接入。六格式和多语言发布门槛不变。

M2 Writer 第二批候选实现：docs/M2-WRITER.md，布局与多段编辑工作区、11 locale 文案，原生 CI 待结果。后续继续 DOCX 桌面安全保存与普通 styles/资源适配、完整富文本排版和 DOC 旧格式，不降低 I18N/PDF/Office 门槛。

M2 Writer 第二批基础回归 129/129 成功，源码 84db8010c281f4eefdb80b7250b57148b3876f44，CI https://github.com/xzcode/office-ex/actions/runs/45426118923。记录时 verify-writer: success; native-macos: queued; native-other (windows-2025): in_progress; native-other (ubuntu-24.04): in_progress。继续时先核对该 CI 原生结果，若失败只修实际问题；下一批重点 DOCX 安全桌面接入、样式/资源、富文本分页精度及 DOC 兼容。保持 m2-writer-20261124，未合并 master。

M2 Writer 第三批候选：DOCX 窄 profile 桌面新建、打开编辑、保存、另存已接入，主进程独立隔离验证器负责生成/重开输出，授权缓存与既有原子保存/备份保护保留。11 个 locale 158 键。查看 docs/M2-WRITER.md 和当前分支 CI；后续再扩展普通 DOCX styles/资源、字体字号颜色工具、表格图片、完整排版、DOC 兼容。禁止把本 profile 宣传为任意 DOCX 可编辑。

M2 Writer 第三批受测源码 c79feaf16191e01b890e16012f7daca98a0a30ef，基础回归 133/133 通过，CI https://github.com/xzcode/office-ex/actions/runs/45428117703；记录时 verify-writer: success; native-other (ubuntu-24.04): in_progress; native-macos: in_progress; native-other (windows-2025): in_progress。下一轮先检查本批原生验收，针对实际失败修复；然后扩展字体/字号/颜色工具、普通 DOCX styles/资源和表格图片，继续 DOC 旧格式。原格式编辑仅支持窄 profile，不宣传通用 Word 兼容；未合并 master。

M2 Writer 第四批候选：第三批 CI 45428117703 三平台全部通过；第四批新增字体/字号/颜色工具、命名样式应用、保留式 styles/theme profile，11 locale 164 键。继续核对本批 CI 后扩展普通 Word 的默认/脚本样式、settings/numbering 依赖闭包、表格、图像及 DOC；不把受控 style fixture 当作全面 Word 兼容。详细边界见 docs/M2-WRITER.md。

M2 Writer 第四批受测源码 75cad7b0ac4a59043b666334b41e70af914dc781，137/137 基础回归、受控样式 DOCX 独立互读与非正文部件逐字节保留均已通过；CI https://github.com/xzcode/office-ex/actions/runs/45432375135。记录时 verify-writer: success; native-other (windows-2025): in_progress; native-macos: in_progress; native-other (ubuntu-24.04): in_progress。后续继续同一 Writer 分支，先核对该原生结果并只修实际问题；随后优先扩大安全 Word profile（docDefaults、字体/语言映射、settings 等影响闭包）、表格结构与布局、图片资产、旧 DOC。源样式目录仍只读，不能把本轮称为完整 Office 兼容。保留全部 22 locale 及六格式发布门槛。

M2 Writer 第五批候选：上一批 45432375135 全任务通过；本批新增规则矩形表格、行列/列宽/边框事务、cell 段落编辑、整行分页及固定表格 DOCX 语义读写，11 locale 176 键。先查本批 CI，修实际失败后继续 inline PNG/JPEG 资源管理与安全原格式读写；超高行不拆分、复杂表格仍只读，保留六格式与多语言正式版门槛。见 docs/M2-WRITER.md。

第五批验收：286cf545 的 145/145 核心回归与独立表格 DOCX 互读已通过；最终边距顺序修正 f72128a9f5c910811429e364399f8ee7487d7452，CI https://github.com/xzcode/office-ex/actions/runs/45434702537，实质 check/interop/license 步骤已通过，记录时 verify-writer: in_progress。下一批先核对最终 native，再开发 PNG/JPEG 图片资源封装、布局与 DOCX 原格式；如有实际失败先修复。保持当前 Writer 分支，未合并 master。

M2 集成第六批候选：增加图片原始字节/摘要、受限 PNG/JPEG 插入缩放删除、OEX 图片封装、独立内嵌图片段落 DOCX 原格式、页面设置及段落 metrics。11 locale 199 键。当前仍非 M2 完整退出：普通 DOCX、分页所见即所得、复杂脚本/多节/图文混排及参考 Office 和 M2-L DOC 等未达原规范；保持当前开发分支及所有 P0 门槛，先核对本轮 CI 结果和验收缺口，后续逐项收敛。

本轮 M2 集成受测源码 2ce8f4599a7e5d0aed2482e510df37f347c20cea，[CI 45437070146](https://github.com/xzcode/office-ex/actions/runs/45437070146) 基础 152/152 成功、独立图片 DOCX 互读及许可证检查通过；记录时 verify-writer: success; native-other (ubuntu-24.04): in_progress; native-other (windows-2025): in_progress; native-macos: queued。初始 CI 的事务权限失败已修复并新增拒绝越权用例。实际 M2 退出缺口表见 docs/M2-WRITER.md；未完成完整 M2/M2-L，未合并 master。后续不再把图片/页面小功能当作整个阶段收尾，按剩余 WRT-01–04 验收条件集中实现，继续当前 Writer 分支。

第七批候选实现：跨相邻段落原子范围编辑、连续文字模式、混合 text/drawing run 与 cell 图片、原子逻辑位置、图像行盒、源媒体保存后撤销修复。204 键/locale。先核对新 CI；M2 退出标准仍按 04-writer/09-i18n，不宣称普通 Word 全兼容或完整 WYSIWYG。M2-L DOC 未达验收。

第七批继续整合下一页分节、独立页面参数与 doc language/direction 元数据，11 locale 211 键，补独立 python-docx 修改后的导入/编辑/保存夹具。首轮原型比较失败已修正。真实双向排版、完整 WYSIWYG、增量重排、更多普通 Office 文件依赖和 M2-L DOC 仍为必需退出门槛。

第七批最终受测 3e0a60f62d543d3f77cde6847f7fc0ee36f46f32；CI https://github.com/xzcode/office-ex/actions/runs/45527322279 基础 162/162、独立混合图片/分节/语言往返及 Python 修改后再导入互通均通过。记录时 verify-writer: success; native-other (ubuntu-24.04): in_progress; native-macos: in_progress; native-other (windows-2025): in_progress。继续时核对实际原生结果并优先修复；集中完成分页富文本编辑、复杂文字 shaping/RTL 和普通 Word 兼容，随后 M2-L DOC。M2 整体未 verified，未合并 master。详见 docs/M2-WRITER.md。

2026-11-25 M2 集中收尾候选：基线 1a2f229，上一轮 CI 45527322279 原生三平台全通过。本轮页面富文本模式、Chromium 原生塑形/视觉簇、节和段落增量缓存、富格式剪贴板、普通默认模板 DOCX 和封闭 Word97 DOC 适配、独立转换会话均已实现。首尔实际 167/167、原有及新增独立互通、构建、许可证通过，另有独立 Chromium 混排页面输入证据。11 locale 221 键。详细实现及边界见 docs/M2-WRITER.md 第八批。M2/M2-L 仍需新版实际原生验收、历史 DOC 更广泛语义结构和参考 Word 排版等价证据；不得将窄 DOC 子集、正文预览、浏览器塑形视为完整 Word 97–2003 兼容。继续保持当前 Writer 分支；无正式 release。

## M3 source recovery (2026-11-25)

The previous M3 source tree was not found in the accessible workspaces or remote branches. Preserve browser output and raw measurement on recovery/m3-20261125 (b4d84c9); continue source reconstruction on m3-sheets-recovery-20261125. See [M3-RECOVERY.md](M3-RECOVERY.md) for scope, verification, and missing source. This first batch restores the typed formula lexer/parser/serializer and locale rules, not the complete calculation engine or M3. The M2 Writer and master branches are untouched.

M3 recovery batch 2 (2026-11-25): restored typed formula evaluator, sparse range/dependency indexes, topological and incremental recalculation, cycle and cache isolation, cancellation/budgets. Formula parser token-kind defect corrected. Local typecheck/architecture and 21/21 targeted tests passed. See docs/M3-RECOVERY.md; next recover spreadsheet commands, formatting, virtual grid/workspace, clipboard and worker integration. M3 and compatibility/release acceptance remain incomplete.

## 2026-11-25 — M3 recovery batches 3–5

Continue from branch m3-sheets-recovery-20261125 and docs/M3-RECOVERY.md. Candidate source for all five recovery batches is now present. Full node regression 198/198, build and dependency audit passed; independent spreadsheet readers and nine headless grid checks passed. The M3/M3-L acceptance gate remains OPEN: scroll p95 50 ms vs <=33 ms goal, narrow XLS/XLSX semantic profiles, real Excel equivalence and three-platform native smoke pending. Do not mark these as formally verified or release. Master/M2 history remains unchanged. Next product development milestone is M4 Slides; complete release-blocking M3 acceptance before claiming full compatibility or releasing.

## 2026-11-25 — M4 source recovered and continued

Continue on `m4-slides-pdf-20261125`, based on `8273c89d3e16d04cbcc9ab613b085040016e9460`. M4 candidate source recovered from the preserved Seoul archive; read `docs/M4-SLIDES-PDF.md`. Eight pinned OFL fonts restored via `npm run setup:fonts`; 236/236 local regressions, independent PPTX/PPT fixture checks, build and license audit passed. Native headless Electron and PDF export verified in part; M4/M4-L acceptance remains open. Do not merge or release before closing remaining compatibility/security/performance/multilingual/native gates.

M4 continuation: 12 native functional checks passed before a final PDF.js stream exception assertion. Corrected rich-text run comparison and PDF stream cancellation Error reason; rebuilt and passed focused PDF cancellation 3/3 with zero exceptions. Full native rerun was blocked earlier by headless image insertion timeout; see `docs/M4-SLIDES-PDF.md` and `docs/verification/m4/pdf-cancel-headless.json`. Keep M4 acceptance open, then address full native repeatability, timing and reference Office.

M4 follow-up: fixed Slides image/text revision race and stale slide selection; 2/2 targeted browser checks; full headless Chromium Slides suite 14/14, 30-slide model-mount median 64.9 ms/max 87 ms, drag p95 16.8 ms/max 100 ms; complete Linux native headless M4 12/12 with no renderer exceptions. Preserved original font-license bytes on Windows checkout. Added dedicated M4 checks across the existing native CI jobs before broad smoke; new runner results pending. Prior CI run 45570279429 had core green, Windows font SHA failure and macOS/Linux broad Sheets smoke timeouts. Review `docs/M4-SLIDES-PDF.md` and the two new JSON evidence files. No merge or release yet.

M4 integration follow-up: on CI 45570900474, core/Ubuntu complete jobs green, Windows M4 native 12/12 and PDF cancel green, Windows broad Sheets timeout; macOS M4 native running at record. Reproduced the actual sheet address draft reset on repaint (B1 became A1); fixed draft preservation and live-workspace navigation after pending cell edits remount the grid. Focused Chromium regression 2/2, typecheck/build passed; native broad smoke now asserts address stability. See `docs/verification/m4/sheets-address-regression.json`. Continue dedicated CI validation and reference Office acceptance; do not treat earlier core/Ubuntu results as all-platform verification of new source.

M4 PDF follow-up: source 341cc7e local broad native 31/31; latest run 45571789798 core and Windows complete green, Linux M4 native legacy reopen/edit/save timeout, macOS in progress at record. Hardened native harness tab-close/open synchronization and timeout diagnostics; local M4 native retest 12/12. Full Chromium PDF suite 11/11 and independent pypdf/Poppler checks pass: Writer four A4 pages, 80 numbered lines exactly once, Slides three 960×540 pt pages, embedded images and Unicode fonts; encrypted input rejected and no external requests or browser errors. Added pinned pypdf test-only dependency and four JSON evidence files in `docs/verification/m4/`. New CI still required; reference Office and physical platform acceptance remain open.

M4 external compatibility: run 45572430061 (a6083f8) core/Ubuntu/Windows complete green; macOS pending at record. Added standard dependency-free PPTX core-properties support with exact native save preservation. Targeted format tests 18/18 and independent interop 7/7 pass; independent python-pptx editing of a supported template and metadata roundtrip verified. Default Office template inheritance/metadata and actual PowerPoint rendering remain outside verified profile. Do not mark M4 fully certified or merge/release on these partial compatibility results.

M4 follow-up: 45574589350 (736cc30) core/Windows/Ubuntu green; macOS recorded all 12 M4 functional PASS lines but timed out at teardown. Extended safe PPTX master/layout profile (solid inherited backgrounds, exact inactive layout retention, per-slide layout identity preserved across edit/reorder/add); block legacy conversion of multi-layout imported sources. Format 19/19, interop 7/7. Fixed explicit Quit intent across asynchronous native document confirmation, cancel IPC clears pending quit, harness uses bounded orderly quit. Linux M4 native 12/12 + orderlyQuit true; dirty Quit cancel/confirm 2/2; IPC 11/11. Latest evidence and limits in M4-SLIDES-PDF.md. New CI pending; active master/placeholder inheritance, richer default templates and actual reference PowerPoint remain open.


M4 consolidated template batch: previous source 4fe77c5 passed all four Actions jobs in 45577967745 (including macOS). Expanded standard-template PPTX master/layout/placeholder inheritance, source metadata/theme and ordinary editable shape/connector profile, soft breaks and line spacing. Fixed consecutive imported PPTX saves after reorder using the main-owned immutable import baseline and existing cache/authority boundaries. Seoul 240/240 core, 9 independent editable fixtures + 2 protected profiles, typecheck/architecture/build/license checks, 13/13 real Electron native checks with orderly Quit passed. See docs/M4-SLIDES-PDF.md and docs/verification/m4/native-template-batch.json. New-source CI, real PowerPoint rendering/legacy profile breadth and physical OS acceptance still required; M4/M4-L not fully verified. Continue on m4-slides-pdf-20261125; no master merge/release.

M4 consolidated paragraph/fit batch: previous a5d8900 passed all four CI jobs (45585817539, all three OS). Added bounded multilevel inherited paragraph/list model, character bullets, decimal numbering, selected-paragraph commands, editor/thumbnail/presentation/PDF marker projection, normal-autofit stored parameters and explicit measured Fit Text/reset. All 22 locale catalogs extended. 243/243 core, 37/37 focused Slides, 11 editable independent source/output fixtures + 2 protected external profiles, 16/16 Chromium Slides UI, 11/11 PDF browser plus independent numbered text/embedded-font inspection, 14/14 real native M4 with orderly Quit. New source requires three-platform CI. Read docs/M4-SLIDES-PDF.md and docs/verification/m4/ paragraph-fit evidence. Dynamic shape expansion, full reference autofit, richer numbering and actual PowerPoint compatibility remain outside verified profile. Do not close M4/M4-L, merge or release on partial reference evidence.

M4 consolidated graphics batch: 7638e95 passed all four CI jobs (45587859248). Implemented bounded nested editable groups and group transforms, reversible contiguous grouping/rigid ungroup, source-preserving image crop, four default vector presets, flips, straight free connectors, preset dash and arrow styles, recursive media and identity handling. UI/22 locales extended; source/native PPTX save and legacy rejection protected. Seoul 246/246 core, 40/40 Slides, independent interop 13 editable samples + 3 protected, 19/19 Chromium; typecheck/architecture/build/dependencies passed. See docs/M4-SLIDES-PDF.md and docs/verification/m4/slides-graphics-browser.json. New-source CI/reference playback and grouped PDF checks open; scaled/flipped ungroup, attached connectors, custom geometry, unsupported effects and active grouped placeholder roles not established. Do not merge/release or call M4 complete until actual acceptance gates are addressed.

M4 nested editor/connector/fit batch: previous 7c1b227 all four CI jobs passed (45590672876). Implemented group-local nested object editor, straight two-ended rectangular attachments, live endpoints, ID remapping/detach on delete/clipboard, native PPTX binding preservation, dynamic normal font fit in same transaction. Seoul full core 249/249, focused 43/43, independent interop 15 editable source/output + 3 protected, Chromium Slides 22/22, PDF browser 11/11 and independent grouped-text/images check. Typecheck/build/architecture/license checks passed. New evidence and boundaries in M4-SLIDES-PDF.md. No master merge/release; latest native CI and actual Office acceptance pending. Avoid claiming arbitrary connectors, automatic shape growth or reference PowerPoint layout equivalence.

M5 batch 1, branch m5-distribution-20261125 from M4 3b49ace (CI 45595130602 all four jobs passed). Added dependency-free native preview packaging (platform-matched archives and optional DEB/DMG/Inno Setup installer), file/hash manifest, packaged-app real launch verifier and gated CI artifact upload; explicit formal-release ledger/readiness check; updated README. Repaired blank DOC paragraph-mark style canonical reconstruction. Writer targeted 6/6, TypeScript, build, architecture, packaged Linux 10/10 including all six format save/reopen and empty DOC; evidence docs/verification/m5/linux-packaged-preview.json records sourceClean:false. New native installer CI and installation evidence, reference Office, full performance, human language QA remain open. Continue M5 on its branch; no master merge or release. Local workspace remains /app/agent/office-ex-m4-20261125 (M5 files added); archive out/m5-preview. Review docs/M5-DISTRIBUTION.md and docs/verification/m5/release-gates.json.

M5 batch 2: prior 7db8704 CI 45597866933 core/Windows/Linux passed, macOS DMG built but extracted app.isPackaged false. Fixed macOS executable/Info.plist identity and added gated native installer install/hash/launch/remove checks on disposable CI, Linux root-owned sandbox helper in DEB only. Sheets viewport node/style/header reuse and bounded drawing layer implemented, 12/12 browser checks, final 100k constructed-model first grid median 690.4 ms, scroll-to-DOM p95 11.5 ms (before 14.7), continuous scroll p95 50 ms still above 33 ms; do not mark performance verified. All 250 core and 31 Linux native checks passed, typecheck/build/architecture/licenses pass. See docs/M5-DISTRIBUTION.md and docs/verification/m5 raw reports. New-source native packaging/installer CI pending. Reference Office/locale human review/performance formal gates remain. No release/master merge.

M5 batch 3: cb1a161 CI 45599395641 all four jobs succeeded, including exact source three-platform archive/installer build, installation/copy, sandboxed launch, uninstall/removal, and uploaded development-preview artifacts. Expanded XLSX ordinary template import, theme/indexed colors, passive metadata/view/print settings, stable source sheet graph and append-only style mapping; byte-exact no-edit XLS/XLSX and same-session source baseline for imported XLSX/PPTX. Rich-template legacy conversion blocked. Seoul 252/252 core, 9/9 focused spreadsheet format, independent openpyxl two saves + three protected profiles, real native 32/32 including consecutive reorder/save/edit/save. Word Heading 1/2 independent edit and byte preservation passed; complex Title remains protected. Read docs/M5-DISTRIBUTION.md, docs/verification/m5/native-desktop-batch3.json. Continue M5; new-source CI, actual reference Office/legacy coverage, QA-05, locale human review and trust remain formal gates. No master merge/release.

M5 batch 4: 2b74917 all four CI jobs passed (45601554499), including three-platform installer lifecycle. Expanded bounded Word Title style, character spacing, kerning threshold, right/first/hanging indent, contextual spacing and bottom paragraph rule, native/Canvas/page/PDF projections, format commands and inspector, candidate labels in 22 locales. Native DOCX preserves named style/theme source and blocks unsupported side rules/legacy conversion. Seoul core 255/255, Word focused 25/25, independent default-template Title edit/reopen and non-main bytes, native desktop 33/33, PDF browser 12/12 and independent text/images/fonts verified. See docs/M5-DISTRIBUTION.md and docs/verification/m5/batch4 and word-paragraph reports. Local Antiword absent, full locked independent interop and native installer CI must run on new source. Keep formal Office reference/layout/legacy, performance, trust and human QA gates open. No master merge/release.

M5 batch 5: d106214 CI 45611433927 all four jobs passed. Implemented bounded Word decimal/bullet numbering registry, multilevel references, source start overrides, native DOCX append-only numbering preservation, list inspector commands, clipboard remapping, global layout counters and derived marker geometry across Canvas/native/PDF. Main retains immutable DOCX import baseline as for XLSX/PPTX. Seoul 258/258 core, 3/3 focused numbering, independent python-docx default named/multilevel/restart corpus and protected Roman profile, PDF browser 13/13 independent marker/text/fonts, focused native list save passed. See docs/M5-DISTRIBUTION.md and docs/verification/m5/word-numbering-pdf reports. Exact Word typography, unsupported numbering formats/RTL/picture bullets, legacy DOC numbering and full formal release gates remain open. No master merge/release.
Final batch 5 native desktop 34/34; raw reports docs/verification/m5/native-desktop-batch5.json, docs/verification/m5/word-numbering-pdf-browser.json, docs/verification/m5/word-numbering-pdf-independent.json. New-source CI pending.

M5 batch 6: previous 877d72e all four CI jobs passed (45615295882). Added original geometric EX mark and deterministic native branding generator, actual window/Dock/Bundle/Windows PE+installer/Linux desktop icons, native-format-first home with internal drafts collapsed, runtime version/About, 22 locale labels, exact-source draft-only release assembly on explicit development tag. Local 259/259 full core/tooling, 8/8 locale checks, build/typecheck/architecture/licenses, Pillow icon decode and actual Linux branded archive 11/11. Evidence docs/verification/m5/linux-branded-preview.json and docs/verification/m5/branding-assets.json; local sourceClean:false. New-source native Windows/macOS/DEB branding/installer CI and real tag-triggered assembly pending. No tag/public release/master merge; only a deliberately pushed matching dev tag after authorization may create a draft. Mark/home chosen provisionally, no visual approval or trademark check. Keep formal gates open.

Post-dev.1 compatibility hotfix: real user XLSX screenshot fell back to structure inspection (original workbook not provided). Bounded Excel extended sheet-title properties import/update; concrete semantic diagnostic; independent XLSX raw-value preview and DOCX/PPTX text/table outline fallback, no semantic save or recalculation. Fixed common independent Word inline image and PPTX picture optional geometry; top-level unrotated top-anchored spAutoFit textbox import/serialization and atomic UI geometry growth. Ordinary default Word auto TableGrid and header/footer, PPTX tables still protected, exact reference Office layout not established. 262/262 core, independent external three-format samples and output readback, native Linux focused 7/7 passed. See docs/M5-DISTRIBUTION.md and docs/verification/m5/office-external-compatibility-hotfix.json. dev.1 public release source 398f1a4 unchanged; new CI pending and no new release published.

Post-dev.1 UI redesign: user-approved Office-inspired direction implemented with shared original SVG ribbon, per-editor groups/tab state, compact shell, Writer default native page/rail/ruler/zoom and alternate text modes, Sheets organized cell/formula/view tools, Presentation contextual collapsible inspector and stable viewport fit. Candidate labels in 22 locales. See docs/UI-RIBBON.md and docs/verification/m5/ribbon for actual native screenshots and raw reports. Local 262/262 core/tooling, 4/4 ribbon native, 7/7 external compatibility native, independent presentation 18 cases passed (obsolete negative spAutoFit fixture corrected to rotated). New-source native CI pending. Published dev.1 unchanged; do not claim formal compatibility, Microsoft visual equivalence or human localization acceptance.

Final full Linux native desktop regression: 34/34 passed on the redesigned UI, including alternate Writer mode navigation, image/table/list saves, crash recovery, and protected source preservation. See docs/verification/m5/ribbon/native-desktop.json.

Workbench V2: minimal user-approved home, single application titlebar, original blue/green/orange icons, recent names/optional last-use timestamps (no renderer paths), real recovery, secondary formats folded. macOS titleBarStyle hidden, original native red/yellow/green controls positioned 14,18, physical left safe area 88px, draggable noninteractive header. No fake OS window controls. Windows/Linux native frame unchanged. Stable live header controls fix prior 03f48f3 CI extracted-app save failure reproduced on immediate PPT text save; local packaged Linux 11/11 pass including all six formats, home native 5/5, core 262/262, settings/i18n 16/16. See docs/UI-RIBBON.md and docs/verification/m5/home-v2. Local packaged sourceClean:false, before final cosmetic refinement; Mac CSS-only local check, new native CI required. Public dev.1 release unchanged.

Final Linux native desktop regression: 34/34 passed on the final Workbench V2 code. Raw report: docs/verification/m5/home-v2/native-desktop.json.

## 2026-11-26 — dev.2 release preparation

Continue on `m5-distribution-20261125`. Workbench V2 and compact original ribbon UI at exact source `e63d20f579af01f056b41f4e31d24de2e73ea17b` passed [Actions run 45765424759](https://github.com/xzcode/office-ex/actions/runs/45765424759): core/interop/licenses and Windows x64, macOS ARM64, Linux x64 native packaging, extracted launch, installation and removal. This evidence applies to the dev.1-versioned candidate source, not the new dev.2 build. The public v0.1.0-dev.1 tag and release remain unchanged.

Prepare v0.1.0-dev.2 by synchronizing root and all 14 workspace versions, internal references, package lock and dependency-inventory lock digest. External dependency entries and reviewed notices are unchanged. Add regression checking version and inventory consistency. See [dev.2 release checklist](releases/v0.1.0-dev.2.md). Require new exact-source native CI before creating an explicit matching tag and release draft. Do not overwrite dev.1, merge master, or claim formal Office compatibility. No tag or public release is created by this preparation commit.

## 2026-11-26 — dev.2 public preview released

[v0.1.0-dev.2](https://github.com/xzcode/office-ex/releases/tag/v0.1.0-dev.2) is public (prerelease, not a formal release), release ID 459062473, published 2026-11-26 07:05:48 UTC. The lightweight tag targets exact verified source `e0a827fa61a9643aa2ba4fc0b62edbd2aedacd42`. [Native CI 45770026211](https://github.com/xzcode/office-ex/actions/runs/45770026211) completed core/interop/license checks and Windows x64, macOS ARM64, Linux x64 extracted and installer lifecycle verification. [Publication run 45772396449](https://github.com/xzcode/office-ex/actions/runs/45772396449) passed, reused those three verified artifacts, and verified checksums of all 20 release assets before publication. Public REST readback confirmed non-draft prerelease, 20 uploaded assets and exact tag target. The dev.1 release remains untouched. The source tag is immutable for this release; subsequent documentation commits do not represent the built source. The assembly evidence's draft-only field records its pre-publication stage; the public API confirms actual publication separately.

Recommended installers: macOS ARM64 DMG (151465933 bytes), Windows x64 EXE (129223865 bytes), Linux x64 DEB (112847386 bytes). Windows unsigned; macOS ad-hoc signed, not notarized. No claim of formal six-format Office compatibility or human locale review. Continue on m5-distribution-20261125, not the one-shot publication branch. Next priorities: ordinary Word auto-layout Table Grid / headers and footers, spreadsheet external profiles and editing fidelity, presentation table / group profiles, based on reproducible independent source fixtures and safe preservation. Keep formal gates open; do not repackage unchanged dev.2 or overwrite either published tag. Master remains unchanged.

## 2026-11-26 — M6.1 batch A

Continue on `m5-distribution-20261125`, see [M6 compatibility](M6-COMPATIBILITY.md). Ordinary default-template TableNormal/TableGrid rectangular DOCX tables now admit bounded source-owned native styles, auto width/autofit and consistent look flags. Retain native semantics, explicit overrides, style spacing and ancillary parts; stored-grid projection is not Word content-driven autofit. Final local 263/263 core, independent producer/readback, desktop build, dependency audit and actual Linux sandboxed native 3/3 passed; new-source CI still required. Merged/nested tables and complex style/geometry remain protected. dev.2 release/tag and master unchanged. Continue M6.1 with richer tables before M6.2 headers/footers; all formal M5 gates remain open.

## 2026-11-26 — M6.1 batch B

Continue on `m5-distribution-20261125` from [M6 compatibility](M6-COMPATIBILITY.md). Prior ordinary-table source 85ee488 passed all three native platforms. New merged-cell grid validation, native gridSpan/vMerge read/write, owning-cell layout, actual merge/split ribbon controls and all 22 candidate locale labels are implemented. Final local 267/267 core, targeted 16/16, independent Office corpus, build, licenses and actual Linux sandboxed native 4/4 passed. Structural row/column insert/delete is disabled in merged tables until split; split retains all content in top-left. Content-bearing continuations, nested tables and complex styles remain protected; cross-page merged-row splitting and exact Word autofit are not established. New-source CI required. dev.2 release/tag/master unchanged. Continue M6.2 headers/footers/page numbers and related safe preservation.

## 2026-11-26 — M6.2 batch A

Continue on `m5-distribution-20261125`; see [M6 compatibility](M6-COMPATIBILITY.md). Prior merged table source 59e8905 passed three-platform native CI 45775804320. Implemented bounded default DOCX running header/footer stories, simple/complex decimal PAGE fields, section inheritance/detached editing, restart/continue, page edge distances, original part preservation and native dialog/page projection. Final local 272/272 core, targeted 8/8, independent corpus, build/licenses and sandboxed Linux native 3/3 passed. PDF projection integrated but dedicated running-story PDF verification not yet executed. Multiple paragraphs, first/even variants, non-PAGE fields and complex story formats remain protected. New-source CI required. Published dev.2 and master remain unchanged.

## 2026-11-26 — M6.3 batch A

Continue on `m5-distribution-20261125`; see [M6 compatibility](M6-COMPATIBILITY.md). M6.2 source 587a44c passed CI 45778305240 (verify plus three native package jobs). Implemented paragraph keep-next (including following table row group), keep-lines, widow/orphan control, explicit paragraph page break and native ribbon controls. Final local 280/280 core, 33/33 targeted, build/licenses and actual sandboxed Linux native 3/3 passed. Independent pypdf readback verifies three PDF pages, header and PAGE 5/6/7, correct paragraph placement; source preserved. Word pagination is bounded, not exact reference Office fidelity; oversized constraints may relax with warnings. New-source CI required. Public dev.2 and master unchanged. Next M6.4 spreadsheets.

## 2026-11-26 — M6.4 / M6.5 batch A

Continue on `m5-distribution-20261125`; see [M6 compatibility](M6-COMPATIBILITY.md). XLSX independent cell border edges/line styles and explicit default row/column geometry are now admitted, projected and preserved on native save. Formula registry expanded 12 → 38 (26 new functions); bounded sparse COUNTBLANK, lazy IFERROR, numeric/text/predicate functions. Local final 284/284 core, 34/34 targeted, independent full Office corpus/readback, build/licenses and sandboxed Linux native 3/3 passed. M6.3 source a205f36 verify/Windows/macOS passed; Linux package job still in progress at this checkpoint. New source CI required. Rich text/named styles, conditional aggregates, lookup functions, charts and conditional formatting remain future work; exact Excel geometry/typography and full BIFF8 compatibility are not established. Published dev.2 and master unchanged.

## 2026-11-26 — M6.5 batch B

Continue on `m5-distribution-20261125`; see [M6 compatibility](M6-COMPATIBILITY.md). Formula registry expanded 38 → 48 (conditional aggregates, VLOOKUP/HLOOKUP, MATCH and scalar INDEX). Sparse same-shape criteria ranges, bounded wildcard matching, exact and sorted numeric approximate lookup, cross-sheet dependency updates and native cache preservation. Local 288/288 full, 38/38 targeted, independent Office corpus and eleven lookup/conditional caches, build/licenses and sandboxed Linux native 3/3 passed. Exact-source CI required. Source 9e4c53c verify/Windows/macOS succeeded, Linux pending; older M6.3 run cancelled. Unsupported implicit output range expansion, array INDEX, mixed/gapped/unsorted approximate lookup, XLOOKUP and new legacy XLS function serialization remain outside profile. Public dev.2 unchanged. Next M6.6 charts and conditional formats.

## 2026-11-26 — M6.6

M6.5 B baseline `4c11eaa1a0adcc8ba38d8db102c1beb462efa435` passed exact-source CI [45793969267](https://github.com/xzcode/office-ex/actions/runs/45793969267).

Numeric conditional rules and three native chart types with source-backed preservation, creation/deletion and live projection. Local 290/290, independent Office corpus, build/licenses, sandboxed native 3/3. See M6-COMPATIBILITY.md and docs/verification/m6/native-charts.json. Unsupported complex profiles guarded. Continue M6.7/M6.8. Exact-source CI required; published dev.2 unchanged.

## 2026-11-26 — M6.7–M6.10 candidate

Native rectangular PPTX table and distribution commands, shared scene/ribbon, double-click/Enter interaction and unchanged relationship preservation. Focused slides 46/46; independent external table and full Office corpus passed; final Linux sandboxed native 3/3. See M6-COMPATIBILITY.md and native-ppt-table.json. Prepare dev.3 version/lock without third-party dependency changes. Exact dev.3 CI, publication and broader acceptance remain pending; dev.2 release and master unchanged.
