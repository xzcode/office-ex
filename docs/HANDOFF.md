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
