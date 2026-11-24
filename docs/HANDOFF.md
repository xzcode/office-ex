# Office EX 开发交接 — M1

更新日期：2026-11-24。仓库：`xzcode/office-ex`，主分支：`master`。当前正在完成 M1 基础设施整批集成与原生验证；最终提交 SHA 和 CI 结果需由验收记录关联，本文不预先宣布各平台通过。

## 必须保留的决定

用户选择自主开发路线：自研 UI、三个领域模型、编辑语义、布局、公式、场景和格式适配，允许通用底层库。生产依赖没有完整 Office 引擎；不要恢复早期已废弃的 ONLYOFFICE／AGPL 方案、upstream lock 或下载／构建脚本。早期方案仅保留在 Git 历史里，未导入上游 Office 编辑器源码。

原仓库 MIT 许可证保持不变。Electron、fflate、xmldom 等通用依赖按固定版本和各自 notice 管理；见 `docs/DEPENDENCIES.md`、`docs/DEPENDENCIES.lock.json` 和 `THIRD_PARTY_NOTICES.md`。不要把 Electron 的根 MIT 描述成所有 Chromium／Node 组件都只有这一份许可证。

用户明确要求 DOC/DOCX、XLS/XLSX、PPT/PPTX 六格式。旧版目标为 Office 97–2003，XLS 为 BIFF8。M2-L、M3-L、M4-L 是首个正式版必经门槛；M1 的 CFB 容器和格式检查不会替代旧格式编辑与双向转换。

## 进入仓库先读

1. `AGENTS.md`、`docs/LICENSING.md`。
2. `docs/specs/README.md` 与本次相关规范；规范优先于旧路线图。
3. `docs/M1.md`、`docs/decisions/0001-m1-boundaries.md`。
4. `docs/specs/features.json` 的实际状态和证据，以及目标提交对应的 CI 输出。

## 已落地的 M1 代码

- npm workspaces、严格 TypeScript、固定直接依赖和单一 lockfile；公共包入口、依赖图与纯领域编译检查。
- `contracts/core/writer/sheets/slides/render`：三种模型、运行时校验、EMU／Unicode、样式／资源、命令、增量 undo/redo、dirty 和保存票据、派生任务、几何／字体接口。
- `ooxml`：有界 ZIP/XML/OPC；保留原始部件和摘要。原 M0 DOCX/XLSX XML 写入原型仍作为回归证据保留。
- `cfb`：原创 MS-CFB v3/v4 读写、FAT/DIFAT/MiniFAT、目录／storage／stream 和安全检查。DOC/XLS/PPT 语义适配尚未实现。
- `format-common`：有界二进制读取、适配器／能力／诊断／保存计划契约。
- `format-service`：在 sandbox worker 中组合容器检测和内部草稿序列化；不让低层容器互相依赖，不让 Electron main 解析文档。
- `apps/desktop`：本地协议、受限 preload/IPC、窗口／菜单、文件 token、身份检查、原子提交、替换前备份、恢复和设置。
- `ui`：多标签、首页／最近文件、中文／英文、浅／深／系统主题、三个草稿预览、Office 结构检查、取消、脏文档关闭提示。
- CI 和真正的 Electron smoke harness：先做基础检查，再验证 macOS ARM64，之后 Windows/Linux；具体通过状态需查看真实运行结果。

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

2026-11-24 本轮核心／渲染的 21 个定向测试和架构检查已独立通过；依赖审查确认 66 个锁定 npm 包、23 份当前安装 notice 文件，并实际 stage 了 Electron 原生 notices。整批测试统计、独立容器互读、各原生平台与 smoke 截图以最终验收产物为准。未收到该平台成功结果前保持“待验证”。

## 接下来如何继续

先收齐并审查 M1 最终源码提交的基础检查、互读、原生 sandbox、文件保存／恢复和界面 smoke 证据，更新追踪矩阵。不要把提交前或其他 commit 的测试结果当成当前构建证据。

通过 M1 后进入 M2 Writer：建立真正的段落／run 编辑、选区／IME、布局盒和分页，再实现 DOCX 语义映射与意图编辑保存。并行设计 MS-DOC 的 M2-L 子阶段，使用共享 CFB 与 Writer 模型。后续继续 M3/BIFF8 XLS、M4/PPT/PPTX/PDF；正式发布前完成六格式往返、转换、参考 Office 和各平台安装验证。

不要新建 Release、宣称已生成安装包、降低 P0 验收，或把 `.oex` 作为正式版兼容目标的替代。修复与验证应围绕具体剩余风险；不重复无关打包或扩大到未授权项目。
