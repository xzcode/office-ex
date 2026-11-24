# M1 验收记录

**结论：M1.1–M1.10 基础设施阶段通过。** 验收日期：2026-11-24；开发版本：`0.1.0-dev.1`。

- 受测提交：[7d0522dcd2a8f6db676a6878429da2ca31a839f6](https://github.com/xzcode/office-ex/commit/7d0522dcd2a8f6db676a6878429da2ca31a839f6)。
- 受测 tree：`25ea2dd3dfadad3c161798f36fef7896eb64fcc8`。
- 实际运行：[GitHub Actions 45375454826](https://github.com/xzcode/office-ex/actions/runs/45375454826)，四个任务均为 `success`。
- 机器可读结果：[report.json](report.json)。包含提交、任务、环境、测试计数、原始产物标识／SHA-256、归档文件摘要、合成夹具源代码摘要和限制。

本目录是在验收后补充的文档与证据。受测版本以以上源码提交为准，不能将文档提交 SHA 描述成重新执行过测试的代码版本。截图与三个 harness JSON 从该轮 Actions artifacts 原样归档，下载时已核对每份 ZIP 的服务端 SHA-256；不是生成的界面示意图。

## 基础检查

| 检查 | 实际结果 |
|---|---|
| 严格类型与架构 | 通过；12 个 workspace、29 个源文件，公共入口、显式依赖、无环、纯领域和 worker 边界 |
| 自动化回归 | 99 项通过，0 失败，0 跳过 |
| ZIP 独立互读 | 原创 ZIP 输出由 Python `zipfile` 读取，Python DEFLATE 输出由原创读取器解析 |
| OOXML 原型语义检查 | 原 M0 DOCX/XLSX XML 输出由 `python-docx 1.2.0`／`openpyxl 3.1.5` 读取 |
| CFB 独立互读 | `olefile 0.47` 读取 v3/v4；目录、小流、空流、4096 边界、8MiB DIFAT；独立修改小／大流后原创读取器复读 |
| 依赖与许可 | 66 个锁定 npm 包、23 份安装包 notice 审查通过；原生构建实际附带 Electron／Chromium notices |
| npm advisory | 2026-11-24 查询报告 0 已知漏洞，不等同于全面安全审计 |

基础环境：Node 24.22.3、npm 11.19.1、Python 3.12.14。原始输出节选见 [verify-checks.txt](verify-checks.txt)。夹具为本项目合成数据，没有把用户文档加入仓库。旧格式识别测试中的有限头部样本仅验证识别器，不是完整 DOC／XLS／PPT 互操作样本。

## 三平台原生结果

实际运行 Electron **45.4.1**、Chromium **156.0.8078.100**，开启 sandbox／context isolation／web security，renderer 和 worker 禁用 Node integration。架构取进程的 `process.arch` 和实际 runner 信息，不依据浏览器 User-Agent 的历史兼容字符串推断。

| 实际平台 | 运行时相关回归 | 原生端到端 | 证据 |
|---|---|---|---|
| macOS 15.8.2，ARM64 | 43 通过，0 跳过 | 15／15 通过 | [JSON](native/darwin-arm64.json) · [输出节选](macos-checks.txt) |
| Windows Server 2025，x64 | 42 通过，1 跳过 | 15／15 通过 | [JSON](native/win32-x64.json) · [输出节选](windows-checks.txt) |
| Ubuntu 24.04.5，x64 | 43 通过，0 跳过 | 15／15 通过 | [JSON](native/linux-x64.json) · [输出节选](linux-checks.txt) |

宿主回归是 99 项中的相关子集，在三个 OS 重新执行，不应累加成更多不同测试。Windows 的设置文件软链接用例在源码中显式跳过；文件服务和恢复的链接测试实际执行。该跳过项没有算作通过。Windows Server 结果不能代替 Windows 10/11 安装验证。

原生端到端检查覆盖：窗口隔离与无 Node、远程／本地 fetch 与弹窗拒绝、伪造能力与伪造拖入拒绝、草稿序列化和原子保存、撤销回到已保存状态、composition 单次事务、三个工作区独立历史、保存期间继续编辑仍 dirty、覆盖前备份、DOCX 结构检查与原字节副本、错误扩展名保存保护、真实 30 秒恢复、进程退出后恢复副本、设置持久化及脏关闭取消／放弃。

原生文件对话框返回值由 harness 注入，输入法采用 DOM composition 事件模拟。这些测试不代表已人工操作所有系统选择器或输入法。进程退出／注入 IO 故障也不等于物理断电试验。

## 原生截图

三平台各归档四张截图。首页、深色草稿、只读检查的主要控件与文字均可见；Slides 截图保留点击“添加矩形”后的滚动位置，其中 macOS／Windows 的画布上缘位于视口外，Linux 的画布完整可见。深色截图保留故意触发的错误扩展名提示，属于保护流程证据。

| 平台 | 首页 | 深色草稿 | 结构检查 | Slides 草稿 |
|---|---|---|---|---|
| macOS | [查看](native/darwin-arm64-home.png) | [查看](native/darwin-arm64-draft-dark.png) | [查看](native/darwin-arm64-inspection.png) | [查看](native/darwin-arm64-slides.png) |
| Windows | [查看](native/win32-x64-home.png) | [查看](native/win32-x64-draft-dark.png) | [查看](native/win32-x64-inspection.png) | [查看](native/win32-x64-slides.png) |
| Linux | [查看](native/linux-x64-home.png) | [查看](native/linux-x64-draft-dark.png) | [查看](native/linux-x64-inspection.png) | [查看](native/linux-x64-slides.png) |

![macOS 原生工作台](native/darwin-arm64-home.png)

## 集成中修复的实际问题

Windows 的自动换行转换改变了 lock 和许可文件的字节，触发审核摘要不一致。`.gitattributes` 为 lock／inventory 固定 LF，对原始 notices 禁止转换，保留上游本来的 LF 或 CRLF；没有放宽摘要核验。

Linux 首次截屏早于 compositor 首帧，且测试机没有中文字体。harness 现在等待原生窗口可见和绘制帧，仅对特定截屏错误最多重试三次；CI 安装系统中文／emoji 字体。最终三平台均未用到截屏重试。Undo／Redo 断言也等待实际重新渲染，避免将 IPC 尚未返回时的旧 DOM 当作最终结果。

首次失败运行与中途替换的运行保留在报告中，不作为最终验收证据。最后一轮成功针对上述同一个受测源码提交。

## 能力边界及后续门槛

M1 的详细范围见 [M1.md](../../M1.md)；逐项对应关系见 [features.json](../../specs/features.json)。宽泛需求的 `status` 与本阶段 `m1Acceptance` 分开记录，后续义务没有删除。

目前 Office 文件只能进行容器／有限版本标记检查和原字节另存；正文、单元格、幻灯片还没有导入为可编辑 Office 语义。`.oex` 是内部草稿，且暂时拒绝尚未封装资源字节的模型。公式仅保存源码；Canvas 字簇位置标注为 estimated。M2／M2-L、M3／M3-L、M4／M4-L 和 M5 的六格式读写、往返、转换、真实 Office 对照与安装发行仍需分别验收。

文件提交保留 Node 最终身份检查与 rename 之间的外部竞争边界；目录同步不可用时显示较弱保证，Windows 截图中可见该提示。未将未知网络盘、所有 OS 版本、所有存储设备、OS 级零流量或全面安全审计描述为已经验证。当前没有发布安装包、签名／公证产物或正式 Release。
