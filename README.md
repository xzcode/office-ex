# Office EX

自主研发、跨平台、离线优先的开源办公套件。原创代码采用 MIT 许可证；文档模型、命令、撤销、格式语义、排版、公式和幻灯片场景由项目自己实现。通用开源依赖负责桌面运行时、压缩和 XML 等基础设施。

## 当前进度：M1 基础设施已完成

M1.1–M1.10 已完成并通过阶段验收：Electron 桌面工作台、多标签会话、共享模型与事务、ZIP/XML/OPC 与 CFB 容器、安全保存、备份、恢复、主题和中英文设置均已落地。源码 `7d0522d` 在 [同一轮 CI](https://github.com/xzcode/office-ex/actions/runs/45375454826) 通过 99 项基础回归，以及 macOS ARM64、Windows x64、Linux x64 各 15 项真实 Electron 检查。完整环境、平台测试计数、跳过项和截图见 [验收记录](docs/verification/m1/README.md)。

**目前是 `0.1.0-dev.1` 开发预览，没有发布正式安装包。** 已验收的是 M1 基础设施范围，完整 Office 兼容能力继续按后续阶段实施。

当前界面可以：

- 新建 `.oex` 内部草稿，验证文字输入、A1:C3 单元格事务、矩形场景、撤销／重做、保存／另存和恢复。
- 检查受支持 Office 文件的实际容器、格式标记、部件／流、大小、摘要和兼容性诊断。
- 对检查成功的源 Office 文件执行**字节完全相同的原样另存**。

**Office 源文件尚未导入为可编辑文档。** `.oex` 是 M1 会话与保存流程的内部草稿，不能交给 Microsoft Office 打开，也不能通过修改扩展名当作 DOCX、XLSX 或 PPTX。分页、公式求值、完整网格、幻灯片编辑、PDF 阅读／输出和六格式转换仍在后续里程碑。

## 正式版格式目标

首个正式版必须覆盖 **DOC/DOCX、XLS/XLSX、PPT/PPTX** 的约定常用功能子集：新建、读取、编辑、原格式保存、另存、复开和新旧格式双向转换，并提供 PDF 阅读／导出。

旧版目标为 **Office 97–2003，XLS 使用 BIFF8**。共享 CFB 容器已进入 M1；DOC、XLS、PPT 的语义读写分别属于 **M2-L、M3-L、M4-L**，全部是首个正式版的强制门槛。能够识别容器不等于已经实现旧格式编辑。

## 国际化规划

M2-I18N 已编写 A 批 11 套候选界面资源：简体中文、繁体中文、英语、日语、韩语、法语、德语、西班牙语、巴西葡萄牙语、俄语、意大利语。B 批 11 个 UI locale 尚未实现。当前处于开发验证阶段，人工术语校订与正式验收尚未完成，详见 [实施记录](docs/M2-I18N.md)。自 M2-I18N 起实施共享语言包、自动识别系统语言、无需重启的切换、独立地区格式、文档语言与 RTL。首个正式版规划 22 个界面 locale（含简繁中、英、日、韩、欧洲主要语言、阿拉伯语、希伯来语、印地语及东南亚主要语言），并验收 Unicode/复杂文字、公式与 PDF 的语言兼容。完整语言表、实施阶段和门槛见 [国际化规范](docs/specs/09-i18n.md)。这些是新增规划，不是现有功能。

## 本地启动

使用 Node.js **24.13.0 及以上的 24.x 版本**；CI 固定为 24.22.3。首次安装需要访问 npm 与 Electron 官方发行文件下载服务，应用运行时默认离线。

```sh
npm ci --ignore-scripts
npm run setup:desktop
npm run start
```

`setup:desktop` 显式安装锁定版本的原生 Electron；`start` 构建后打开桌面窗口。Linux 需要图形会话、可用的 Chromium sandbox 和中文字体；Ubuntu CI 安装 `fonts-noto-cjk`、`fonts-noto-color-emoji`，使用 Xvfb。系统字体不随应用分发，sandbox 配置见工作流。

## 开发验证

Python 只用于独立文件读取和兼容性夹具，不参与生产运行。下面的命令应在已安装 `tests/requirements.txt` 的 Python 环境中运行：

```sh
python -m pip install -r tests/requirements.txt
npm run check
npm run test:interop
npm run audit:licenses
npm run build
npm run smoke:desktop
```

`check` 包括严格类型检查、模块边界／纯领域检查和自动化测试。`test:interop` 使用独立 Python 读取器检查原创 OOXML 原型和 CFB 输出。`smoke:desktop` 启动真正的 Electron，记录平台、版本、断言结果和截图到 `out/native/`；它不使用 `--no-sandbox`。

## 核心组件

| 组件 | 当前用途 |
|---|---|
| 原创 TypeScript 核心 | 文档模型、事务／历史、资源／样式、几何、格式契约和安全策略 |
| Electron 45.4.1 | 桌面窗口、本地文件授权、隔离 preload、sandboxed renderer／worker |
| fflate 0.8.3 | 通用 DEFLATE 压缩／解压；ZIP 和 OPC 策略由项目实现 |
| @xmldom/xmldom 0.9.12 | 通用 XML DOM；资源预算和安全限制由项目实现 |
| 原创 CFB 实现 | MS-CFB v3/v4 容器读写和流／目录保留，为旧版格式提供基础 |

完整固定版本、传递依赖和许可证见 [依赖审查](docs/DEPENDENCIES.md)与 [第三方 notices](THIRD_PARTY_NOTICES.md)。

## 文档

- [M1 实现、API、阶段和验证说明](docs/M1.md)
- [M1 验收报告与原生截图](docs/verification/m1/README.md)
- [M1 架构决策](docs/decisions/0001-m1-boundaries.md)
- [开发规范 v1.1](docs/specs/README.md)
- [多语言与地区适配规范](docs/specs/09-i18n.md)
- [六格式契约](docs/specs/03-formats.md)
- [功能追踪矩阵](docs/specs/features.json)
- [开发交接](docs/HANDOFF.md)
- [许可证规则](docs/LICENSING.md)
- [M0 历史基线](docs/M0.md)与 [M0 验证](docs/M0_VERIFICATION.md)

规范和追踪矩阵区分计划、实现与验证。所有兼容性声明都需要对应格式、功能、样本和实际环境证据。
