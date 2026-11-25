# Office EX

自主研发、跨平台、离线优先的开源办公套件。原创代码采用 MIT 许可证；文档模型、命令、撤销、格式语义、排版、公式和幻灯片场景由项目自己实现。通用开源依赖负责桌面运行时、压缩和 XML 等基础设施。

## 当前进度：M5 开发预览交付

**当前版本为 `0.1.0-dev.1` 开发预览，不是正式 Office 兼容产品。** M1 基础设施已验收；M2 Writer、M3 Sheets、M4 Slides/PDF 已实现限定格式子集，M5 正在完成六格式验收、分发和性能工作。

- Writer：段落与字符样式、分页、限定 DOCX 表格／图片／分节、限定 Word 97–2003 DOC 原格式读写。
- Sheets：虚拟网格、稀疏单元格、公式计算、格式、限定 XLSX／BIFF8 XLS 原格式读写。
- Slides：幻灯片与图形、母版／版式的已支持子集、段落列表、主题、嵌套组、裁剪、连接线、限定 PPTX／PPT 原格式读写。
- PDF：离线阅读、搜索、翻页、Writer／Slides PDF 输出及 Unicode 字体嵌入。
- 已接入六种格式的新建、编辑、原格式保存、复开和受支持的双向转换。未知特性或不安全源文件只读／拒绝，不强制转换，不静默删除。

开发预览构建仍需按 [M5 交付与验收](docs/M5-DISTRIBUTION.md) 进行平台核验。完整 Office 参考应用、复杂文件、旧格式、人工语言审校及性能门槛仍未全部通过。M1 历史基础设施验收见 [M1 验收记录](docs/verification/m1/README.md)。

## 正式版格式目标

首个正式版必须覆盖 **DOC/DOCX、XLS/XLSX、PPT/PPTX** 的约定常用功能子集：新建、读取、编辑、原格式保存、另存、复开和新旧格式双向转换，并提供 PDF 阅读／导出。

旧版目标为 **Office 97–2003，XLS 使用 BIFF8**。共享 CFB 容器已进入 M1；DOC、XLS、PPT 的语义读写分别属于 **M2-L、M3-L、M4-L**，全部是首个正式版的强制门槛。能够识别容器不等于已经实现旧格式编辑。

## 实施与验收记录

[Writer / DOC(X)](docs/M2-WRITER.md) · [Sheets / XLS(X)](docs/M3-RECOVERY.md) · [Slides / PPT(X) / PDF](docs/M4-SLIDES-PDF.md) · [M5 分发与正式版门槛](docs/M5-DISTRIBUTION.md)

22 套界面语言候选目录已建立，支持热切换、独立地区格式、文档语言与 RTL 布局。候选翻译不等于人工审校完成；正式版仍须通过 [I18N-01–05](docs/specs/09-i18n.md) 与 QA-07。

## 本地启动

使用 Node.js **24.13.0 及以上的 24.x 版本**；CI 固定为 24.22.3。首次安装需要访问 npm 与 Electron 官方发行文件下载服务，应用运行时默认离线。

```sh
npm ci --ignore-scripts
npm run setup:fonts
npm run setup:desktop
npm run start
```

`setup:desktop` 显式安装锁定版本的原生 Electron；`start` 构建后打开桌面窗口。Linux 需要图形会话、可用的 Chromium sandbox 和中文字体；Ubuntu CI 安装 `fonts-noto-cjk`、`fonts-noto-color-emoji`，使用 Xvfb。系统字体不随应用分发，sandbox 配置见工作流。

## 开发预览打包

在对应原生平台安装锁定依赖、字体及 Electron 后：

```sh
npm run build
# 将实际源码提交 SHA 设为 OFFICE_EX_SOURCE_SHA（CI 自动提供）
npm run package:preview
npm run verify:preview -- out/m5-preview/Office-EX-0.1.0-dev.1-linux-x64.manifest.json # Linux x64 示例
npm run release:readiness
```

`npm run package:preview -- --installer` 在对应构建机上生成 Linux DEB、macOS DMG 或 Windows 安装器。分发仍处于预览验证阶段。普通分支推送仅生成 CI 产物；维护者主动推送与源码版本一致的开发预览标签（如 `v0.1.0-dev.1`），且三平台安装、启动及移除校验全部通过后，工作流才组装未公开的 GitHub 预发布草稿。不会自动公开或标记 Latest。打包文件附 SHA256SUMS、内容清单及第三方许可证。不要绕过系统安全策略或关闭 Chromium sandbox。正式发布检查使用 `npm run release:readiness -- --require-formal`。

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

### M2 Writer 第三批开发预览（开发分支）

新增窄范围 DOCX 新建/打开/编辑/保存/另存候选功能。未知内容、普通复杂 Word 文件以及未完成的 XLS/PPT/DOC 等仍只读结构检查。现有 .oex 内部草稿保留；不是改扩展名转换。完整验证和限制见 [M2 Writer](docs/M2-WRITER.md)，不得在原生 CI 未通过前宣称本批完整验收。

### M2 Writer 第四批开发预览

新增字体、字号、颜色、已有命名样式工具，以及受限的 DOCX 样式和主题读取与原部件保留。普通 Word 文档的完整样式、复杂部件和旧 DOC 仍在开发中；当前候选范围及验收见 [M2 Writer](docs/M2-WRITER.md)。

### M2 Writer 第五批开发预览

新增规则矩形表格、行列操作、单元格段落编辑、整行分页预览与固定列宽 DOCX 表格支持。复杂表格及图片仍在后续阶段；本批验收见 [M2 Writer](docs/M2-WRITER.md)。

### M2 Writer 图片与页面设置开发预览

增加受限 PNG/JPEG 内嵌图片段落、原始图片资源保存、页面纸型/方向/边距与段落间距。本阶段仍为限定功能开发版，不等于完整 Word 或旧 DOC 兼容。验收范围及剩余门槛见 [Writer 记录](docs/M2-WRITER.md)。

### Writer 连续文字与行内图片开发预览

新增连续相邻段落的原子范围编辑、纯文本剪贴板和图文混排（独立 DrawingML run，含单元格图片）。仍为受限 DOCX 与估算分页，非完整 M2 验收。详见 [Writer](docs/M2-WRITER.md)。

### Writer 分节和文档语言

限定下一页单栏分节及独立纸张参数；保留文字语言和直接书写方向标记。RTL/复杂文字精确布局、完整 Word 互操作与 DOC 仍在开发中。
