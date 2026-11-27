# Office EX

[English](README.md) | **简体中文**

自主研发、跨平台、离线优先的开源办公套件，涵盖文字处理、电子表格和演示文稿。

Office EX 使用 TypeScript 开发，作为桌面应用运行。文档模型、编辑命令、撤销重做、排版、公式和 Office 格式适配器由项目自主实现。Electron 与通用开源库提供运行时、压缩、XML 和 PDF 等基础设施，不将 Microsoft Office、LibreOffice 或 ONLYOFFICE 作为内嵌编辑引擎。

## 下载

当前已发布的最新开发预览版为 **[v0.1.0-dev.4](https://github.com/xzcode/office-ex/releases/tag/v0.1.0-dev.4)**。

| 平台 | 安装包 |
| --- | --- |
| Windows x64 | [下载 EXE](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.4/Office-EX-0.1.0-dev.4-win32-x64-setup.exe) |
| macOS Apple Silicon | [下载 DMG](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.4/Office-EX-0.1.0-dev.4-darwin-arm64.dmg) |
| Linux x64 | [下载 DEB](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.4/Office-EX-0.1.0-dev.4-linux-x64.deb) |

查看[全部版本](https://github.com/xzcode/office-ex/releases)、[SHA-256 校验文件](https://github.com/xzcode/office-ex/releases/download/v0.1.0-dev.4/SHA256SUMS)及[发布验收记录](https://github.com/xzcode/office-ex/actions/runs/45955956363)。

目前为开发预览版，而非正式生产版本。Windows 安装包尚未签名，macOS 应用采用临时签名且未公证。请保持系统安全保护开启，并备份重要文档。

## 功能

已发布的 dev.4 预览版实现并验证了以下功能的约定子集。

| 模块 | 功能 |
| --- | --- |
| Writer 文字 | DOC / DOCX 编辑、文字样式、表格与合并单元格、分节排版、页眉页脚、自动页码、首页及奇偶页、多段落页眉页脚、段落分页 |
| Sheets 表格 | XLS / XLSX 编辑、虚拟化单元格网格、格式设置、48 个公式函数、条件格式、基础柱状图、折线图和饼图、富文本及普通命名样式 |
| Slides 演示 | PPT / PPTX 编辑、文本及图形、图片、组合、原生合并表格及单元格格式、对齐和等距排列、演示模式 |
| PDF | PDF 阅读与受支持文档的导出 |
| 桌面 | 默认独立文档窗口（可选标签页）、八种文件关联、标题栏拖动及面板调宽、撤销重做、原格式保存与另存、历史备份、崩溃恢复、主题、22 个界面语言区域 |

在受支持范围内保留原始文件的原生部件。对于超出编辑能力的复杂文档，采用保护策略，避免静默丢弃或扁平化内容。嵌入式 Office 对象、部分生成器的简化编号、未定义样式引用、旧版二进制格式、高级排版、公式、图表、表格、打印、无障碍与参考 Office 排版精度尚未全面兼容。内部 .oex 草稿不是 Microsoft Office 格式。

## 开发进度

- **M0–M6：** 已完成开发预览版交付。dev.4 候选版本通过 316 项核心测试，以及 Windows、macOS、Linux 的原生打包和安装验收。
- **M7.1–M7.5：** 已随 dev.4 交付。继续推进真实文档兼容性、稳定性和体验；dev.4 之后新增能力属于后续开发版本。

默认 `master` 分支目前保留早期实现基线。最新源码和文档请查看[活跃开发分支](https://github.com/xzcode/office-ex/tree/m5-distribution-20261125)。需要与安装包一致的源码时，请使用 [dev.4 标签](https://github.com/xzcode/office-ex/tree/v0.1.0-dev.4)。本文将已发布能力与正在开发的功能区分说明。

## 从源码构建

使用 Node.js 24.x（至少 24.13.0）。首次安装需要下载 npm 依赖和 Electron，正常文档编辑在本地运行。

```sh
git clone -b m5-distribution-20261125 https://github.com/xzcode/office-ex.git
cd office-ex
npm ci --ignore-scripts
npm run setup:desktop
npm run start
```

Linux 需要图形会话、正常工作的 Chromium 沙箱及合适的字体。不要关闭沙箱。

## 开发与验证

Python 用于独立兼容性夹具和文件回读测试，不参与生产编辑引擎运行。

```sh
python -m pip install -r tests/requirements.txt
npm run check
npm run test:interop
npm run audit:licenses
npm run build
npm run smoke:desktop
```

查看活跃开发分支的[开发交接](https://github.com/xzcode/office-ex/blob/m5-distribution-20261125/docs/HANDOFF.md)、[M6 兼容性说明](https://github.com/xzcode/office-ex/blob/m5-distribution-20261125/docs/M6-COMPATIBILITY.md)和 [M7 稳定性工作](https://github.com/xzcode/office-ex/blob/m5-distribution-20261125/docs/M7-STABILITY.md)。

## 开源许可证

原创代码使用 [MIT 许可证](LICENSE)。第三方组件遵循各自的许可证，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。Office EX 为独立项目，与 Microsoft 无隶属关系。
