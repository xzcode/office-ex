# 01 产品范围与架构

PRD-01 自研 MIT 桌面套件，Electron/TypeScript；自主编辑、排版、公式、场景、事务及格式语义映射。不 fork/嵌入/包装其他完整办公引擎。通用 ZIP/XML/CFB、字符编码、字体、图像及 PDF 基础库可在许可证/安全审计后使用；独立 Office 读取器仅作测试。不得擅自改技术路线。
PRD-02 Windows/macOS/Linux，macOS ARM64 首测，Windows x64/Linux x64 各自原生验证。离线优先，无账号、广告、遥测、服务器、默认联网更新。中文/英文为 M1 基线；正式版支持 22 个界面 locale，独立地区格式、文档语言、RTL、IME、Unicode、键盘交互。国际化按 I18N-01–05（[09-i18n.md](09-i18n.md)）。
PRD-03 P0 正式版必须覆盖 DOCX/DOC、XLSX/XLS、PPTX/PPT 的常用功能子集：新建、读取、编辑、原格式保存、另存、双向转换、复开。现代格式为 Transitional OOXML；旧版为 Office 97–2003（XLS BIFF8）。不承诺完整 Office 功能；旧格式不可被仅改扩展名或强制转换替代。M2-M4 可先交付内部现代格式里程碑，M5 正式版必须通过六种格式门槛。任意未知/危险功能先限制操作、只读或拒绝，绝不静默损坏。PDF 阅读/导出 P0；CSV/ODF/RTF、复杂图表、修订、母版等 P1；Office 95 及更早格式、宏执行、协作、插件和 AI P2/另议。

## 规划模块（尚未建立全部目录）
| 模块 | 职责 | 允许依赖 |
|---|---|---|
| packages/contracts | ID、单位、诊断、跨边界 DTO | 无 |
| packages/core | 会话、命令、撤销、资源、能力 | contracts |
| packages/writer | 文字模型、分页、命令 | core/contracts |
| packages/sheets | 稀疏网格、公式计算、命令 | core/contracts |
| packages/slides | 场景、形状、命令 | core/contracts |
| packages/format-common | 有界读取、格式识别、适配接口、写入计划 | contracts/core |
| packages/ooxml | OPC/ZIP/XML 与三种现代格式映射 | format-common、领域模型 |
| packages/cfb | 通用复合二进制容器与流 | contracts/format-common |
| packages/legacy-doc | MS-DOC 读写 | cfb、format-common、writer |
| packages/legacy-xls | MS-XLS 读写 | cfb、format-common、sheets |
| packages/legacy-ppt | MS-PPT 读写 | cfb、format-common、slides |
| packages/render | 布局/坐标/绘制基础设施 | core/contracts、模型 |
| packages/pdf | PDF 适配与打印输出 | render/contracts |
| packages/i18n | 类型化消息键、语言资源、locale 注册/回退、地区格式（M2-I18N 规划） | contracts |
| packages/ui | 界面、交互、主题 | 领域公开 API、render/core、i18n |
| apps/desktop | Electron shell、preload、平台文件与权限 | 公开 API |
| tests | 独立语义验证、夹具、回归 | 公开 API |

ARC-01 模型、公式、事务不依赖 DOM、Electron、文件系统、网络或格式层。UI 不直接变更模型或拼接文件格式。六个格式适配器共享三个领域模型；格式差异通过能力/来源元数据表达，绝不开发六个编辑器。禁止循环和跨包内部路径。src/index.ts 为公共出口。现有 OOXML 原型类型非稳定契约。
ARC-02 TypeScript strict、外部 unknown 校验、判别联合穷尽、无未经说明 any，禁止 eval/new Function 和执行文档字符串。命名统一，边界文档明确单位及错误；公开 Result/错误码包括 invalid、unsupported、unsafe、limit、conflict、cancelled、io、internal。IPC/worker 使用可序列化 DTO；大量解析/计算可取消；revision 不匹配时丢弃派生结果。
ARC-03 npm workspace、固定直接依赖版本、唯一锁文件、许可证/NOTICE/安全清单；依赖仅经明确理由引入。Electron/fflate/xmldom 仍是候选、未安装。CFB 底层通用库可评估，若不能满足安全与写入要求则自研；禁止拿完整文档转换器代替实现。Python 仅测试。
ARC-04 提交注明需求 ID、改动、验证、未完成项。不做无关重构/重复构建；未通过的功能不得标 verified。默认日志不含文档内容、公式、路径、密码。