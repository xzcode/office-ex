# M2 Writer — 第一批：真实段落事务与 DOCX 窄语义适配
2026-11-24。状态：首批核心实现已通过限定范围验收，M2 整阶段尚未完成。基线 e68578a（11 个候选 UI locale 已通过原生 CI 45380700385）。不修改已验证 M1 或国际化证据。

## 实现
- packages/writer/src/paragraphs.ts：以段落 ID 和段落全局 UTF-16 offset 定位文本，使用 Intl.Segmenter 验证 grapheme 边界（跨 run）；保留未修改片段/样式，事务化替换、格式化、拆分、合并、对齐、显式分页符。插入继承前一 run（起点继承第一非空 run）；新节点使用调用者提供的唯一 operation ID 前缀。禁止在此 API 扁平化表格或非文本 inline。分合段使用前段样式，历史由 DocumentSession 管理。
- packages/ooxml/src/writer-docx.ts：基于已有受限 OPC/XML/ZIP 容器；独立 `writeWriterDocx(model)`、`openWriterDocx(bytes,id,title)`、`saveWriterDocx(source,model)`。源码句柄使用 WeakMap 保存原包与身份，不允许伪造句柄或修改暴露的原包。无更改时原字节返回；支持范围内更改仅替换 main part，重开并比较规范化语义 XML。
- 首个支持 profile：Transitional DOCX，仅 [Content_Types].xml、根关系、主部件，主部件路径由 OPC 关系解析，不硬编码；单 section 显式页宽/高和四边距；多个段落、空段、文本 run、Unicode、空白、tab/line break、独立 page break、直接粗斜下划线/字体/字号/不透明 RGB、对齐/缩进/段间距/精确行距。字号以半点、几何以 twip 舍入，不能保证源布局完全一致。文本 controls 用明确元素表示；导出将默认字符样式展开为直接样式，导入重建 ID；测试对比语义而非 JSON 字节相同。
- 安全策略：额外 part/relationship、宏/签名/外链、Strict、未识别属性/节点、命名样式、主题色、图片、表格、分节、复杂字体或语言/方向属性等暂阻止语义适配，不静默丢弃。常见 Microsoft Word 产生的包含多个样式等部件的 DOCX 通常仍只读结构检查，后续逐项支持。所有原始文件保留；这批 API 不直接进行文件 IO。
- 不增加第三方生产依赖，不更改现有 M0 fixture 或格式检测报告，避免将 C1 容器识别冒充编辑能力。Writer UI 暂未接入新命令与 adapter；桌面打开 DOCX 仍为原来的结构检查。真实分页、复杂脚本 shaping、文档语言/方向元数据、图片/表格、打印、DOC 和完整 Office 互操作未完成。无需重复桌面原生 smoke（本批未修改 UI/IPC/保存代码）；仍执行基础回归、独立 DOCX/CFB 互读与许可证检查。

## 验证入口
- tests/writer-paragraphs.test.ts：跨 run、组合字符、格式拆分、原子分合段、空段、分页符、unsupported/冲突与 undo/redo。
- tests/writer-docx.test.ts：新建/重开、编辑/原样保存、非默认 part 路径与前缀、未知内容阻断、原包不变。
- tests/writer-docx-fixture.mjs + tests/writer_docx_interop.py：生成实际原始与编辑后 DOCX，独立 python-docx 检查多语言正文、直接字体格式、对齐、纸宽及分页段落；不是微软 Office GUI 证据。
- npm run check; npm run test:interop; npm run audit:licenses。工作流 .github/workflows/m2-writer.yml 在 m2-writer-20261124 独立运行。

## 下一批
Writer 布局与输入：可测量行盒、分页、grapheme selection/IME 与实际富文本编辑；补齐文档语言/方向及字体回退、资源字节封装。再逐项实现普通 DOCX 的 styles/relationship/未知部件影响闭包，接入安全原格式保存，推进 M2-L DOC。首个正式版仍需六格式、22 locale、PDF 的全部 P0 门槛。

## 本批验收

受测源码 7692c0c7da8e561f3d15da71edf1ab8f2b314a06；[CI 45424554932](https://github.com/xzcode/office-ex/actions/runs/45424554932) 全部成功。Ubuntu 24.04 / Node 24.22.3 / Python 3.12.14：13 workspace、37 源文件；严格类型和架构检查通过；121 项自动化测试全部通过，0 失败、0 跳过（较上一阶段新增 13 项 Writer 用例）。独立 python-docx 读取新建和编辑后 DOCX，核对多语言正文、字体格式、对齐、分页段落与纸张宽度；旧 OOXML/CFB 互读、66 项锁定 npm 包及许可清单检查均通过。

首轮 CI 45424318323 因 XML localName 允许 null 而编译失败；已显式检查空值，补充控制字符专用 run 与 accessor 安全测试，以上最终 CI 为实际证据。本轮不改桌面 UI/IPC/文件服务，因此未重新跑三平台原生 GUI。正式 M2、M2-L、复杂格式与 Microsoft Office 实测仍未通过完整验收。

## 第二批：布局与输入工作区（实现候选）
- render/writer-layout.ts：基于 FontMetrics 的有界行盒和页盒，支持单栏横排 LTR、Latin/CJK 基础断行、直接/命名样式、页边距、对齐、缩进、间距、显式分页符、软换行和固定四空格宽 tab；段落跨页。派生结果含 revision/contentId/environmentId，不写回源模型。提供 UTF-16 grapheme 光标、点击命中、同段选区框。
- Writer 每段原生 textarea，跨 run 最小 grapheme diff、单 run 稳定 ID、选区粗斜下划线和对齐、Enter 选区替换并分段、段首退格/段末 Delete 合并、显式分页符。单页 Canvas 预览与编辑共享 snapshot，点选定位到段落输入，支持页数/翻页。composition 时阻止可能丢弃输入的切换与关闭，真实物理输入法仍待测。
- 11 套 locale 同步增加 18 个界面键（155 个/locale），无新外部依赖。内部 OEX 编辑保存保持原有路径；未接入桌面 DOCX 编辑保存。编辑区与排版预览分离，未完成 WYSIWYG，不宣称打印精度。Canvas cluster 宽度及断行估算；复杂 bidi/印度文字/泰文等 shaping、跨样式 grapheme、图片/表格和超限内容拒绝预览，不覆盖编辑内容。负缩进、justify、精确过小行距等有限处理；未达到 QA-04。
- 新增确定性布局测试与原生 Writer 操作 smoke；全套 CI 在当前分支运行。实际结果另记，不借用前批证据。

### 第二批验收记录

受测源码 84db8010c281f4eefdb80b7250b57148b3876f44，[CI 45426118923](https://github.com/xzcode/office-ex/actions/runs/45426118923)。基础 verify-writer 已通过：13 workspaces、39 源文件，129/129 自动化测试成功，0 失败、0 跳过；新增 8 项布局/输入用例覆盖自动分页、空段、显式换页、选区与命中、样式、CJK 断行、取消与资源预算、grapheme diff 和原子分段。独立 python-docx 语义互读和依赖许可审查通过。

记录时任务状态：verify-writer: success; native-macos: queued; native-other (windows-2025): in_progress; native-other (ubuntu-24.04): in_progress。尚未完成的原生任务不得记为通过。当前只在独立 Writer 分支，未合并 master。排版属于估算 LTR profile，完整 Word 排版、真实 IME、复杂脚本与 DOCX 桌面安全读写仍待后续验收。

## 第三批：DOCX 桌面打开与安全保存（实现候选）

- 新增明确的「新建 DOCX」入口；OEX 草稿和 DOCX 原格式会话各自保留独立格式，禁止仅更改扩展名。DOCX 会话标题显示文件名，另存改变文件关联。恢复快照仍为 OEX，恢复为未保存副本，不自动覆盖原 DOCX。
- `inspectForEditor` 在结构检查之后尝试窄 profile 语义读取；只有实际 Transitional DOCX、正确扩展名、明确支持内容及 32 MiB 以内源文件进入编辑。普通 Word 文件、未知部件/属性、复杂样式、图片表格、签名宏及其他格式维持只读和原样另存。能力报告仅标记 WRT-04 窄范围 C2 候选，不提升六格式整体兼容等级。
- `office:save-docx` 传递受限 OEX 模型快照而不是调用者提供的 DOCX 输出。主进程校验窗口、会话、内容身份、文件 token 和原始摘要；独立临时隐藏 sandboxed Chromium 窗口采用单独非持久 session、专用 preload、CSP、固定本地资源白名单及 Web Worker。主进程从授权缓存取得原文档，worker 负责导入/序列化/重开，主进程只接收验证器来源的有界字节；正常 renderer 无验证器 IPC 权限。复用原子提交、备份、身份冲突检测及恢复清理。没有引入外部生产依赖。
- 11 套候选语言新增 3 个键（158/locale）并更新首页、M2 标记及能力说明。DOCX 原格式支持仍是狭窄的文本/直接格式/基本页面参数；未实现完整 Word 兼容、格式化工具栏扩展、图片、表格、文档语言方向元数据、DOC 及完整 WYSIWYG。
- 新增 4 项核心/IPC 回归与 native smoke 的直接越权请求阻断、新建、另存、复开、编辑、备份及外部文件冲突；正式结果以本批 CI 为准。仍须独立读者及参考 Microsoft Office 验收；未假定所有原生任务成功。

### 第三批验收记录

受测源码 c79feaf16191e01b890e16012f7daca98a0a30ef；[CI 45428117703](https://github.com/xzcode/office-ex/actions/runs/45428117703)。基础 verify-writer 通过：13 workspaces、42 源文件，133/133 自动化测试成功，0 失败、0 跳过；新增 4 项覆盖语义编辑入口、正确扩展名、模型快照生成、无修改字节原样保留、未知源阻断及 IPC 身份和边界校验。独立 python-docx 读取新建与编辑后文件，依赖许可检查通过。

记录时任务状态：verify-writer: success; native-other (ubuntu-24.04): in_progress; native-macos: in_progress; native-other (windows-2025): in_progress。未结束原生任务不得标记为通过。上一批 45426118923 的 Windows/Linux 原生任务已通过，macOS 因新工作流取消。完整桌面 DOCX 读写须以本批实际原生结果验收；未合并 master。普通 Word 文件的 styles/资源、图片表格、文档语言方向与完整 WYSIWYG 仍未实现，旧 DOC 的 M2-L 门槛不变。恢复 DOCX 快照为新的 OEX 草稿副本，避免覆盖原文件。

## 第四批：格式工具栏与命名样式/主题 profile（实现候选）

- 第三批受测 c79feaf 的 CI 45428117703 全部任务已结束并成功：133/133 基础回归，macOS、Windows、Linux 原生桌面验收通过。第四批独立验收，不复用前批证明。
- Writer 在现有直接格式命令上新增字体名称输入/建议（不是已安装字体枚举）、半磅字号、RGB 颜色，及已有段落/字符样式选择。选区内只改指定属性；光标无选区时不执行直接字符格式。UI 显示选区起始 run 的有效属性；混合选区精细状态和新建样式仍待完善。11 locale 各 164 个消息键。
- 扩展 `NamedStyle` 的可选 kind、bold/italic toggle，保留旧 OEX schema 可读。解析 docDefaults、支持的段落/字符命名样式、basedOn 链、明确的样式引用及样式内主题色/一致的主题字体。粗/斜体按命名样式 toggle 与直接格式区分。未知未使用样式保留但不暴露；已引用、默认或继承链上不可理解的样式阻断编辑。主题颜色限 RGB/系统色 lastClr，禁止 tint/shade、复杂字体映射与主题修改；直接正文主题属性仍不支持。链接/next 等样式元数据原样保留，编辑器暂不自动切换 next style。允许已知根元数据部件，保持字节，不更新其中统计字段；禁止未知关系、settings、numbering、复杂对象或外部依赖。
- 只允许 source-owned 样式/主题 profile；样式定义和文档默认值不可修改，原格式保存只替换主部件，完整保留非主部件，原样另存字节不变，重开核验语义。样式/主题源不写入普通新建 DOCX；从 OEX 恢复的有命名样式文档暂不能转换为新 DOCX。未实现常见完整 Word 样式/主题、脚本映射、复杂格式、完整 WYSIWYG 和 DOC，不能声称所有 Word 文档可编辑。
- 新增受控 DOCX 夹具与独立 python-docx 读取验收，包括命名样式/主题、直接覆盖、原始非正文部件、header/footer 偏移。新增基础及 native smoke 用例；实际结果见本批 CI，不把受控夹具当作 Microsoft Office 实测。无新生产依赖。

### 第四批验收记录

受测源码 75cad7b0ac4a59043b666334b41e70af914dc781；[CI 45432375135](https://github.com/xzcode/office-ex/actions/runs/45432375135)。基础 verify-writer 已通过：13 workspaces、43 源文件；137/137 自动化测试通过、0 失败、0 跳过（新增 4 项样式测试）。命名样式继承及 toggle、主题字体/颜色、直接样式覆盖、样式定义不可改、未知依赖和循环阻断均通过。独立 python-docx 实际读取受控的原始与编辑后样式 DOCX，校验段落/字符样式 ID、继承引用、字体/字号/RGB、页眉偏移及所有非正文部件逐字节一致。原有 DOCX/XLSX/CFB 互读及依赖审查通过。

记录时任务状态：verify-writer: success; native-other (windows-2025): in_progress; native-macos: in_progress; native-other (ubuntu-24.04): in_progress。原生工具栏、样式选择与完整桌面保存以本次实际任务最终结果为准，未完成不得标为通过。当前代码未合并 master。此为源样式只读保留 profile，不是任意普通 Word 文件兼容：settings/numbering、复杂样式、直接主题属性及非一致字体映射暂受限；混合选区精确状态、自动 next style、新建样式、表格、图片、完整 WYSIWYG、参考 Office 和旧 DOC 仍待开发/验收。

## 第五批：矩形表格事务、分页与 DOCX（实现候选）

- 第四批 CI 45432375135 对源码 75cad7b 三平台 native 及 137/137 基础回归全部通过；本批单独验证。
- 模型扩充可选 cellPadding、uniform border 和 defaultParagraphStyleId；保留 schemaVersion 1 旧草稿的读取能力。Writer 支持 top-level 规则矩形表格，默认列宽以 twip 分配至当前纸张正文宽度；行、列插入删除、列宽调整、边框开关和删除表格均使用 WRT-03 事务、稳定 ID、撤销/重做。删除末行/末列必须显式删除表格；单节仅剩一个表格时阻止删除，避免产生空 section。插列从相邻列分配宽度，删列将宽度归还相邻列，调整列宽保持总宽度（单列允许不超过正文宽度）。同一单元格可编辑多个段落、直接格式和已有命名样式，可分段/合段，禁止跨 cell 合并、cell 内硬分页、嵌套表格与不支持的 inline 对象。新增 cell 使用导入模型的默认段落样式。
- 排版复用已有 glyph/line 计算，以各 cell 内容宽度和 padding 布局，行高取最高 cell；整行跨页移动，生成 tableCells 与段落的逻辑 caret/selection 共享快照。Canvas 绘制 uniform border；点击在 cell 矩形内优先匹配该 cell 的段落。当前行不能拆成两页；超高行或超宽文字明确拒绝预览，不裁剪/删改源模型。复杂脚本仍沿用已有 layout 限制。
- DOCX 支持严格固定宽度矩形表格：显式 tblW、fixed tblLayout、tblGrid、四边 tblCellMar、可选一致六边 single RGB border，行/列/单元格均须规则矩形及 tcW 与 grid 一致，单元格含普通段落。单 section 现有页宽/边距限制不变。默认 table style 不可解释时阻止有损编辑；源样式/主题与其它允许非正文部件字节不变。合并/跨列/纵向合并、autofit、嵌套、复杂边框、行高、单元格额外属性及未知内容不在本 profile，保持只读。不宣称所有 Word 表格可编辑。
- UI 新增表格工具栏、行列数量输入、增删行列、列宽（pt）、边框开关、删除、可直接编辑的矩形 cell grid。现有格式工具对选中 cell 的段落有效；分页预览仍与编辑区分离。11 套 locale 新增 12 键，176 键/套；无新生产依赖。
- 新增 8 项定向核心回归及独立 python-docx 表格互读，原生 smoke 覆盖创建、cell 内容、行列编辑、撤销、DOCX 保存与复开。CI 实际结果另记。下一批为图片资源封装与 DOCX inline 图片，之后补普通 Word 表格、复杂排版、旧 DOC 及全部 P0 门槛。

### 第五批验收记录

首批源码 286cf545befab1b2c9a28644c0ab3e9a6a559227，[CI 45434512688](https://github.com/xzcode/office-ex/actions/runs/45434512688) 的 verify-writer 成功：13 workspace、44 源文件，145/145 自动化回归通过、0 失败、0 跳过（新增 8 项表格用例）。独立 python-docx 验证真实 DOCX 表格行列、单元格多段落、直接文字格式、列宽、边框、编辑新增行与非正文部件字节一致。由于复核 Transitional CT_TblCellMar 顺序（top/left/bottom/right）后提交修正，新工作流替换旧 native 任务；不能把旧任务算作最终通过。

最终受测源码 f72128a9f5c910811429e364399f8ee7487d7452，[CI 45434702537](https://github.com/xzcode/office-ex/actions/runs/45434702537)；check、test:interop、audit:licenses 三个实质验证步骤均已通过。 修正增加 XML 边距顺序独立检查以及嵌套表格拒绝样例。记录时任务状态：verify-writer: in_progress。仅以最终工作流实际结果认定 native 验收，未合并 master。仍限显式固定矩形表格，整行分页，复杂表格/拆行、图片及 DOC 未完成，未在 Microsoft Word 实测。
