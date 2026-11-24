# M2 Writer — 第一批：真实段落事务与 DOCX 窄语义适配
2026-11-24。状态：实现候选，待本批 CI 验证。基线 e68578a（11 个候选 UI locale 已通过原生 CI 45380700385）。不修改已验证 M1 或国际化证据。

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
