# 02 核心模型与事务

MOD-01 一份会话一个权威模型。DocumentKind writer|sheets|slides，各自 schemaVersion/documentId；稳定不透明节点 ID，重排不变。源格式（doc/docx 等）、原始资源/流/part、映射及能力报告由格式适配层管理；禁止让不同文件格式侵入编辑逻辑。内部 DTO 可序列化且可迁移，资源用 ID 和 SHA256 引用，未证明无未知引用不可清理。未知新 schema 不覆盖，迁移失败保留原快照。
MOD-02 统一几何整数 EMU：1 inch=914400 EMU，1 pt=12700，1 twip=635，96dpi CSS px=9525；旋转 1/60000 度，x 向右 y 向下。源格式精度映射及舍入显式，CSS 像素/缩放不存入文档。表格列宽按格式字符度量与字体转换。颜色显式 RGBA 或 theme 引用，缺失字体/样式诊断。文本逻辑位置 nodeId+UTF16 offset，移动/删除不切断 grapheme；RTL 逻辑/视觉位置分离。行列 1-based，现代上限 1048576×16384，旧 XLS 上限 65536×256；目标格式检查不能截断。
MOD-03 命令含 documentId、expectedRevision、type、payload；能力、范围、版本、权限检查先于执行；原子事务返回 revision/change set/inverse 或恢复快照。失败不得改模型/历史/dirty。批量粘贴/拖动/IME 提交为单事务。撤销重做 revision 单调递增，按持久内容 ID 判断 dirty；撤销后新编辑清空 redo。默认最多 200 项、64 MiB 历史预算，优先增量操作，不每次复制整个模型。
MOD-04 layout/recalc/decode 用只读 revision 快照，任务 ID、取消、过期丢弃。布局与公式缓存/视口/选区不作为源内容。会话维护源身份、capabilities、revision、lastSavedContentId、dirty、saveState。保存时出现新编辑仍 dirty；关闭脏文档提示保存/另存/放弃/取消。
MOD-05 逻辑接口（待实现）：
- detect(bytes) → format/version/diagnostics
- open(bytes, limits, cancellation) → model + provenance + capability report
- execute(command), undo(), redo() → transaction result
- snapshot() → immutable DTO; layout(snapshot,env) → layout; hitTest(layout,point) → position
- planSave(snapshot,original,targetFormat) → allowed/warnings/blockers/patch plan
- serialize(plan) → validated bytes; commitSave(token,expectedIdentity,bytes) → result
- diagnostic: code/severity/featureId/sourceLocation/messageKey/safeAction，稳定排序，不记录内容。
验收：失败无副作用，跨标签历史独立，撤销恢复已保存状态，并发保存不误清 dirty，几何及 Unicode 往返，版本迁移保留原件。