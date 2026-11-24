# Office EX 开发规范 v1.0
2026-11-24。规范基线，非实现结果。MUST 为强制；SHOULD 为默认（例外需 ADR）；MAY 为可选。规范要求不得当作已完成能力。
- [产品及架构](01-architecture.md)
- [模型与事务](02-model.md)
- [现代及旧版格式](03-formats.md)
- [Writer](04-writer.md)
- [Sheets](05-sheets.md)
- [Slides/PDF](06-slides-pdf.md)
- [桌面、安全及保存](07-security.md)
- [质量与阶段门槛](08-quality.md)
- [功能追踪矩阵](features.json)

本目录为现行开发契约，与旧 M0 路线图冲突时以本目录为准。需求 ID 稳定，不复用已废弃 ID。变更记录原因、兼容性影响、迁移及测试，不能为了通过测试降低验收。实施流程：需求 ID → 模型/接口 → 夹具 → 实现 → 验证 → 更新矩阵与交接。未验证不宣传支持。