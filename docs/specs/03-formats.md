# 03 文件格式与转换规范

## 统一原则
FMT-01 不信任扩展名：识别 ZIP/OPC、CFB 签名及内部部件/流版本；区分加密 OOXML、旧版 Office、其他 CFB、伪扩展名与损坏文件。识别错误需可读诊断。六种目标均实现读写适配；每个适配器输出领域模型+来源映射+能力报告，不能只取纯文本。
FMT-02 打开时保留原始包/流、未知记录、关系与哈希；未编辑时原样复制字节或无操作。编辑保存优先最小修改；对每个命令确定影响闭包（索引、偏移、引用、缓存、样式、资源）。未知内容安全时原样保留；不确定关联时限制编辑/保存或只读。禁止整包重建后默默丢失未知功能。
FMT-03 保存前构造计划、检查能力、完整性与目标限制；生成临时文件，独立重读验证语义后原子提交。跨格式转换明确展示不可表达功能，默认阻止；若未来允许简化副本，需明确列出降级项、用户确认及不同目标路径。签名、加密、宏、OLE、外链不执行，不能默默去除保护或声称签名仍有效。原件始终保留。目标格式限制如 XLS 行列/公式长度、样式索引及版本能力须预检。
FMT-04 逐项记录格式、版本、read/write/preserve/blocked、源位置、诊断与夹具。双向转换 doc↔docx、xls↔xlsx、ppt↔pptx 通过统一模型，不调用外部办公套件。新建导出和原格式往返分开验收。

## OOXML（ECMA-376）
FMT-05 首版 Transitional；Strict 需识别，未验证则明确限制。OPC 以 Content Types、根关系、相对 part URI、Relationship Id/Type/Target/TargetMode 建图；按 namespace URI+localName 解析，不依赖前缀/硬编码主部件路径。内部目标必须存在，ID/规范化路径不能冲突；外部目标不连接。mc:AlternateContent、扩展、未知属性、样式等需保留或阻止影响命令。未修改 part 解压字节哈希一致；被修改 part 保留未受影响未知节点。写入必要关系/类型/部件和正确 XML escaping。
- DOCX: document/body、styles、numbering、sectPr、media/relationships。
- XLSX: workbook/sheet relationships、worksheet、sharedStrings/inlineStr、styles、f/v、date1904、merges。
- PPTX: presentation/slide relationships、shape tree、slide size、layout/master/theme、media；不能省略必要母版与关系。
验证：不同前缀/合法非默认 part URI、关系缺失/重复、未知内容保留和危险编辑阻止。

## 旧版 Office 97–2003（必须通过正式版门槛）
LEG-01 CFB/OLE Structured Storage 通用读写（MS-CFB）：签名 D0 CF 11 E0 A1 B1 1A E1，header/sector size、FAT/DIFAT/MiniFAT、directory storage/stream、mini stream、sector 链。使用有界 little-endian reader，校验版本、大小、偏移、计数、安全整数、重复/循环/交叉链、截断和层级；不执行 OLE 对象、不提取到任意目录。保留未知 stream 字节和 storage 结构（逻辑内容，而非原扇区物理位置）。CFB writer 正确分配扇区、目录、FAT、MiniFAT 等，并经独立解析器读取。识别加密 OOXML 的 EncryptionInfo/EncryptedPackage，不误判为 doc。
LEG-02 DOC（MS-DOC）：WordDocument、0Table/1Table、Data 等；按 FIB 版本/标记定位结构；处理 CLX/piece table、CP/FC 映射、Unicode/压缩字符及 codepage、CHPX/PAPX、节/样式、表格和内嵌图片的 P0 子集；写入时正确重建受影响指针、长度、文本片段和格式索引。未识别格式运行/域/对象阻止危险操作。新建输出 Word 97–2003 可读的受测 profile；旧版本导入兼容须逐个验证。
LEG-03 XLS（MS-XLS，BIFF8）：Workbook stream、BOF/EOF、BoundSheet、SST/CONTINUE、数字/RK/字符串/布尔/错误、XF/FORMAT/FONT、行列/合并、Formula token/cache、1900/1904 日期。记录边界/长度有界；工作表偏移、共享字符串、格式索引和相对/绝对公式引用按规范处理。BIFF8 公式 token 与自研 AST 映射；未支持 token 保留/阻止修改，不能替换为数值。旧 XLS 上限 65536 行、256 列（IV），超限转换拒绝，不截断。不承诺 BIFF5/7。
LEG-04 PPT（MS-PPT）：PowerPoint Document、Current User、Pictures 等；记录头、container/atom、persist object 与编辑链、slide/master/layout、文本/图形/OfficeArt/图片；处理 ID、长度、偏移与资源映射。写入可验证的有效 persist 映射和必要记录；不能用图片代替可编辑文本/形状。未知记录按影响范围保留或阻止修改；复杂媒体/动画等先只读。
LEG-05 实施顺序：共享 CFB 读→写及安全回归；三个旧格式分别检测/解析→领域映射→最小新建输出→原格式意图编辑→双向转换→参考 Office 重开。先做 XLS 再 DOC 再 PPT；各自可并行，但不能以仅成功读取宣称可写。P0 是普通未加密、无宏文档的约定子集，Office 95 及更早版本需识别/拒绝，不在首版保证内。
LEG-06 夹具：Office 97/2000/2002/2003 各来源能取得的可授权合成样本；来源和版本必须可证。参考较新 Office 可打开旧格式验证，不能代替未实际测试的旧版应用兼容性声明。至少每种格式验证新建、源文件原样另存、指定编辑后原格式重开、双向转换、未知记录保护、损坏扇区/流/记录拒绝。不能获得某历史版本环境时明确标注未测并保留门槛，不造证据。

## 标准来源
- ECMA-376: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- MS-OE376: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376
- MS-CFB: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
- MS-DOC: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/
- MS-XLS: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls
- MS-PPT: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt
- Office formats: https://learn.microsoft.com/en-us/office/compatibility/office-file-format-reference

以公开规范自主实现；开发前锁定使用的文档修订与条款，检查 Microsoft Open Specifications 的知识产权声明及适用承诺，不把规范公开等同于所有专利/素材均无条件授权。