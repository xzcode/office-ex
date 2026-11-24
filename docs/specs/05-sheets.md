# 05 Sheets

SHT-01 Workbook→有序 Sheet ID、styles、date system；Sheet→稀疏 cells、行列属性、merge、view。类型 blank|string|number|boolean|error|formula；空值不同于空字符串/0，日期存 Excel 序列与格式，保留 1900/1904（含 1900 闰年兼容）。数值 binary64，显示格式不作为源值。公式源、AST、缓存、状态分离。
SHT-02 稀疏索引+虚拟视口，严禁分配整个表；现代 1048576×16384，XLS 65536×256，转换必须拒绝越界。P0 单元格/范围编辑、批量 TSV 粘贴、工作表增删改名/重排、行列增删、基本样式及数字格式、合并拆分；合并非左上数据时明确阻止/确认，不静默清除。操作单事务，撤销恢复全部。
SHT-03 自研 lexer→parser→AST→resolver→dependency graph→evaluator，Excel 优先级/结合性/类型转换，禁止 eval/new Function 或外部公式引擎。支持 A1、绝对/混合引用、范围、带引号 sheet 名、括号、数值/字符串/布尔、算术/比较/连接；P0 SUM、AVERAGE、MIN、MAX、COUNT、COUNTA、IF、AND、OR、NOT、ROUND、ABS。实施前逐函数定义参数/类型/空值/错误/范围/惰性分支/精度，并用 Office 结果夹具验证。
SHT-04 稳定 sheet ID+坐标引用；复制调整相对分量，绝对不动；插删、改名重写受影响引用，删除目标 #REF!。范围索引避免展开整列；局部拓扑重算、循环检测、错误传播；结果 valid/stale/unsupported/error，未知公式保留源式及旧缓存、明确未计算，不写 0 或伪造值。限制步数、AST 深度与任务时间，可取消。复杂数组、共享公式、名称、结构化引用、外部链接、宏、图表等未理解的影响闭包阻止结构编辑。图表首版检测并安全保留，创建编辑 P1。
SHT-05 XLSX f/v 与 BIFF8 formula tokens 均由同一 AST 语义映射；公式序列化不代表求值。源日期系统、类型、样式和引用须往返；未知 token 不静默转换。目标 XLS 公式/格式能力预检。
验收：双格式数据类型/公式/跨表/绝对相对/错误/循环/IF 惰性/日期/行列改动，Office 对照计算，100k 非空单元格虚拟化（XLS 按实际合法分布），越界转换阻断、原件不变。