# Office EX

自主研发、跨平台、离线优先的开源办公套件。Office EX 自己实现编辑器、排版、公式计算、幻灯片与用户界面，通用开源组件只承担桌面运行时、ZIP/XML、字体和渲染等基础设施。**不基于其他 Office 套件进行二次开发。**

## 状态

M0-R2：自主路线已确定。原创 TypeScript OOXML 写入原型可以生成基础 DOCX/XLSX XML 部件，经测试工具封装 ZIP 后由独立解析库读取。尚无桌面可安装包，不宣称完整 Office 兼容性。

- [技术方案、阶段和验收门槛](docs/M0.md)
- [许可证与依赖规则](docs/LICENSING.md)
- [验证记录](docs/M0_VERIFICATION.md)
- [兼容性基线](docs/COMPATIBILITY.md)
- [交接](docs/HANDOFF.md)

## M0 开发验证

Node.js 22.13+、TypeScript 5.8.3（开发环境）、Python 3.10+（仅测试）。执行：

```sh
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r tests/requirements.txt
tsc -p tsconfig.json
npm test
```

另可运行 tests/compat.py generate .work/fixtures 和 tests/compat.py check .work/fixtures；这些是扩展兼容性夹具，暂非自研编辑器通过记录。

MIT license. 保留原仓库许可证；第三方依赖进入生产构建前须进行许可证与安全审计。
