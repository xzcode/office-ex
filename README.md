# Office EX

面向个人及开发者的本地优先、跨平台开源办公套件。目标为 DOCX、XLSX、PPTX 的创建、编辑、保存与兼容交换，并提供 PDF 阅读和后续编辑。

## 当前阶段

M0 技术基线：首选 ONLYOFFICE Desktop Editors 10.0.0，尚未完成源码构建或实际编辑器验收。当前仓库没有可安装的 Office EX。

- [M0 技术决策与验收门槛](docs/M0.md)
- [许可证与第三方声明](docs/LICENSING.md)
- [上游源码锁定清单](upstream/lock.json)
- [兼容性用例与操作方法](docs/COMPATIBILITY.md)
- [M0 实际验证记录](docs/M0_VERIFICATION.md)
- [后续交接](docs/HANDOFF.md)

## 开发工具

需要 Python 3.10+。执行以下命令检查 Linux 构建环境（不安装任何依赖）：

```sh
python3 scripts/preflight.py
python3 scripts/preflight.py --network
```

在具备网络、磁盘条件的专用环境，获取并核对上游源码快照（不是完整的构建工作区，prepare 需要全新路径）：

```sh
python3 scripts/prepare_sources.py prepare .work/upstream
python3 scripts/prepare_sources.py verify .work/upstream
```

生成兼容性基线（虚拟环境内安装测试依赖）：

```sh
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r tests/requirements.txt
python tests/compat.py generate .work/fixtures
python tests/compat.py check .work/fixtures
```

## License

现阶段独立开发工具及文档采用根目录 MIT 许可证。未来引入的 ONLYOFFICE 代码及与其组成的衍生程序必须遵守其 AGPLv3 与适用附加条款，不能以 MIT 代替。参见 docs/LICENSING.md。
