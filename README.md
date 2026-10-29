# 温升报告自动生成器 v4.0

从测温 CSV 生成 `.xlsx` 温升报告，包括总表、工况附表、可编辑的原生折线图和图片。

## 功能

- 测温数据写入 A 列扫描时间、B 列扫描编号和 C 列起的温度通道。图表使用 A 列时间及 C 列起的温度数据，不把扫描编号当成测温值。
- 附表图表位于 E9。PowerZ 图片在不超过 14 个通道时位于 Q1；通道更多时紧接最后一个温度列。
- 总表图片从第 30 行起，每行四张，排列在表格下方的图片格中。
- 支持分批更新工况，并保留报告中未参与本次生成的内容。
- 未勾选且尚未生成的工况不占用总表容量；总表最多保存 12 个工况、23 个测温通道。
- 删除工况前要求确认，确认后同步删除对应附表和总表数据，剩余工况自动左移；新增结果接在末尾，附表顺序与总表一致。
- 删除失败时保留原报告和配置；错误弹窗显示原因，不显示程序堆栈。
- 优先使用 Microsoft Excel；Excel 无法启动时使用 WPS 表格。

## 运行

适用于 Windows。安装 Python 3.10 或更新版本，并安装 Microsoft Excel 或 WPS 表格。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe thermal_gui.py
```

安装好独立环境后，也可双击 `启动温升报告生成器.cmd`。程序已处理中文安装路径下的 Qt 插件定位问题。

使用独立 EXE 时无需另外安装 Python。可执行文件已内置两份 `.xlsx` 模板，生成报告时仍需安装 Excel 或 WPS 表格。

详细操作见 [使用说明.md](使用说明.md)。

## 打包

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean 温升报告自动生成器_v4.0.spec
```

独立 EXE 输出到 `dist` 目录。工程配置保存在 EXE 旁的 `configs` 目录，更新程序时请保留该目录。

## 测试范围

曾在 Microsoft Excel 16.0 与 WPS Office 12.1.0.28505 上使用真实 CSV 核对 14、15 通道报告，包括全部导入值、图表引用、图片位置及跨软件打开。其他电脑上的具体版本需要自行确认。

本轮容量与工况删除功能已在 Microsoft Excel 16.0（构建 20527）验证：取消删除、未生成工况删除、中间工况删除与排版、报告被占用时保护原文件、删除后新增、连续删除、保留其他工况曲线和公共图片，以及删除最后一个工况。WPS 的这些新增操作尚未单独复测。

2026-10-27 使用 PyInstaller 6.22.3 打包的独立 EXE 已通过隔离目录验收：内置模板新建工程、生成带图表和图片的报告、取消删除、确认删除及总表左移排版。打包配置也修正了中文路径下 Qt 插件目录识别失败的问题。

2026-10-29 修复 Windows 中文输入法下，单击通道名称/限值后直接输入引发 Qt 栈溢出闪退的问题。通道和工况名称选中后立即进入编辑状态；新增中文输入、限值修改、配置保存重载和复选框回归测试。

自动回归测试：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

本仓库仅提供 v4.0 源码与去除旧外部链接缓存的空白模板，不包含用户的报告、CSV、配置或截图。
