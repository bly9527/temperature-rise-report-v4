# 温升报告自动生成器 v4.0

从测温 CSV 生成 `.xlsx` 温升报告，包括总表、工况附表、可编辑的原生折线图和图片。

## 功能

- 测温数据写入 A 列扫描时间、B 列扫描编号和 C 列起的温度通道。图表使用 A 列时间及 C 列起的温度数据，不把扫描编号当成测温值。
- 附表图表位于 E9。PowerZ 图片在不超过 14 个通道时位于 Q1；通道更多时紧接最后一个温度列。
- 总表图片从第 30 行起，每行四张，排列在表格下方的图片格中。
- 支持分批更新工况，并保留报告中未参与本次生成的内容。
- 优先使用 Microsoft Excel；Excel 无法启动时使用 WPS 表格。

## 运行

适用于 Windows。安装 Python 3.10 或更新版本，并安装 Microsoft Excel 或 WPS 表格。

```powershell
python -m pip install -r requirements.txt
python thermal_gui.py
```

也可从项目的 Releases 下载 Windows 可执行文件。可执行文件已内置两份 `.xlsx` 模板，运行时仍需安装 Excel 或 WPS 表格。

详细操作见 [使用说明.md](使用说明.md)。

## 打包

```powershell
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean 温升报告自动生成器_v4.0.spec
```

## 测试范围

曾在 Microsoft Excel 16.0 与 WPS Office 12.1.0.28505 上使用真实 CSV 核对 14、15 通道报告，包括全部导入值、图表引用、图片位置及跨软件打开。其他电脑上的具体版本需要自行确认。

本仓库仅提供 v4.0 源码与去除旧外部链接缓存的空白模板，不包含用户的报告、CSV、配置或截图。
