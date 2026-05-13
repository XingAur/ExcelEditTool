# ExcelEditTool

ExcelEditTool 是一款 Windows 桌面端 Excel 汇总工具，面向日常表格整理、物料清单汇总、财务明细合并等场景。它可以在本机读取 `.xlsx` 文件，预览原始数据，按指定字段分组，并对数量、金额等列自动求和。

## 产品特点

- 本地处理 Excel 文件，适合处理日常工作表和多 sheet 表格。
- 支持选择表头行，适配不同格式的原始表。
- 支持按多个字段分组，例如规格型号、单位、单价等。
- 支持对多个数字列求和，例如数量、金额等。
- 原始数据支持表头下拉多选过滤，汇总只基于过滤后的数据。
- 原始数据和汇总结果都支持双击编辑单元格。
- 汇总结果支持拖拽调整列顺序。
- 支持将汇总结果带格式复制到 Excel。
- 支持导出新的 `.xlsx` 文件，默认命名为 `原Excel文件名_汇总.xlsx`。

## 使用流程

1. 打开 ExcelEditTool。
2. 选择本地 `.xlsx` 文件。
3. 根据实际表格设置表头行。
4. 选择分组列和求和列。
5. 如有需要，在原始数据预览中使用表头筛选或直接编辑单元格。
6. 点击“生成汇总”查看结果。
7. 复制结果到 Excel，或导出为新的 `.xlsx` 文件。

## 下载与安装

Windows 安装包请到 GitHub Releases 下载：

- [Latest Release](https://github.com/XingAur/ExcelEditTool/releases/latest)
- 安装包文件名：`ExcelEditTool_Tauri_Setup.exe`

安装信息：

- 默认安装目录：`D:\Programs\ExcelEditTool`
- 安装完成后可从开始菜单或桌面快捷方式启动
- 运行环境：Windows 10/11，系统需具备 Microsoft Edge WebView2 Runtime

## 截图

当前界面以类 Excel 的上下分栏方式展示原始数据和汇总结果，右侧提供分组列、求和列与生成操作。
