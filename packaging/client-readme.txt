ExcelEditTool - 测试说明

启动方式：
1. 双击 ExcelEditTool.exe。
2. 选择本地 .xlsx 文件。
3. 切换 sheet，确认表头行。
4. 选择分组列和求和列。
5. 可以双击原始预览单元格修改数据。
6. 点击“生成汇总”。
7. 可以双击汇总结果单元格继续修改。
8. 可以点击“复制结果”粘贴到 Excel，也可以点击“导出”选择保存位置。

导出文件命名：
原Excel文件名_汇总.xlsx

注意：
- 当前版本优先支持 .xlsx。
- 如果 Windows 首次启动提示安全拦截，请选择“仍要运行”。
- 安装包默认安装到 D:\ExcelEditTool；如果当前电脑没有 D 盘，则回退到当前用户的本地程序目录。
- 推荐使用 ExcelEditTool_Setup.exe 安装包；安装时会自动检测 Microsoft Edge WebView2 Runtime，缺失时会调用 Microsoft 官方 Bootstrapper 静默安装。
- 如果使用 zip 免安装版，仍依赖系统已安装 WebView2 Runtime；Windows 11 和较新的 Windows 10 通常已自带。
- 当前安装包随附的是 WebView2 Evergreen Bootstrapper，首次安装缺少 Runtime 时需要联网下载运行时。若要完全离线分发，可改为随包附带 Evergreen Standalone Installer，安装包体积会明显变大。
