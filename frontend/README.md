# ExcelEditTool Frontend

ExcelEditTool 的前端和桌面客户端代码位于此目录，界面使用 React + TypeScript + Fluent UI，Excel 读写和汇总处理使用 ExcelJS，桌面壳使用 Tauri。

## 开发

```powershell
npm install
npm run dev -- --host 127.0.0.1 --port 5173
```

启动桌面开发客户端：

```powershell
npm run tauri
```

## 验证

```powershell
npm test
npm run lint
npm run build
```

## 打包

从项目根目录运行：

```powershell
.\scripts\package_tauri_client.ps1
```
