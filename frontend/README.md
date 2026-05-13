# ExcelEditTool Frontend

这是 ExcelEditTool 的 React + Tauri 客户端。Excel 文件在本机前端用 ExcelJS 处理，不需要启动后端服务。

## 开发

```powershell
npm install
npm run dev -- --host 127.0.0.1 --port 5173
```

启动 Tauri 开发客户端：

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
..\scripts\package_tauri_client.ps1
```
