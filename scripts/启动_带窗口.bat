@echo off
cd /d "D:\CLIProxyAPI_8.0.16_windows_amd64"
echo 正在启动 CLIProxyAPI (带窗口调试模式)...
echo 如果需要关闭，直接关闭本窗口或双击【停止服务.bat】即可。
echo.
"D:\CLIProxyAPI_8.0.16_windows_amd64\cli-proxy-api.exe" -config "D:\CLIProxyAPI_8.0.16_windows_amd64\config.yaml"
pause
