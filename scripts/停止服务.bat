@echo off
echo 正在停止 CLIProxyAPI 服务...
taskkill /f /im cli-proxy-api.exe >nul 2>&1
echo CLIProxyAPI 已成功停止！
ping 127.0.0.1 -n 2 >nul
