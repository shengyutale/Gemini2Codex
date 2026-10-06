# Codex 配置与使用 Gemini 大模型完整保姆级指南

> **前言**：本教程专为零代码基础用户编写。通过 CLI Proxy API + CC Switch，让官方 Codex 能够一键接入 Google Gemini 系列大模型，并在顶部下拉菜单中直接点选。

---

## 一、核心原理（大白话通俗比喻）

整套方案由四个部分协同运作，如同给外国电器配电源转换插头：

```
[Google 官方 Gemini]  <---(需要科学上网梯子)
         │
[CLI Proxy API]       <---【后台转接头】静默开机自启，把 Google 接口翻译成 Codex 能听懂的语言
         │
   [CC Switch]        <---【可视化总控管家】图形化界面，一键把转接头参数写进 Codex 并生成菜单
         │
     [Codex]          <---【工作台】在界面顶部的下拉三角里自由选择并使用 Gemini
```

---

## 二、准备工作与下载地址

1. **科学上网环境**：开启本地代理软件（如 Clash / v2rayN 等），确保本地端口开启（常见为 `7890` 或 `10809`）。
2. **两款必备工具下载**：
   * **CLI Proxy API（转接头）**：
     * 官方开源项目：[GitHub - CLIProxyAPI](https://github.com/router-for-me/CLIProxyAPI)
     * 下载地址：[CLIProxyAPI Releases 页面](https://github.com/router-for-me/CLIProxyAPI/releases)
     * 下载说明：Windows 系统请下载 `CLIProxyAPI_x.x.xx_windows_amd64.zip`，解压到固定文件夹（如 `D:\CLIProxyAPI_8.0.16_windows_amd64\`）。
   * **CC Switch（可视化总控管家）**：
     * 官方开源项目：[GitHub - cc-switch](https://github.com/farion1231/cc-switch)
     * 下载地址：[CC Switch Releases 页面](https://github.com/farion1231/cc-switch/releases)
     * 下载说明：Windows 系统下载安装包 `CC-Switch-vx.xx.x-Windows.msi`，双击默认下一步安装完成，桌面会出现 `CC Switch` 图标。
3. **Google 账号**：拥有一个正常可用的 Google 邮箱账号。

---

## 三、完整配置四步曲

### 第一步：配置转接头参数（config.yaml）

1. 打开文件夹：`D:\CLIProxyAPI_8.0.16_windows_amd64\`。
2. 找到 `config.yaml` 文件，右键 -> 【打开方式】 -> 【记事本】。
3. 检查并确保三处核心参数正确填写：
   * **服务端口（门牌号，默认 8317）**：
     ```yaml
     server:
       port: 8317
     ```
   * **自定义通行暗号（给 Codex 用的密码）**：
     ```yaml
     access:
       api-keys:
         - "local-gemini-key"
     ```
   * **网络代理（梯子地址与端口）**：
     ```yaml
     requests:
       proxy-url: "http://127.0.0.1:7890"
     ```
4. 按 `Ctrl + S` 保存并关闭记事本。

---

### 第二步：授权 Google 账号

> 每个 Google 账号**只需授权一次**，授权后凭证长期保存在本地，永久有效。

1. 键盘按下快捷键 `Win + R`，输入 `cmd` 按回车打开命令窗口。
2. 复制并粘贴以下命令，按回车进入程序目录：
   ```cmd
   cd /d D:\CLIProxyAPI_8.0.16_windows_amd64
   ```
3. 复制并粘贴以下命令，按回车发起授权：
   ```cmd
   .\cli-proxy-api.exe -antigravity-login -config .\config.yaml
   ```
4. 电脑默认浏览器会自动弹出 Google 授权登录页面（若未弹出，把黑框里的网址复制到浏览器打开）。
5. 登录您的 Google 账号，点击【允许】/【授权】。
6. 看到黑框提示登录成功后关闭黑框即可。登录凭证会自动保存在电脑的 `C:\Users\用户名\.cli-proxy-api\` 中。

---

### 第三步：配置后台静默自启（彻底告别手动命令与黑框）

在没有配置自启优化前，原先每次启动都需要手动敲以下命令：
```cmd
cd /d D:\CLIProxyAPI_8.0.16_windows_amd64
.\cli-proxy-api.exe -config .\config.yaml
```
手动启动不仅容易误关黑色窗口导致服务中断，而且每次电脑开机都必须手动敲一遍。现在通过全自动脚本实现一键配置：

本项目在 `scripts/` 目录下提供了全套自动化脚本，只需双击运行 **`【一键配置】开机自启与桌面图标.vbs`**，即可全自动完成后台静默运行、开机自启以及桌面快捷图标配置。

#### `【一键配置】开机自启与桌面图标.vbs` 代码展示：
```vbscript
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

targetDir = "D:\CLIProxyAPI_8.0.16_windows_amd64"
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' 1. 将后台启动脚本与停止脚本复制到 D 盘程序目录
vbsSource = fso.BuildPath(scriptDir, "启动_后台静默.vbs")
vbsDest = fso.BuildPath(targetDir, "启动_后台静默.vbs")
If fso.FileExists(vbsSource) Then
    fso.CopyFile vbsSource, vbsDest, True
End If

stopSource = fso.BuildPath(scriptDir, "停止服务.bat")
stopDest = fso.BuildPath(targetDir, "停止服务.bat")
If fso.FileExists(stopSource) Then
    fso.CopyFile stopSource, stopDest, True
End If

' 2. 创建开机自启动快捷方式
startupFolder = ws.SpecialFolders("Startup")
Set lnkStartup = ws.CreateShortcut(fso.BuildPath(startupFolder, "CLIProxyAPI_静默自启.lnk"))
lnkStartup.TargetPath = vbsDest
lnkStartup.WorkingDirectory = targetDir
lnkStartup.WindowStyle = 0
lnkStartup.IconLocation = fso.BuildPath(targetDir, "cli-proxy-api.exe") & ",0"
lnkStartup.Save

' 3. 创建桌面快捷方式
desktopFolder = ws.SpecialFolders("Desktop")
Set lnkDesktop = ws.CreateShortcut(fso.BuildPath(desktopFolder, "CLIProxyAPI (双击静默启动).lnk"))
lnkDesktop.TargetPath = vbsDest
lnkDesktop.WorkingDirectory = targetDir
lnkDesktop.WindowStyle = 0
lnkDesktop.IconLocation = fso.BuildPath(targetDir, "cli-proxy-api.exe") & ",0"
lnkDesktop.Save

msg = "CLIProxyAPI 开机自启与桌面图标配置成功！" & vbCrLf & vbCrLf
msg = msg & "1. 电脑开机后将在后台自动静默运行（无黑框）。" & vbCrLf
msg = msg & "2. 桌面上已生成【CLIProxyAPI (双击静默启动)】快捷方式。" & vbCrLf
msg = msg & "3. 如需停止服务，运行 D 盘程序目录下的【停止服务.bat】即可。" & vbCrLf & vbCrLf
msg = msg & "现在 Codex 已可直接从下拉菜单选择 Gemini 模型使用！"
MsgBox msg, 64, "配置成功"
```

1. **静默运行原理**：通过 VBScript 脚本调用程序，彻底隐藏黑色控制台窗口。
2. **开机自启**：将启动脚本放入 Windows 系统自启文件夹（`shell:startup`），开机即在后台静默运行。
3. **桌面快捷控制**：
   * **启动快捷方式**：桌面上双击【CLIProxyAPI (双击静默启动)】即可在后台无感运行。
   * **停止快捷方式**：运行【停止服务.bat】即可一键彻底关闭后台转接头进程。
4. **如何验证程序是否在运行**：
   * 在任意浏览器打开：`http://localhost:8317/`，只要能打开页面就说明正在运行；
   * 或者打开任务管理器（`Ctrl + Shift + Esc`），在“后台进程”中能看到 `cli-proxy-api.exe`。

---

### 第四步：在 CC Switch 中一键对接 Codex

1. 双击桌面上的 **【CC Switch】** 快捷方式打开软件界面。
2. 在软件左侧或顶部选择目标软件为 **Codex**。
3. 点击 **【添加服务商】**（或编辑现有服务商），填写如下信息：
   * **服务商名称**：`Gemini CLI Proxy API`
   * **API 接口地址 (Base URL)**：`http://localhost:8317/v1`
   * **API 密钥 (API Key)**：`local-gemini-key`
4. 在页面下方的 **【模型列表 (Model Catalog)】** 中添加需要使用的模型：
   * `gemini-3.8-flash-high`（主力高智商思考模型）
   * `gemini-3.5-flash-lite`（轻量极速模型）
   * `gemini-pro-agent`（专用智能体模型）
5. 点击 **【保存】**，然后在卡片上点击 **【启用】**（或设为当前）。
6. **底层自动完成**：
   CC Switch 会自动生成 `C:\Users\用户名\.codex\cc-switch-model-catalog.json` 并修改 `config.toml`，使 Codex 具备下拉菜单能力。

---

## 四、在 Codex 中切换与使用

1. 打开 **Codex** 桌面应用。
2. 点击界面顶部的**模型选择框（下拉三角图标）**。
3. 在展开的列表中，直接点击勾选 **`gemini-3.8-flash-high`**。
4. 现在输入任何问题，Codex 就会直接调用 Google 的 Gemini 大模型为您服务！

---

## 五、日常使用技巧与常见排查（FAQ）

### Q1：提问时提示网络错误或连接超时？
* **自查 1（梯子）**：检查电脑科学上网软件是否正常开启，节点是否通畅。
* **自查 2（转接头）**：在浏览器打开 `http://localhost:8317/`，如果无法打开，说明转接头未启动，双击桌面的【CLIProxyAPI (双击静默启动)】即可恢复。

### Q2：我想切回原版的 OpenAI 官方 GPT 模型怎么办？
* 打开桌面上的 **CC Switch** 软件，点击选中【OpenAI Official】或默认服务商即可一键切回。
* 以后想再用 Gemini，点回【Gemini CLI Proxy API】即可。

### Q3：Google 账号失效或需要换账号怎么办？
* 重新执行【第二步】中的 `.\cli-proxy-api.exe -antigravity-login -config .\config.yaml` 命令，在浏览器中登录新账号授权即可，新凭证会自动覆盖旧凭证。

---

## 六、关键路径与文件速查表

| 功能类别 | 文件 / 文件夹完整路径 | 说明 |
| :--- | :--- | :--- |
| **转接头程序** | `D:\CLIProxyAPI_8.0.16_windows_amd64\cli-proxy-api.exe` | 核心转接程序 |
| **转接头配置** | `D:\CLIProxyAPI_8.0.16_windows_amd64\config.yaml` | 端口与代理设置 |
| **Google凭证** | `C:\Users\用户名\.cli-proxy-api\` | Google 授权密钥存储目录 |
| **总控管家程序** | `C:\Users\用户名\AppData\Local\Programs\CC Switch\cc-switch.exe` | 可视化切换软件 |
| **桌面快捷方式** | `C:\Users\用户名\Desktop\CC Switch.lnk` | CC Switch 桌面图标 |
| **Codex核心配置** | `C:\Users\用户名\.codex\config.toml` | Codex 软件主配置文件 |
| **Codex模型菜单** | `C:\Users\用户名\.codex\cc-switch-model-catalog.json` | 下拉三角模型数据源 |
| **完整指南文档** | `C:\Users\用户名\Documents\Codex\Wiki\Codex配置Gemini大模型完整指南.md` | 本地 Markdown 手册 |
