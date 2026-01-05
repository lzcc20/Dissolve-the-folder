# 1. 强制脚本以管理员权限运行
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# --- 语言选择界面 ---
Clear-Host
Write-Host " ==========================================" -ForegroundColor Cyan
Write-Host "      Language Selection / 语言选择        " -ForegroundColor White -BackgroundColor DarkCyan
Write-Host " ==========================================" -ForegroundColor Cyan
Write-Host "  1. 中文 (Chinese)"
Write-Host "  2. English"
Write-Host ""
$LangChoice = Read-Host " Please select (1/2) / 请选择"

# --- 语言资源包 ---
if ($LangChoice -eq "2") {
    $L = @{
        Title = "Folder Dissolver Management v1.0"
        Intro = "Folder Dissolver is a productivity tool for Windows. Extracts content to parent directory and deletes the container via right-click menu."
        Overview = "Software Overview"
        Version = "Current Version"
        Updated = "Last Updated"
        Env = "OS Environment"
        Loc = "Install Loc"
        Detected = "[Installed] Please select:"
        NotDetected = "[Not Installed] Please select:"
        Opt1 = "1. Install / Update"
        Opt2 = "2. Uninstall (Clean all)"
        Opt3 = "3. Exit"
        Input = " Enter number and press Enter"
        Uninstalling = "`n Uninstalling..."
        Un_Reg = " [+] Registry entry removed."
        Un_File = " [+] Script file removed."
        Done = " Done!"
        ExitKey = " Press [Enter] to exit"
        MenuName = "Dissolve Folder"
        Target = "Target"
        Path = "Path"
        Stats = "Statistics"
        Files = "Files"
        Dirs = "Folders"
        Size = "Size"
        Warn = "Warning: Content moves to parent; original folder deleted."
        Ops = "Action: Press Enter to Execute | Ctrl+C to Cancel"
        CountMsg = "Starting in {0} seconds..."
        Processing = "Processing..."
        Complete = "Folder Dissolved."
        InstDone = "Installation Complete! Context menu integrated."
    }
} else {
    $L = @{
        Title = "文件夹解散助手 管理工具 v1.0"
        Intro = "文件夹解散助手是一款专为 Windows 用户设计的增强型生产力工具。通过右键菜单集成，实现一键将文件夹内容提取至上级目录并自动销毁原文件夹。"
        Overview = "软件概况"
        Version = "当前版本"
        Updated = "最后更新"
        Env = "运行环境"
        Loc = "安装位置"
        Detected = "[检测到已安装] 请选择操作："
        NotDetected = "[检测到未安装] 请选择操作："
        Opt1 = "1. 立即安装 / 更新组件"
        Opt2 = "2. 彻底卸载 (清除脚本与右键菜单)"
        Opt3 = "3. 退出"
        Input = " 请输入数字并按回车"
        Uninstalling = "`n 正在卸载..."
        Un_Reg = " [+] 已移除右键菜单注册表项。"
        Un_File = " [+] 已移除 C 盘脚本文件。"
        Done = " 卸载完成！"
        ExitKey = " 按 [回车键] 退出"
        MenuName = "解散此文件夹"
        Target = "目标"
        Path = "路径"
        Stats = "统计"
        Files = "文件"
        Dirs = "目录"
        Size = "大小"
        Warn = "警告: 内容将移至父目录，原文件夹将删除。"
        Ops = "操作: 按 Enter 立即执行 | Ctrl+C 取消操作"
        CountMsg = "将在 {0} 秒后开始执行..."
        Processing = "正在处理中..."
        Complete = "文件夹已解散。"
        InstDone = "安装完成！右键菜单已集成。"
    }
}

# 2. 定义全局路径
$InstallDir = "C:\Scripts"
$ScriptPath = Join-Path $InstallDir "DissolveFolder.ps1"
$RegistryBase = "Registry::HKEY_CLASSES_ROOT\Directory\shell\DissolveFolder"

# --- 模式选择 + 概况显示界面 ---
Clear-Host
Write-Host " ==========================================" -ForegroundColor Cyan
Write-Host "      $($L.Title)      " -ForegroundColor White -BackgroundColor DarkCyan
Write-Host " ==========================================" -ForegroundColor Cyan
Write-Host " $($L.Intro)"
Write-Host " ------------------------------------------"
Write-Host " 🚀 $($L.Overview)" -ForegroundColor Yellow
Write-Host " $($L.Version)：v1.0 (Stable)"
Write-Host " $($L.Updated)：2026-01"
Write-Host " $($L.Env)：Windows 10 / 11"
Write-Host " $($L.Loc)：$InstallDir"
Write-Host " ==========================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $ScriptPath) {
    Write-Host " $($L.Detected)" -ForegroundColor Yellow
    Write-Host " $($L.Opt1)"
    Write-Host " $($L.Opt2)"
    Write-Host " $($L.Opt3)"
} else {
    Write-Host " $($L.NotDetected)" -ForegroundColor Yellow
    Write-Host " $($L.Opt1)"
    Write-Host " $($L.Opt3)"
}
Write-Host ""
$Choice = Read-Host "$($L.Input)"

# --- 逻辑处理 ---

if ($Choice -eq "3") { exit }

# 卸载逻辑
if ($Choice -eq "2") {
    Write-Host "$($L.Uninstalling)" -ForegroundColor Cyan
    if (Test-Path $RegistryBase) { Remove-Item -Path $RegistryBase -Recurse -Force }
    Write-Host "$($L.Un_Reg)" -ForegroundColor Green
    if (Test-Path $ScriptPath) { Remove-Item -Path $ScriptPath -Force }
    Write-Host "$($L.Un_File)" -ForegroundColor Green
    Write-Host "`n $($L.Done)" -ForegroundColor Yellow
    Read-Host "$($L.ExitKey)"
    exit
}

# 安装逻辑
if ($Choice -eq "1") {
    if (-not (Test-Path $InstallDir)) { New-Item -Path $InstallDir -ItemType Directory | Out-Null }

    $ScriptContent = @"
param([string]`$FolderPath)
`$Host.UI.RawUI.WindowTitle = "$($L.Title)"
try { `$Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(60, 22) } catch {}

if (-not (Test-Path -Path `$FolderPath -PathType Container)) { exit }

`$FolderName = Split-Path `$FolderPath -Leaf
`$ParentPath = Split-Path `$FolderPath -Parent
`$Files = Get-ChildItem -LiteralPath `$FolderPath -File
`$SubFolders = Get-ChildItem -LiteralPath `$FolderPath -Directory
`$TotalSize = (`$Files | Measure-Object -Property Length -Sum).Sum / 1KB

Clear-Host
Write-Host "`n ==========================================" -ForegroundColor Cyan
Write-Host "          $($L.Title)           " -ForegroundColor White -BackgroundColor DarkCyan
Write-Host " ==========================================" -ForegroundColor Cyan
Write-Host "  [$($L.Target)]: `$FolderName" -ForegroundColor Yellow
Write-Host "  [$($L.Path)]: `$FolderPath"
Write-Host " ------------------------------------------"
Write-Host "  [$($L.Stats)]:"
Write-Host "    - $($L.Files): `$(`$Files.Count)"
Write-Host "    - $($L.Dirs): `$(`$SubFolders.Count)"
Write-Host "    - $($L.Size): `$([Math]::Round(`$TotalSize, 2)) KB"
Write-Host " ------------------------------------------"
Write-Host "  [$($L.Warn)]" -ForegroundColor DarkYellow
Write-Host " =========================================="
Write-Host "`n  [$($L.Ops)]" -ForegroundColor Gray

`$Counter = 5
while (`$Counter -gt 0) {
    `$Msg = "$($L.CountMsg)" -f `$Counter
    Write-Host "`r  `$Msg " -NoNewline -ForegroundColor Red
    if (`$Host.UI.RawUI.KeyAvailable) {
        `$Key = `$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        if (`$Key.VirtualKeyCode -eq 13) { break }
    }
    Start-Sleep -Milliseconds 100
    `$SubCount += 100
    if (`$SubCount -ge 1000) { `$Counter--; `$SubCount = 0 }
}

Write-Host "`n`n  $($L.Processing)" -ForegroundColor Green
`$Items = Get-ChildItem -LiteralPath `$FolderPath
foreach (`$Item in `$Items) {
    `$DestPath = Join-Path -Path `$ParentPath -ChildPath `$Item.Name
    if (Test-Path -LiteralPath `$DestPath) {
        `$NewName = "`$(`$Item.BaseName)_`$(Get-Date -Format 'HHmmss')`$(`$Item.Extension)"
        `$DestPath = Join-Path -Path `$ParentPath -ChildPath `$NewName
    }
    Move-Item -LiteralPath `$Item.FullName -Destination `$DestPath -Force
}
Remove-Item -LiteralPath `$FolderPath -Force
Write-Host "  [$($L.Complete)]" -ForegroundColor Green
Start-Sleep -Seconds 2
"@

    $Utf8NoBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($ScriptPath, $ScriptContent, $Utf8NoBom)

    if (-not (Test-Path $RegistryBase)) { New-Item -Path $RegistryBase -Force | Out-Null }
    Set-ItemProperty -Path $RegistryBase -Name "(Default)" -Value "$($L.MenuName)"
    Set-ItemProperty -Path $RegistryBase -Name "Icon" -Value "powershell.exe"
    $CmdPath = Join-Path $RegistryBase "command"
    if (-not (Test-Path $CmdPath)) { New-Item -Path $CmdPath -Force | Out-Null }
    Set-ItemProperty -Path $CmdPath -Name "(Default)" -Value "powershell.exe -ExecutionPolicy Bypass -File `"$ScriptPath`" `"%1`""

    Write-Host "`n $($L.InstDone)" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""
    Read-Host " $($L.ExitKey)"
}