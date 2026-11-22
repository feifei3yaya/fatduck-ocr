# -*- coding: utf-8 -*-
"""解除目录内文件的 Windows 下载锁定（Zone.Identifier）。

从网盘/IM 下载的文件会被 Windows 标记为"来自互联网"，首次运行触发
SmartScreen 蓝色警告。本脚本清除该标记，不影响文件本身。

用法（由 解除锁定.bat 调用，也可手动执行）：
    powershell -NoProfile -ExecutionPolicy Bypass -File unblock_files.ps1
"""
$dir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

Write-Host "============================================"
Write-Host "  肥鸭文档转换 - 解除 Windows 下载锁定"
Write-Host "============================================"
Write-Host ""
Write-Host "本工具未购买代码签名证书，从网盘/IM 下载的文件会被 Windows"
Write-Host "标记为`"来自互联网`"，首次运行弹出蓝色 SmartScreen 警告。"
Write-Host "本脚本解除该标记（不影响文件本身），解压后运行一次即可。"
Write-Host ""

try {
    Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction Stop
    Write-Host "[完成] 已解除本目录及子目录全部文件的安全锁定。"
}
catch {
    Write-Host "[失败] 自动解除未成功：$_"
    Write-Host "        请手动操作：右键 肥鸭文档转换.exe - 属性 - 勾选`"解除锁定`" - 确定。"
}

Write-Host ""
Read-Host "按回车键关闭"
