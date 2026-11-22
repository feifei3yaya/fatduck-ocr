; -*- coding: utf-8 -*-
; 肥鸭文档转换 —— Inno Setup 安装脚本（全中文界面）。
;
; 编译（需 Inno Setup 6/7）：
;     ISCC.exe installer.iss
; 产物：dist\肥鸭文档转换_安装包_v{#MyAppVersion}.exe
;
; 前置：dist\肥鸭文档转换.exe 已由 build_exe.py 生成；
;       runtime\ 由 build_runtime.py 生成；
;       models\PP-OCRv6 为项目内置识别模型（OpenVINO IR）；
;
; 安装布局与便携包一致（ocr_worker 按脚本同级目录定位模型与字库）：
;     {app}\肥鸭文档转换.exe
;     {app}\ocr_worker.py
;     {app}\runtime\...            （CPU 推理环境）
;     {app}\models\PP-OCRv6\...    （det/rec OpenVINO 模型，主用）
;     {app}\assets\ppocrv6_dict.txt / ppocrv5_dict.txt
;
; 语言：安装向导与卸载程序均使用简体中文（installer_zh.isl）。

#define MyAppName "肥鸭文档转换"
#include "version.iss"
#define MyAppPublisher "Fat Duck OCR Contributors"
#define MyAppExeName "肥鸭文档转换.exe"
#define MyAppAssumed "肥鸭文档转换"

[Setup]
AppId={{B8F4D2A1-6C3E-4F7B-9A2D-5E8C1B7F3D46}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
VersionInfoVersion={#MyAppVersion4}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} {#MyAppVersion} —— 本地 OCR 图片/PDF 转 Excel 工具
VersionInfoCopyright=Copyright (c) 2026 Fat Duck OCR Contributors
AppCopyright=Copyright (c) 2026 Fat Duck OCR Contributors
DefaultDirName={autopf}\{#MyAppAssumed}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=肥鸭文档转换_安装包_v{#MyAppVersion}
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName} {#MyAppVersion}
LicenseFile=LICENSE
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
ShowLanguageDialog=no
Uninstallable=yes

[Languages]
Name: chinesesimplified; MessagesFile: "installer_zh.isl"

[Tasks]
Name: desktopicon; Description: 创建桌面快捷方式; GroupDescription: 附加任务:; Flags: unchecked

[Files]
Source: dist\{#MyAppExeName}; DestDir: {app}; Flags: ignoreversion
Source: ocr_worker.py; DestDir: {app}; Flags: ignoreversion
Source: VERSION; DestDir: {app}; Flags: ignoreversion
Source: runtime\*; DestDir: {app}\runtime; Flags: ignoreversion recursesubdirs createallsubdirs
Source: models\PP-OCRv6\det\*.xml; DestDir: {app}\models\PP-OCRv6\det; Flags: ignoreversion
Source: models\PP-OCRv6\det\*.bin; DestDir: {app}\models\PP-OCRv6\det; Flags: ignoreversion
Source: models\PP-OCRv6\rec\*.xml; DestDir: {app}\models\PP-OCRv6\rec; Flags: ignoreversion
Source: models\PP-OCRv6\rec\*.bin; DestDir: {app}\models\PP-OCRv6\rec; Flags: ignoreversion
Source: assets\ppocrv6_dict.txt; DestDir: {app}\assets; Flags: ignoreversion
Source: assets\ppocrv5_dict.txt; DestDir: {app}\assets; Flags: ignoreversion
Source: LICENSE; DestDir: {app}; Flags: ignoreversion
Source: privacy_notice.txt; DestDir: {app}; Flags: ignoreversion
Source: 使用说明_SmartScreen.txt; DestDir: {app}; Flags: ignoreversion
Source: tools\unblock_dist.bat; DestDir: {app}; DestName: 解除锁定.bat; Flags: ignoreversion
Source: tools\unblock_files.ps1; DestDir: {app}; Flags: ignoreversion

[Icons]
Name: {group}\{#MyAppName}; Filename: {app}\{#MyAppExeName}
Name: {group}\使用说明; Filename: {app}\使用说明_SmartScreen.txt
Name: {group}\卸载 {#MyAppName}; Filename: {uninstallexe}
Name: {commondesktop}\{#MyAppName}; Filename: {app}\{#MyAppExeName}; Tasks: desktopicon

[Run]
Filename: {app}\{#MyAppExeName}; Description: 立即运行 {#MyAppName}; Flags: nowait postinstall skipifsilent

[Code]

{ 卸载时询问是否一并删除本地配置与日志。
  配置在 %APPDATA%\肥鸭文档转换，日志在 %LOCALAPPDATA%\肥鸭文档转换\logs。
  任务历史（jobs）与 WebView2 缓存不在此处删除。 }
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
    answer: Integer;
begin
    if CurUninstallStep = usPostUninstall then
    begin
        answer := MsgBox('是否同时删除本地配置与日志？' #13#13
            + '配置：%APPDATA%\肥鸭文档转换' #13
            + '日志：%LOCALAPPDATA%\肥鸭文档转换\logs' #13#13
            + '选择「是」一并删除；选择「否」保留（换机重装后可继续沿用设置）。',
            mbConfirmation, MB_YESNO or MB_DEFBUTTON2);
        if answer = IDYES then
        begin
            DelTree(ExpandConstant('{userappdata}\肥鸭文档转换'), True, True, True);
            DelTree(ExpandConstant('{localappdata}\肥鸭文档转换\logs'), True, True, True);
        end;
    end;
end;
