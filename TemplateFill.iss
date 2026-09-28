; TemplateFill 模板文档生成器 — Inno Setup 安装包脚本
; 编译：ISCC.exe TemplateFill.iss
; AppVersion 由 build_release.py 从 version.py 自动同步，勿手改

#define MyAppName "TemplateFill 模板文档生成器"
#define MyAppVersion "1.3.0"
#define MyAppPublisher "TemplateFill"

[Setup]
; 固定 AppId（本项目专属，勿与其他应用共用，勿改动）：保证升级/卸载识别同一程序
AppId={{9F70B630-8753-426E-AB00-83267504A600}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\TemplateFill
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=TemplateFill_{#MyAppVersion}_setup
; LZMA2 极限压缩 + 固实包：Qt DLL 压缩率约 50~60%
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
SetupIconFile=TemplateFill.ico
UninstallDisplayIcon={app}\TemplateFill.exe

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; GroupDescription: "附加任务:"; Flags: checkedonce

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
Source: "dist\TemplateFill\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\TemplateFill.exe"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\TemplateFill.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\TemplateFill.exe"; Description: "立即运行 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 注意：用户数据一律不删——输出文档在桌面、示例模板在 %USERPROFILE%\TemplateFill、
; .labels.json 在模板文件旁、.tplfill 项目文件由用户自选位置保存
