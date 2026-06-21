; Inno Setup スクリプト — 給与自動計算（ユーザー単位インストール）
; ビルド: Inno Setup Compiler で本ファイルを開き [Build] → Output\給与自動計算_setup.exe
;
; 管理者権限不要・ユーザーごとにインストールできる構成。
; インストール先(LocalAppData\Programs)は書込可能なため、アプリの自動更新が機能する。

#define AppName "給与自動計算"
#define AppVersion "1.0.0"
#define AppExe "給与自動計算.exe"
#define Publisher "給与自動計算"

[Setup]
AppId={{8F3A2C10-6B4D-4E2A-9C1F-PAYROLL000001}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DisableDirPage=no
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=給与自動計算_setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成する"; GroupDescription: "追加アイコン:"

[Files]
Source: "dist\給与自動計算.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\{#AppName} をアンインストール"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "今すぐ起動する"; Flags: nowait postinstall skipifsilent

; アンインストール時、ユーザーデータ(%APPDATA%\給与自動計算)は残す（誤削除防止）。
; 完全削除したい場合は %APPDATA%\給与自動計算 を手動で削除する。
