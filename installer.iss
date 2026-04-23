; 초록멀티 v1.5 — Inno Setup 설치 스크립트
; 빌드: ISCC.exe installer.iss

#define AppName "초록멀티"
#define AppVersion "1.5.1"
#define AppVersionDisplay "v1.5"
#define AppPublisher "초록등대 동호회"
#define AppExeName "초록멀티 v1.5.exe"
#define AppDistName "초록멀티 v1.5"

[Setup]
AppId={{B3F9A4D2-5E7C-4F8B-A1D6-3E2C7B8F9A01}
AppName={#AppName} {#AppVersionDisplay}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersionDisplay}
AppPublisher={#AppPublisher}
VersionInfoVersion={#AppVersion}
DefaultDirName={autopf}\{#AppName} {#AppVersionDisplay}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=installer_out
OutputBaseFilename=초록멀티_v1.5_setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; 자동 업데이트: 실행 중인 초록멀티를 설치 프로그램이 종료시키고 (기본 yes)
; 설치 후 다시 실행하도록 Restart Manager 사용을 명시적으로 활성화.
CloseApplications=yes
RestartApplications=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} {#AppVersionDisplay}
SetupIconFile=data\icon.ico
; 한국어 안내 메시지 사용

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 만들기"; GroupDescription: "추가 아이콘:"

[Files]
Source: "dist\{#AppDistName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; WinFSP MSI 는 임시 폴더에 풀고 설치 후 자동 삭제 (dontcopy 는 설치 디렉토리에 남기지 않음)
Source: "installer_deps\winfsp.msi"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: NeedWinFsp

[Icons]
Name: "{group}\{#AppName} {#AppVersionDisplay}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName} {#AppVersionDisplay}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName} {#AppVersionDisplay}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; WinFSP 미설치 시 msiexec 으로 조용히 설치. /qb 는 "기본 UI, 취소 불가" — 진행 표시줄만
; 나오므로 사용자에게 설치 진행이 보임. UAC 권한이 필요하므로 runascurrentuser+shellexec 로
; 전달. MSIRESTARTMANAGERCONTROL=Disable 로 재시작 프롬프트 억제.
Filename: "msiexec.exe"; \
  Parameters: "/i ""{tmp}\winfsp.msi"" /qb /norestart MSIRESTARTMANAGERCONTROL=Disable"; \
  StatusMsg: "NAS 드라이브 마운트용 WinFSP 를 설치 중입니다..."; \
  Flags: runhidden waituntilterminated; \
  Check: NeedWinFsp

Filename: "{app}\{#AppExeName}"; \
  Description: "{cm:LaunchProgram,{#AppName} {#AppVersionDisplay}}"; \
  Flags: nowait postinstall skipifsilent

[Code]
function NeedWinFsp(): Boolean;
var
  WinFspDll32, WinFspDll64: String;
begin
  WinFspDll32 := ExpandConstant('{pf32}') + '\WinFsp\bin\winfsp-x64.dll';
  WinFspDll64 := ExpandConstant('{pf}') + '\WinFsp\bin\winfsp-x64.dll';
  // 둘 다 없을 때만 설치 필요
  Result := not (FileExists(WinFspDll32) or FileExists(WinFspDll64));
end;
