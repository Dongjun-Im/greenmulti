; 초록멀티 v1.5 — Inno Setup 설치 스크립트
; 빌드: ISCC.exe installer.iss

#define AppName "초록멀티"
#define AppVersion "1.5.2"
#define AppVersionDisplay "v1.5"
#define AppPublisher "초록등대 동호회"
#define AppExeName "초록멀티 v1.5.exe"
#define AppDistName "초록멀티 v1.5"

[Setup]
; AppId 는 절대 바꾸지 말 것 — 고정 GUID 로 유지해야 같은 제품으로 인식되어
; 이전 버전 위에 새 버전을 설치할 때 자동으로 기존 설치를 교체한다.
AppId={{B3F9A4D2-5E7C-4F8B-A1D6-3E2C7B8F9A01}
; 제품 이름 자체는 버전 없이 "초록멀티" 로 고정한다. 그래야 시작 메뉴와
; 바탕화면 바로가기, 시작 메뉴 그룹 이름이 버전마다 중복 생성되지 않는다.
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersionDisplay}
AppPublisher={#AppPublisher}
VersionInfoVersion={#AppVersion}
; 설치 폴더도 버전 없는 이름으로 — 업그레이드 시 같은 폴더에 덮어쓴다.
DefaultDirName={autopf}\{#AppName}
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
; 제어판의 "프로그램 추가/제거" 목록에는 버전을 함께 표시해 어느 버전이
; 설치되어 있는지 알 수 있게 한다.
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

[InstallDelete]
; ── 구 버전(v1.3, v1.4) 바로가기·실행 파일 잔해 정리 ──
; 이전 설치본은 AppName 에 버전이 포함돼 있어서 시작 메뉴 그룹·바탕화면·
; 실행 파일이 별도 이름으로 남는다. 새 버전 설치 전에 아래 항목을 지워
; "초록멀티 v1.4" / "초록멀티 v1.5" 같은 아이콘이 여러 개 공존하지 않게 한다.

; 설치 폴더 안의 구 실행 파일
Type: files; Name: "{app}\초록멀티 v1.3.exe"
Type: files; Name: "{app}\초록멀티 v1.4.exe"

; 시작 메뉴 구 그룹 폴더(내부 바로가기·uninstaller 링크 포함)
Type: filesandordirs; Name: "{commonprograms}\초록멀티 v1.3"
Type: filesandordirs; Name: "{commonprograms}\초록멀티 v1.4"
Type: filesandordirs; Name: "{userprograms}\초록멀티 v1.3"
Type: filesandordirs; Name: "{userprograms}\초록멀티 v1.4"

; 바탕화면 구 바로가기(전 사용자·현재 사용자 양쪽 모두)
Type: files; Name: "{commondesktop}\초록멀티 v1.3.lnk"
Type: files; Name: "{commondesktop}\초록멀티 v1.4.lnk"
Type: files; Name: "{userdesktop}\초록멀티 v1.3.lnk"
Type: files; Name: "{userdesktop}\초록멀티 v1.4.lnk"

[Icons]
; 아이콘 이름에서 버전을 뺀다 — 다음 업그레이드 때 아이콘이 중복되지 않음.
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

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
