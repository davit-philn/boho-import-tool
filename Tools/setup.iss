; =============================================================
;  setup.iss — Inno Setup script cho BOHO IMPORT BOM/THDM
;  Đặt file này tại: Tools\setup.iss
;  Chạy bằng: python build.py  hoặc  iscc setup.iss
; =============================================================

#define AppName    "BOHO IMPORT BOM-THDM"
#define AppVersion "2.2.5"
#define AppDir     "BOHO_IMPORT_BOM_THDM"
#define AppExe     "BOHO_IMPORT_BOM_THDM.exe"
#define Publisher  "BOHO"
#define ODBCFile   "redist\msodbcsql.msi"
#define HasODBC    FileExists(ODBCFile)

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
; Cài vào C:\BOM_Tool (không dùng Program Files — tránh lỗi quyền ghi config)
DefaultDirName=C:\BOM_Tool
DisableProgramGroupPage=yes
; Output
OutputDir=installer_output
OutputBaseFilename=BOHO_ImportBOM_THDM_Setup_v{#AppVersion}
; Nén tốt nhất
Compression=lzma2/ultra64
SolidCompression=yes
; Cần admin để cài ODBC Driver
PrivilegesRequired=admin
; Icon cho file installer
SetupIconFile=icon.ico

[Languages]
Name: "default"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; \
  Description: "Tạo shortcut trên Desktop"; \
  GroupDescription: "Shortcut:"; \
  Flags: checkedonce

[Files]
; ── Toàn bộ thư mục dist/BOHO_IMPORT_BOM_THDM/ (onedir build) ──────────────
;    Bao gồm _internal/ chứa Python runtime và tất cả dependencies
Source: "dist\{#AppDir}\*"; \
  DestDir: "{app}"; \
  Flags: ignoreversion recursesubdirs createallsubdirs

; ── Config: CK_Mapping_v5.xlsx — luôn update khi nâng cấp ─────────────────
Source: "config\CK_Mapping_v5.xlsx"; \
  DestDir: "{app}\config"; \
  Flags: ignoreversion

; ── db_config.json — CHỈ copy lần đầu, KHÔNG ghi đè ───────────────────────
;    Tránh xóa DB config user đã cấu hình khi cập nhật tool
Source: "config\db_config.json"; \
  DestDir: "{app}\config"; \
  Flags: ignoreversion onlyifdoesntexist

; ── ODBC Driver 17 (chỉ bundle nếu file tồn tại lúc build) ─────────────────
;    Tải về tại: https://aka.ms/downloadmsodbcsql
;    Đặt vào: Tools\redist\msodbcsql.msi
#if HasODBC
Source: "{#ODBCFile}"; \
  DestDir: "{tmp}"; \
  Flags: deleteafterinstall; \
  Check: NeedODBCDriver
#endif

[Icons]
; Start Menu
Name: "{userprograms}\{#AppName}\{#AppName}"; \
  Filename: "{app}\{#AppExe}"

; Desktop shortcut (tùy chọn)
Name: "{userdesktop}\{#AppName}"; \
  Filename: "{app}\{#AppExe}"; \
  Tasks: desktopicon

[Run]
; ── Cài ODBC Driver 17 silent (chỉ chạy nếu đã bundle + chưa cài) ──────────
#if HasODBC
Filename: "msiexec.exe"; \
  Parameters: "/i ""{tmp}\msodbcsql.msi"" /quiet IACCEPTMSODBCSQLLICENSETERMS=YES"; \
  StatusMsg: "Đang cài Microsoft ODBC Driver 17 for SQL Server..."; \
  Flags: waituntilterminated runhidden; \
  Check: NeedODBCDriver
#endif

; ── Mở app sau khi cài xong ─────────────────────────────────────────────────
Filename: "{app}\{#AppExe}"; \
  Description: "Mở BOHO Import BOM/THDM ngay bây giờ"; \
  Flags: nowait postinstall skipifsilent

[Code]
// Trả về True nếu ODBC Driver 17 CHƯA được cài
function NeedODBCDriver: Boolean;
begin
  Result := not RegKeyExists(
    HKLM,
    'SOFTWARE\ODBC\ODBCINST.INI\ODBC Driver 17 for SQL Server'
  );
end;
