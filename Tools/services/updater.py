"""
services/updater.py — Auto-update checker cho BOHO IMPORT BOM/THDM.

Flow:
  1. check_update_async(callback) → background thread fetch version.json từ GitHub
  2. Nếu có bản mới: callback(info_dict)
  3. download_update(url, progress_cb) → tải ZIP về %TEMP%
  4. apply_update(zip_path) → ghi PowerShell updater, sys.exit()
  5. PS1 chờ process tắt → extract → xcopy vào app dir → relaunch
"""
import os
import sys
import json
import threading
import tempfile
import subprocess
from typing import Callable, Optional

# ── Constants ────────────────────────────────────────────────────────────────
CURRENT_VERSION = "2.2"   # auto-updated by build.py

VERSION_URL = (
    "https://raw.githubusercontent.com/davit-philn/"
    "boho-import-tool/master/version.json"
)


# ── Helpers ──────────────────────────────────────────────────────────────────
def _ver_tuple(v: str):
    try:
        return tuple(int(x) for x in str(v).strip().split("."))
    except Exception:
        return (0,)


# ── Public API ───────────────────────────────────────────────────────────────
def check_update_async(callback: Callable[[Optional[dict]], None]):
    """
    Kiểm tra bản mới trong background thread (không block UI).
    callback(info)  — info = dict nếu có bản mới, None nếu đã mới nhất hoặc lỗi.
    """
    import urllib.request

    def _worker():
        try:
            req = urllib.request.Request(
                VERSION_URL,
                headers={"User-Agent": f"BOHO-IMPORT-BOM/{CURRENT_VERSION}"},
            )
            with urllib.request.urlopen(req, timeout=6) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            remote = data.get("version", "0")
            if _ver_tuple(remote) > _ver_tuple(CURRENT_VERSION):
                callback(data)
            else:
                callback(None)
        except Exception:
            callback(None)   # silent fail — không làm phiền user nếu offline

    threading.Thread(target=_worker, daemon=True).start()


def download_update(
    url: str,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> str:
    """
    Tải ZIP update về %TEMP%\\boho_update\\boho_new.zip.
    progress_cb(bytes_downloaded, total_bytes) — tùy chọn.
    Trả về đường dẫn file ZIP.
    """
    import urllib.request

    tmp_dir = os.path.join(tempfile.gettempdir(), "boho_update")
    os.makedirs(tmp_dir, exist_ok=True)
    zip_path = os.path.join(tmp_dir, "boho_new.zip")

    req = urllib.request.Request(
        url, headers={"User-Agent": f"BOHO-IMPORT-BOM/{CURRENT_VERSION}"}
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        total = int(resp.headers.get("Content-Length", 0))
        downloaded = 0
        with open(zip_path, "wb") as f:
            while True:
                buf = resp.read(65536)
                if not buf:
                    break
                f.write(buf)
                downloaded += len(buf)
                if progress_cb:
                    progress_cb(downloaded, total)

    return zip_path


def apply_update(zip_path: str):
    """
    Ghi PowerShell updater script → launch ngầm → sys.exit().
    Script PS1:
      - Chờ app tắt (2s)
      - Backup db_config.json (giữ credentials)
      - Extract ZIP vào temp
      - xcopy nội dung vào app dir (đè file cũ)
      - Phục hồi db_config.json
      - Relaunch EXE
    """
    if getattr(sys, "frozen", False):
        app_exe = sys.executable
    else:
        # Dev mode: giả lập đường dẫn
        app_exe = os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..", "main.py")
        )
    app_dir  = os.path.dirname(os.path.abspath(app_exe))
    ps1_path = os.path.join(tempfile.gettempdir(), "boho_updater.ps1")

    # Escape đường dẫn cho PowerShell
    def _ps(p):
        return p.replace("'", "''")

    extract_tmp = os.path.join(tempfile.gettempdir(), "boho_update_extract")
    db_cfg_bak  = os.path.join(tempfile.gettempdir(), "boho_db_config.bak")

    ps1 = f"""
Add-Type -AssemblyName System.Windows.Forms

# Tray icon thông báo "Đang cập nhật..."
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.Visible = $true
$notify.BalloonTipTitle = 'BOHO IMPORT - Dang cap nhat'
$notify.BalloonTipText  = 'Dang ap dung ban moi, vui long doi...'
$notify.ShowBalloonTip(15000)

# Chờ process tắt hẳn (tối đa 15s) trước khi copy đè EXE
$processName = [System.IO.Path]::GetFileNameWithoutExtension('{_ps(app_exe)}')
$waited = 0
while ((Get-Process -Name $processName -ErrorAction SilentlyContinue) -and $waited -lt 15) {{
    Start-Sleep -Seconds 1
    $waited++
}}
Start-Sleep -Seconds 1

$zipPath    = '{_ps(zip_path)}'
$extractTmp = '{_ps(extract_tmp)}'
$appDir     = '{_ps(app_dir)}'
$appExe     = '{_ps(app_exe)}'
$dbCfgSrc   = Join-Path $appDir 'config\\db_config.json'
$dbCfgBak   = '{_ps(db_cfg_bak)}'

# Backup db_config.json (credentials user)
if (Test-Path $dbCfgSrc) {{
    Copy-Item $dbCfgSrc $dbCfgBak -Force
}}

# Extract ZIP
if (Test-Path $extractTmp) {{ Remove-Item $extractTmp -Recurse -Force }}
try {{
    Expand-Archive -Path $zipPath -DestinationPath $extractTmp -Force
}} catch {{
    Add-Content '$env:TEMP\\boho_update_error.log' "Extract failed: $_"
    $notify.Dispose()
    exit 1
}}

# Copy nội dung subfolder đầu tiên → app dir
$inner = Get-ChildItem $extractTmp -Directory | Select-Object -First 1
if ($inner) {{
    try {{
        Copy-Item (Join-Path $inner.FullName '*') $appDir -Recurse -Force -ErrorAction Stop
    }} catch {{
        Add-Content "$env:TEMP\boho_update_error.log" "Copy failed: $_ | src=$($inner.FullName) dst=$appDir"
        $notify.BalloonTipTitle = 'BOHO IMPORT - Cap nhat that bai'
        $notify.BalloonTipText  = "Khong the ghi de file. Xem log: $env:TEMP\boho_update_error.log"
        $notify.ShowBalloonTip(8000)
        Start-Sleep -Seconds 3
        $notify.Dispose()
        exit 1
    }}
}}

# Phục hồi db_config.json
if (Test-Path $dbCfgBak) {{
    Copy-Item $dbCfgBak $dbCfgSrc -Force
    Remove-Item $dbCfgBak -Force
}}

# Dọn temp
Remove-Item $extractTmp -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $zipPath    -Force          -ErrorAction SilentlyContinue

# Thông báo hoàn tất
$notify.BalloonTipTitle = 'BOHO IMPORT - Cap nhat xong!'
$notify.BalloonTipText  = 'Dang khoi dong lai ung dung...'
$notify.ShowBalloonTip(4000)
Start-Sleep -Seconds 1
$notify.Dispose()

# Relaunch app
Start-Process $appExe

# Xóa script này
Remove-Item $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""

    with open(ps1_path, "w", encoding="utf-8") as f:
        f.write(ps1)

    subprocess.Popen(
        [
            "powershell",
            "-WindowStyle", "Hidden",
            "-ExecutionPolicy", "Bypass",
            "-File", ps1_path,
        ],
        creationflags=subprocess.CREATE_NO_WINDOW
        if sys.platform == "win32"
        else 0,
    )
    sys.exit(0)
