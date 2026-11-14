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
CURRENT_VERSION = "2.3.2"   # auto-updated by build.py

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


def _app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _log_check(msg: str):
    """Ghi đè log chẩn đoán lần check gần nhất (cạnh exe) — để debug máy
    không tự cập nhật được mà không cần rebuild lại app để thêm log."""
    try:
        import datetime
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_path = os.path.join(_app_dir(), "update_check.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def _ssl_context():
    """SSL context dùng CA bundle của certifi (đóng gói kèm app) thay vì phụ
    thuộc kho chứng chỉ Windows của máy — tránh CERTIFICATE_VERIFY_FAILED /
    "unable to get local issuer certificate" khi máy chưa cache đủ root cert
    (đã gặp thực tế: PowerShell/trình duyệt cùng máy verify OK nhưng app thì
    không, vì không tự trigger AIA fetch như native Windows networking)."""
    import ssl
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


# ── Public API ───────────────────────────────────────────────────────────────
def check_update_async(callback: Callable[[Optional[dict]], None]):
    """
    Kiểm tra bản mới trong background thread (không block UI).
    callback(info)  — info = dict nếu có bản mới, None nếu đã mới nhất hoặc lỗi.
    Luôn ghi log chẩn đoán (OK hoặc lỗi thật) vào update_check.log cạnh exe.
    """
    import urllib.request

    def _worker():
        try:
            req = urllib.request.Request(
                VERSION_URL,
                headers={"User-Agent": f"BOHO-IMPORT-BOM/{CURRENT_VERSION}"},
            )
            with urllib.request.urlopen(req, timeout=6, context=_ssl_context()) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            remote = data.get("version", "0")
            if _ver_tuple(remote) > _ver_tuple(CURRENT_VERSION):
                _log_check(f"OK - co ban moi {remote} (dang chay {CURRENT_VERSION})")
                callback(data)
            else:
                _log_check(f"OK - da la ban moi nhat ({CURRENT_VERSION})")
                callback(None)
        except Exception as e:
            _log_check(f"LOI: {type(e).__name__}: {e}")
            callback(None)   # silent fail phía UI — không làm phiền user nếu offline

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
    with urllib.request.urlopen(req, timeout=120, context=_ssl_context()) as resp:
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


# Script cập nhật chạy trong PowerShell: hiện cửa sổ tiến trình trên taskbar (bấm vào để xem),
# đợi app tắt → giải nén → ghi đè file → mở lại app. Các __TOKEN__ được thay bằng đường dẫn thật.
_PS1_TEMPLATE = r"""Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$sync = [hashtable]::Synchronized(@{ Status = 'Đang chờ ứng dụng đóng...'; Done = $false; Ok = $true; Error = '' })

# Cửa sổ tiến trình — hiện trên taskbar, bấm vào để xem đang chạy
$form = New-Object System.Windows.Forms.Form
$form.Text            = 'BOHO IMPORT - Đang cập nhật'
$form.ClientSize      = New-Object System.Drawing.Size(420, 110)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.ShowInTaskbar   = $true
$form.TopMost         = $true
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 10)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Location = New-Object System.Drawing.Point(20, 20)
$lbl.Size     = New-Object System.Drawing.Size(380, 24)
$lbl.Text     = $sync.Status
$form.Controls.Add($lbl)

$bar = New-Object System.Windows.Forms.ProgressBar
$bar.Location = New-Object System.Drawing.Point(20, 56)
$bar.Size     = New-Object System.Drawing.Size(380, 22)
$bar.Style    = 'Marquee'
$bar.MarqueeAnimationSpeed = 30
$form.Controls.Add($bar)

# Không cho đóng khi đang cập nhật (thu nhỏ xuống taskbar vẫn được)
$form.Add_FormClosing({ if (-not $sync.Done) { $_.Cancel = $true } })

$work = {
    param($sync, $zipPath, $extractTmp, $appDir, $appExe, $dbCfgBak)
    $errLog = Join-Path $env:TEMP 'boho_update_error.log'
    try {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($appExe)
        $waited = 0
        while ((Get-Process -Name $name -ErrorAction SilentlyContinue) -and $waited -lt 15) {
            Start-Sleep -Seconds 1
            $waited++
        }
        Start-Sleep -Seconds 1

        $dbCfgSrc = Join-Path $appDir 'config\db_config.json'
        if (Test-Path $dbCfgSrc) { Copy-Item $dbCfgSrc $dbCfgBak -Force }

        $sync.Status = 'Đang giải nén bản mới...'
        if (Test-Path $extractTmp) { Remove-Item $extractTmp -Recurse -Force }
        Expand-Archive -Path $zipPath -DestinationPath $extractTmp -Force

        $sync.Status = 'Đang ghi đè file...'
        $inner = Get-ChildItem $extractTmp -Directory | Select-Object -First 1
        if ($inner) {
            Copy-Item (Join-Path $inner.FullName '*') $appDir -Recurse -Force -ErrorAction Stop
        }

        if (Test-Path $dbCfgBak) {
            Copy-Item $dbCfgBak $dbCfgSrc -Force
            Remove-Item $dbCfgBak -Force
        }
        Remove-Item $extractTmp -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        $sync.Status = 'Hoàn tất — đang mở lại ứng dụng...'
    } catch {
        $sync.Ok = $false
        $sync.Error = $_.ToString()
        Add-Content $errLog ("Update failed: " + $_.ToString())
    }
    $sync.Done = $true
}

$ps = [PowerShell]::Create()
$null = $ps.AddScript($work).AddArgument($sync).AddArgument('__ZIP__').AddArgument('__EXTRACT__').AddArgument('__APPDIR__').AddArgument('__APPEXE__').AddArgument('__DBBAK__')
$handle = $ps.BeginInvoke()

$form.Show()
while (-not $sync.Done) {
    $lbl.Text = $sync.Status
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 80
}
$lbl.Text = $sync.Status
$bar.Style = 'Continuous'
$bar.Value = 100

if ($sync.Ok) {
    $until = (Get-Date).AddMilliseconds(1200)
    while ((Get-Date) -lt $until) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 }
    Start-Process '__APPEXE__'
} else {
    $form.TopMost = $false
    [System.Windows.Forms.MessageBox]::Show("Cập nhật thất bại: " + $sync.Error + "`nXem log: " + (Join-Path $env:TEMP 'boho_update_error.log'), 'BOHO IMPORT', 'OK', 'Error') | Out-Null
}
$form.Close()
$ps.Dispose()
Remove-Item $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""


def apply_update(zip_path: str):
    """
    Ghi PowerShell updater script → launch ngầm → sys.exit().
    Script PS1:
      - Hiện cửa sổ tiến trình trên taskbar
      - Chờ app tắt
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

    ps1 = _PS1_TEMPLATE
    for key, val in (("__ZIP__", zip_path), ("__EXTRACT__", extract_tmp), ("__APPDIR__", app_dir),
                     ("__APPEXE__", app_exe), ("__DBBAK__", db_cfg_bak)):
        ps1 = ps1.replace(key, _ps(val))

    with open(ps1_path, "w", encoding="utf-8-sig") as f:
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
