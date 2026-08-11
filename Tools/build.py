"""
build.py — Đóng gói BOHO IMPORT BOM/THDM thành .exe (PyInstaller onedir)

Chạy:
    cd Tools
    python build.py            # bản Release (windowed, không console)
    python build.py --debug    # bản Debug (console window, verbose log)

Output:
    dist/BOHO_IMPORT_BOM_THDM/           ← chạy trực tiếp
    installer_output/BOHO_IMPORT_BOM_THDM_v2.1_Portable.zip  ← bản Portable
    installer_output/BOHO_ImportBOM_THDM_Setup_v2.1.exe       ← bản Installer (cần Inno Setup)
"""
import sys
import os
import shutil
import zipfile
import subprocess
import tempfile
import argparse
import json

# Đảm bảo console Windows in được ký tự UTF-8
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ── Cấu hình ─────────────────────────────────────────────────────────────────
APP_NAME     = "BOHO_IMPORT_BOM_THDM"
VERSION      = "2.2.5"
SPEC_FILE    = "BOHO_IMPORT_BOM_THDM.spec"
DIST_APP_DIR = os.path.join("dist", APP_NAME)
OUTPUT_DIR   = "installer_output"
PORTABLE_DIR = os.path.join(OUTPUT_DIR, f"{APP_NAME}_v{VERSION}_Portable")
PORTABLE_ZIP = f"{PORTABLE_DIR}.zip"

# Config files sẽ được copy vào thư mục dist SAU KHI build (không bundle vào exe)
CONFIG_FILES = [
    "CK_Mapping_v5.xlsx",   # Mapping column — user có thể update
    "db_config.json",        # DB connection — user tự điền
]

INNO_CANDIDATES = [
    r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    r"C:\Program Files\Inno Setup 6\ISCC.exe",
    os.path.expanduser(r"~\AppData\Local\Programs\Inno Setup 6\ISCC.exe"),
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def _banner(msg: str):
    print(f"\n{'─'*60}")
    print(f"  {msg}")
    print(f"{'─'*60}")

def _ok(msg):   print(f"  [OK]  {msg}")
def _err(msg):  print(f"  [LỖI] {msg}")
def _warn(msg): print(f"  [CẢNH BÁO] {msg}")
def _info(msg): print(f"  {msg}")


# ── Bước 1: Chạy PyInstaller ──────────────────────────────────────────────────
def run_pyinstaller(debug: bool = False):
    label = "[DEBUG]" if debug else "[1/4]"
    _banner(f"{label} PyInstaller → {DIST_APP_DIR}/")
    if debug:
        _info("Chế độ DEBUG: console window bật, --debug all")

    # Build vào thư mục TEMP để tránh AV (Cortex XDR) lock file trong project dir
    tmp_dir  = tempfile.mkdtemp(prefix="boho_build_")
    tmp_dist = os.path.join(tmp_dir, "dist")
    tmp_work = os.path.join(tmp_dir, "work")
    os.makedirs(tmp_dist, exist_ok=True)
    os.makedirs(tmp_work, exist_ok=True)
    _info(f"Temp build dir: {tmp_dir}")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        SPEC_FILE,
        "--noconfirm",
        "--distpath", tmp_dist,
        "--workpath", tmp_work,
    ]
    if debug:
        # Bật console window và verbose PyInstaller log (override spec console=False)
        cmd += ["--console", "--debug", "all"]
    _info(f"Lệnh: {' '.join(cmd)}")
    ret = subprocess.run(cmd)
    if ret.returncode != 0:
        _err("PyInstaller build thất bại!")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        sys.exit(1)

    tmp_app = os.path.join(tmp_dist, APP_NAME)
    exe_path = os.path.join(tmp_app, f"{APP_NAME}.exe")
    if not os.path.exists(exe_path):
        _err(f"Không tìm thấy {exe_path} — build có thể thất bại!")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        sys.exit(1)
    _ok(f"Build OK: {exe_path}")

    # Sao chép kết quả từ TEMP → dist/ thực (dùng dirs_exist_ok để không cần xóa trước)
    os.makedirs("dist", exist_ok=True)
    _info(f"Sao chép → {DIST_APP_DIR}")
    shutil.copytree(tmp_app, DIST_APP_DIR, dirs_exist_ok=True)
    _ok(f"Đã sao chép → {DIST_APP_DIR}")

    # Dọn temp
    shutil.rmtree(tmp_dir, ignore_errors=True)


# ── Bước 2: Copy config vào dist/ (kế bên .exe) ──────────────────────────────
def copy_config():
    _banner("[2/4] Copy config/ vào dist/BOHO_IMPORT_BOM_THDM/config/")
    config_src = "config"
    config_dst = os.path.join(DIST_APP_DIR, "config")
    os.makedirs(config_dst, exist_ok=True)

    all_ok = True
    for fname in CONFIG_FILES:
        src = os.path.join(config_src, fname)
        dst = os.path.join(config_dst, fname)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            _ok(fname)
        else:
            _warn(f"{fname} không tìm thấy tại {src}")
            all_ok = False

    if not all_ok:
        _warn("Thiếu file config — app có thể không kết nối được DB khi chạy lần đầu.")


# ── Bước 3: Tạo bản Portable (ZIP) ───────────────────────────────────────────
def make_portable_zip():
    _banner(f"[3/4] Portable ZIP → {PORTABLE_ZIP}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if os.path.exists(PORTABLE_DIR):
        shutil.rmtree(PORTABLE_DIR, ignore_errors=True)

    shutil.copytree(DIST_APP_DIR, PORTABLE_DIR, dirs_exist_ok=True)
    _ok(f"Sao chép dist → {PORTABLE_DIR}")

    if os.path.exists(PORTABLE_ZIP):
        os.remove(PORTABLE_ZIP)

    total_files = 0
    with zipfile.ZipFile(PORTABLE_ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for root, _dirs, files in os.walk(PORTABLE_DIR):
            for fname in files:
                fp = os.path.join(root, fname)
                arcname = os.path.relpath(fp, os.path.dirname(PORTABLE_DIR))
                zf.write(fp, arcname)
                total_files += 1

    size_mb = os.path.getsize(PORTABLE_ZIP) / 1024 / 1024
    _ok(f"{PORTABLE_ZIP}  ({size_mb:.1f} MB, {total_files} files)")


# ── Bước 0: Patch version trong SOURCE trước khi PyInstaller compile ──────────
def patch_source_versions():
    """
    Patch version strings trong source files TRƯỚC khi PyInstaller chạy.
    Đảm bảo EXE được compile với đúng VERSION, CURRENT_VERSION, APP_VERSION.
    """
    _banner("[0/4] Patch version trong source files")
    import re as _re
    import datetime as _dt

    # 1) version.json ở repo root
    vj_src = os.path.join(os.path.dirname(__file__), "..", "version.json")
    if os.path.exists(vj_src):
        with open(vj_src, encoding="utf-8") as f:
            vj = json.load(f)
        vj["version"] = VERSION
        vj["download_url"] = (
            f"https://github.com/davit-philn/boho-import-tool/"
            f"releases/download/v{VERSION}/{APP_NAME}_v{VERSION}_Portable.zip"
        )
        vj["release_date"] = _dt.date.today().isoformat()
        with open(vj_src, "w", encoding="utf-8") as f:
            json.dump(vj, f, ensure_ascii=False, indent=4)
        _ok(f"version.json: version = \"{VERSION}\"")
    else:
        _warn("Không tìm thấy version.json ở repo root.")

    # 2) services/updater.py — CURRENT_VERSION
    updater_path = os.path.join("services", "updater.py")
    if os.path.exists(updater_path):
        with open(updater_path, encoding="utf-8") as f:
            src = f.read()
        new_src = _re.sub(r'(CURRENT_VERSION\s*=\s*")[^"]*(")', rf'\g<1>{VERSION}\g<2>', src)
        if new_src != src:
            with open(updater_path, "w", encoding="utf-8") as f:
                f.write(new_src)
            _ok(f"updater.py: CURRENT_VERSION = \"{VERSION}\"")
        else:
            _info("updater.py: CURRENT_VERSION đã đúng.")

    # 3) services/utils.py — APP_VERSION (hiển thị trên title bar / splash)
    utils_path = os.path.join("services", "utils.py")
    if os.path.exists(utils_path):
        with open(utils_path, encoding="utf-8") as f:
            src = f.read()
        new_src = _re.sub(r'(APP_VERSION\s*=\s*")[^"]*(")', rf'\g<1>{VERSION}\g<2>', src)
        if new_src != src:
            with open(utils_path, "w", encoding="utf-8") as f:
                f.write(new_src)
            _ok(f"utils.py: APP_VERSION = \"{VERSION}\"")
        else:
            _info("utils.py: APP_VERSION đã đúng.")

    # 4) setup.iss — AppVersion
    iss_path = "setup.iss"
    if os.path.exists(iss_path):
        with open(iss_path, encoding="utf-8") as f:
            src = f.read()
        new_src = _re.sub(r'(#define AppVersion\s+")[^"]*(")', rf'\g<1>{VERSION}\g<2>', src)
        if new_src != src:
            with open(iss_path, "w", encoding="utf-8") as f:
                f.write(new_src)
            _ok(f"setup.iss: AppVersion = \"{VERSION}\"")


# ── Bước 3b: Copy version.json vào dist/ (SAU khi PyInstaller build xong) ────
def copy_version_json_to_dist():
    _banner("[3b] Copy version.json → dist/")
    vj_src = os.path.join(os.path.dirname(__file__), "..", "version.json")
    vj_dst = os.path.join(DIST_APP_DIR, "version.json")
    if os.path.exists(vj_src):
        shutil.copy2(vj_src, vj_dst)
        _ok(f"version.json → {vj_dst}  (v{VERSION})")
    else:
        _warn("Không tìm thấy version.json ở repo root — bỏ qua.")


# ── Bước 4: Inno Setup Installer (tùy chọn) ──────────────────────────────────
def run_inno_setup():
    _banner("[4/4] Inno Setup → Setup.exe (tùy chọn)")
    iscc = next((c for c in INNO_CANDIDATES if os.path.exists(c)), None)
    if not iscc:
        _info("Không tìm thấy Inno Setup 6 → bỏ qua bước này.")
        _info("Tải về tại: https://jrsoftware.org/isdl.php")
        return

    iss_file = "setup.iss"
    if not os.path.exists(iss_file):
        _warn("Không tìm thấy setup.iss → bỏ qua.")
        return

    _info(f"ISCC: {iscc}")
    ret = subprocess.run([iscc, iss_file])
    if ret.returncode != 0:
        _err("Inno Setup build thất bại!")
    else:
        setup_exe = os.path.join(OUTPUT_DIR, f"BOHO_ImportBOM_THDM_Setup_v{VERSION}.exe")
        _ok(f"{setup_exe}")


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Luôn chạy từ thư mục chứa build.py (Tools/)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    parser = argparse.ArgumentParser(description="Build BOHO IMPORT BOM/THDM")
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Build debug variant: bật console window + PyInstaller verbose log",
    )
    args = parser.parse_args()

    mode_tag = "DEBUG" if args.debug else "RELEASE"
    print(f"\n{'='*60}")
    print(f"  BOHO IMPORT BOM/THDM v{VERSION} — Build Script [{mode_tag}]")
    print(f"  Python {sys.version.split()[0]} | PyInstaller onedir")
    print(f"{'='*60}")

    patch_source_versions()        # Patch source TRƯỚC — EXE sẽ compile đúng version
    run_pyinstaller(debug=args.debug)
    copy_config()
    copy_version_json_to_dist()    # Copy version.json vào dist SAU khi build xong
    if not args.debug:
        make_portable_zip()
        run_inno_setup()
    else:
        _info("Bỏ qua ZIP/Installer trong debug build.")

    print(f"""
{'='*60}
  HOÀN THÀNH BUILD!

  Chạy trực tiếp:
    {DIST_APP_DIR}\\{APP_NAME}.exe

  Bản Portable (gửi cho user):
    {PORTABLE_ZIP}
    → Giải nén → chạy {APP_NAME}.exe

  Bản Installer (nếu có Inno Setup):
    {OUTPUT_DIR}\\BOHO_ImportBOM_THDM_Setup_v{VERSION}.exe
{'='*60}
""")
