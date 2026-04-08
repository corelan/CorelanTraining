# -*- coding: utf-8 -*-
from __future__ import print_function

"""
corelanWin7VMinstall.py

(c) Corelan Consulting bv - 2026
www.corelan-consulting.com
www.corelan-training.com
www.corelan-certified.com
www.corelan.be

Python 2/3 compatible installer helper for older Windows 7 SP1 VMs.

IMPORTANT:
- Run this script as Administrator.
- Assumes Python 2 is already available to run this script.
- Uses only Python standard library functionality for downloads/extraction.
"""

import os
import sys
import ssl
import ctypes
import shutil
import zipfile
import socket
import subprocess
import traceback

try:
    import winreg
except ImportError:
    import _winreg as winreg

try:
    from urllib.request import urlopen, Request
except ImportError:
    from urllib2 import urlopen, Request

try:
    text_type = unicode
except NameError:
    text_type = str

SCRIPT_NAME = "corelanWin7VMinstall.py"
TEMP_FOLDER = r"C:\corelantemp"
SYMBOL_PATH = r"srv*c:\symbols*http://msdl.microsoft.com/download/symbols"

PYTHON2_X86_URL = "https://www.python.org/ftp/python/2.7.18/python-2.7.18.msi"
PYTHON2_X64_URL = "https://www.python.org/ftp/python/2.7.18/python-2.7.18.amd64.msi"
PYTHON32_URL = "https://www.python.org/ftp/python/3.8.10/python-3.8.10.exe"
PYTHON64_URL = "https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe"
WINDBG_URL = "https://go.microsoft.com/fwlink/p/?LinkId=323507"
MONA_URL = "https://github.com/corelan/mona/raw/master/mona.py"
WINDBGLIB_URL = "https://github.com/corelan/windbglib/raw/master/windbglib.py"
PYKD_EXT_X86_URL = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x86.zip"
PYKD_EXT_X64_URL = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x64.zip"
VCREDIST_X86_URL = "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe"
VCREDIST_X64_URL = "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe"
DOTNET_URL = "https://go.microsoft.com/fwlink/?linkid=2088631"
SYMBOLS_ZIP_URL = "https://www.corelan-training.com/downloads/win7symbols.zip"

PYTHON2_X86_INSTALLER = "python-2.7.18.msi"
PYTHON2_X64_INSTALLER = "python-2.7.18.amd64.msi"
PYTHON32_INSTALLER = "python-3.8.10.exe"
PYTHON64_INSTALLER = "python-3.8.10-amd64.exe"
WINDBG_INSTALLER = "sdksetup.exe"
MONA_FILE = "mona.py"
WINDBGLIB_FILE = "windbglib.py"
PYKD_EXT_X86_ZIP = "pykd-ext-x86.zip"
PYKD_EXT_X64_ZIP = "pykd-ext-x64.zip"
VCREDIST_X86_FILE = "vcredist_x86.exe"
VCREDIST_X64_FILE = "vcredist_x64.exe"
DOTNET_FILE = "NPD48-x86-x64-AllOS-ENU.exe"
SYMBOLS_ZIP_FILE = "win7symbols.zip"

LOCALAPPDATA = os.environ.get("LOCALAPPDATA", r"C:\Users\Default\AppData\Local")
ENGINE_EXT_64 = os.path.join(LOCALAPPDATA, "DBG", "EngineExtensions")
ENGINE_EXT_32 = os.path.join(LOCALAPPDATA, "DBG", "EngineExtensions32")

PYTHON27_X86_ROOT = r"C:\Python27"
PYTHON27_X64_ROOT = r"C:\Python27-64"
PYTHON38_X86_ROOT = os.path.join(LOCALAPPDATA, "Programs", "Python", "Python38-32")
PYTHON38_X64_ROOT = os.path.join(LOCALAPPDATA, "Programs", "Python", "Python38")

PROGRAM_FILES_X86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
PROGRAM_FILES = os.environ.get("ProgramFiles", r"C:\Program Files")
WINDIR = os.environ.get("WINDIR", r"C:\Windows")

REGSVR32_64 = os.path.join(WINDIR, "System32", "regsvr32.exe")
REGSVR32_32 = os.path.join(WINDIR, "SysWOW64", "regsvr32.exe")


def log(msg):
    print(msg)
    sys.stdout.flush()


def abort(msg, code=1):
    log("*** " + msg)
    raise SystemExit(code)


def ensure_dir(path):
    if path and not os.path.isdir(path):
        os.makedirs(path)


def remove_file(path):
    try:
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def remove_tree(path):
    try:
        if os.path.isdir(path):
            shutil.rmtree(path)
    except Exception:
        pass


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def confirm_continue(message):
    while True:
        try:
            answer = raw_input(message + " (Y/N) ")
        except NameError:
            answer = input(message + " (Y/N) ")
        answer = (answer or "").strip().lower()
        if answer in ("y", "yes"):
            return
        if answer in ("n", "no"):
            abort("Aborted by user.")
        log("Please enter Y or N.")


def safe_step(step_name, func, *args, **kwargs):
    log("\n[+] {0}".format(step_name))
    try:
        result = func(*args, **kwargs)
        if result is None:
            return True
        return result
    except SystemExit as e:
        log("    FAILED: {0}".format(e))
    except Exception as e:
        log("    FAILED: {0}".format(e))
        try:
            tb = traceback.format_exc()
            if tb:
                for line in tb.splitlines():
                    log("    " + line)
        except Exception:
            pass
    return False


def get_ssl_context():
    try:
        if hasattr(ssl, "SSLContext"):
            try:
                # Prefer the most permissive practical setup for old Win7 boxes.
                if hasattr(ssl, "PROTOCOL_TLSv1_2"):
                    ctx = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
                elif hasattr(ssl, "PROTOCOL_SSLv23"):
                    ctx = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
                else:
                    ctx = ssl.create_default_context()
            except Exception:
                ctx = ssl.create_default_context()

            try:
                ctx.verify_mode = ssl.CERT_NONE
            except Exception:
                pass
            try:
                if hasattr(ctx, "check_hostname"):
                    ctx.check_hostname = False
            except Exception:
                pass
            return ctx
    except Exception:
        pass
    return None


def download_file(url, out_file, label):
    log("    " + label)
    remove_file(out_file)
    headers = {"User-Agent": "Corelan-Win7-Installer/1.0"}
    req = Request(url, headers=headers)
    ctx = get_ssl_context()

    if ctx is not None:
        response = urlopen(req, context=ctx, timeout=120)
    else:
        response = urlopen(req, timeout=120)

    try:
        ensure_dir(os.path.dirname(out_file))
        with open(out_file, "wb") as f:
            while True:
                chunk = response.read(1024 * 64)
                if not chunk:
                    break
                f.write(chunk)
    finally:
        try:
            response.close()
        except Exception:
            pass

    if not os.path.isfile(out_file) or os.path.getsize(out_file) <= 0:
        raise RuntimeError("Download produced empty file: {0}".format(out_file))


def run_checked(cmd, description, cwd=None):
    log("    " + description)
    try:
        rc = subprocess.call(cmd, cwd=cwd)
    except OSError as e:
        raise RuntimeError("{0} failed: {1}".format(description, e))
    if rc != 0:
        raise RuntimeError("{0} failed with exit code {1}".format(description, rc))


def run_capture(cmd, cwd=None):
    try:
        p = subprocess.Popen(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        output = p.communicate()[0]
        if not isinstance(output, text_type):
            try:
                output = output.decode("utf-8")
            except Exception:
                try:
                    output = output.decode("mbcs")
                except Exception:
                    output = output.decode("latin-1", "replace")
        return p.returncode, output
    except OSError as e:
        return 1, text_type(e)


def broadcast_environment_change():
    HWND_BROADCAST = 0xFFFF
    WM_SETTINGCHANGE = 0x001A
    SMTO_ABORTIFHUNG = 0x0002
    try:
        ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 0, u"Environment", SMTO_ABORTIFHUNG, 5000, 0)
    except Exception:
        try:
            ctypes.windll.user32.SendMessageTimeoutA(HWND_BROADCAST, WM_SETTINGCHANGE, 0, b"Environment", SMTO_ABORTIFHUNG, 5000, 0)
        except Exception:
            pass


def check_internet():
    hosts = [
        ("www.python.org", 443),
        ("github.com", 443),
        ("go.microsoft.com", 443),
        ("www.corelan-training.com", 443),
    ]
    for hostname, port in hosts:
        s = socket.create_connection((hostname, port), 20)
        s.close()
        log("    OK   {0}:{1}".format(hostname, port))


def extract_zip(zip_path, dst_folder):
    remove_tree(dst_folder)
    ensure_dir(dst_folder)
    zf = zipfile.ZipFile(zip_path, "r")
    try:
        zf.extractall(dst_folder)
    finally:
        zf.close()


def install_local_symbols():
    zip_path = os.path.join(TEMP_FOLDER, SYMBOLS_ZIP_FILE)
    extract_path = os.path.join(TEMP_FOLDER, "symbols_extract")
    target_root = r"C:\symbols"

    log("    Preparing local symbol store at {0}".format(target_root))

    extract_zip(zip_path, extract_path)
    ensure_dir(target_root)

    symbols_folder = None
    top_level_symbols = os.path.join(extract_path, "symbols")
    if os.path.isdir(top_level_symbols):
        symbols_folder = top_level_symbols
    else:
        for root, dirs, files in os.walk(extract_path):
            for d in dirs:
                if d.lower() == "symbols":
                    symbols_folder = os.path.join(root, d)
                    break
            if symbols_folder:
                break

    if not symbols_folder:
        raise RuntimeError("Could not find 'symbols' folder inside win7symbols.zip")

    nested_symbols = os.path.join(symbols_folder, "symbols")
    if os.path.isdir(nested_symbols):
        symbols_folder = nested_symbols

    for item in os.listdir(symbols_folder):
        src = os.path.join(symbols_folder, item)
        dst = os.path.join(target_root, item)

        if os.path.exists(dst):
            if os.path.isdir(dst):
                shutil.rmtree(dst)
            else:
                os.remove(dst)

        if os.path.isdir(src):
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)

    log("    Symbols installed into {0}".format(target_root))


def install_python27():
    py27_x86 = os.path.join(TEMP_FOLDER, PYTHON2_X86_INSTALLER)
    py27_x64 = os.path.join(TEMP_FOLDER, PYTHON2_X64_INSTALLER)

    run_checked(
        ["msiexec.exe", "/i", py27_x86, "/qn", "ALLUSERS=1", "TARGETDIR={0}".format(PYTHON27_X86_ROOT)],
        "Installing Python 2.7.18 32-bit"
    )
    run_checked(
        ["msiexec.exe", "/i", py27_x64, "/qn", "ALLUSERS=1", "TARGETDIR={0}".format(PYTHON27_X64_ROOT)],
        "Installing Python 2.7.18 64-bit"
    )


def install_python38():
    py32 = os.path.join(TEMP_FOLDER, PYTHON32_INSTALLER)
    py64 = os.path.join(TEMP_FOLDER, PYTHON64_INSTALLER)

    log32 = os.path.join(TEMP_FOLDER, "python38-x86-install.log")
    log64 = os.path.join(TEMP_FOLDER, "python38-x64-install.log")

    args32_primary = [
        py32,
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_pip=1",
        "Include_test=0",
        "Include_launcher=1",
        "TargetDir={0}".format(PYTHON38_X86_ROOT),
        "/log", log32,
    ]
    args64_primary = [
        py64,
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_pip=1",
        "Include_test=0",
        "Include_launcher=1",
        "TargetDir={0}".format(PYTHON38_X64_ROOT),
        "/log", log64,
    ]

    args32_fallback = [
        py32,
        "/quiet",
        "TargetDir={0}".format(PYTHON38_X86_ROOT),
        "Include_pip=1",
        "/log", log32,
    ]
    args64_fallback = [
        py64,
        "/quiet",
        "TargetDir={0}".format(PYTHON38_X64_ROOT),
        "Include_pip=1",
        "/log", log64,
    ]

    try:
        run_checked(args32_primary, "Installing Python 3.8.10 32-bit")
    except Exception:
        log("    Primary install command failed for Python 3.8.10 32-bit, trying fallback")
        run_checked(args32_fallback, "Installing Python 3.8.10 32-bit (fallback)")

    try:
        run_checked(args64_primary, "Installing Python 3.8.10 64-bit")
    except Exception:
        log("    Primary install command failed for Python 3.8.10 64-bit, trying fallback")
        run_checked(args64_fallback, "Installing Python 3.8.10 64-bit (fallback)")


def install_dotnetframework_48():
    dotnet_install = os.path.join(TEMP_FOLDER, DOTNET_FILE)
    run_checked([dotnet_install, "/quiet", "/norestart"], "Installing .Net Framework 4.8 (this may take a while)")


def install_vcredist_2010():
    x86 = os.path.join(TEMP_FOLDER, VCREDIST_X86_FILE)
    x64 = os.path.join(TEMP_FOLDER, VCREDIST_X64_FILE)
    run_checked([x86, "/quiet", "/norestart"], "Installing VC++ 2010 SP1 x86")
    if os.path.exists(x64):
        run_checked([x64, "/quiet", "/norestart"], "Installing VC++ 2010 SP1 x64")


def install_windbg_classic():
    installer = os.path.join(TEMP_FOLDER, WINDBG_INSTALLER)
    run_checked(
        [installer, "/features", "OptionId.WindowsDesktopDebuggers", "/ceip", "off", "/q"],
        "Installing WinDBG Classic via sdksetup.exe"
    )


def detect_windbg_root():
    candidates = [
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "10", "Debuggers"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.1", "Debuggers"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.0", "Debuggers"),
    ]

    for c in candidates:
        x86_exe = os.path.join(c, "x86", "windbg.exe")
        x64_exe = os.path.join(c, "x64", "windbg.exe")
        if os.path.isfile(x86_exe) or os.path.isfile(x64_exe):
            return c
    return None


def create_admin_cmd_shortcut_hack(windbg_root):
    desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.path.join(r"C:\Users\Public", "Desktop")

    shortcut_path = os.path.join(desktop, "Corelan Admin Command Prompt.lnk")

    windir = os.environ.get("WINDIR", r"C:\Windows")
    target = os.path.join(windir, "System32", "cmd.exe")

    startup_folder = os.path.join(windbg_root, "x86")
    if not os.path.isdir(startup_folder):
        startup_folder = os.path.join(windir, "System32")

    vbs = '''
Set oWS = WScript.CreateObject("WScript.Shell")
Set oLink = oWS.CreateShortcut("{shortcut}")
oLink.TargetPath = "{target}"
oLink.WorkingDirectory = "{workdir}"
oLink.IconLocation = "{icon}"
oLink.Description = "Corelan Admin Command Prompt"
oLink.Save
'''.format(
        shortcut=shortcut_path,
        target=target,
        workdir=startup_folder,
        icon=target + ",0"
    )

    vbs_path = os.path.join(desktop, "create_shortcut.vbs")

    try:
        f = open(vbs_path, "wb")
        try:
            f.write(vbs.encode("ascii"))
        finally:
            f.close()
        subprocess.call(["cscript.exe", "//nologo", vbs_path])
    finally:
        try:
            os.remove(vbs_path)
        except Exception:
            pass

    with open(shortcut_path, "rb") as f:
        data = bytearray(f.read())

    if len(data) <= 0x15:
        raise RuntimeError("Shortcut file is too small to patch RunAs flag")

    data[0x15] = data[0x15] | 0x20

    with open(shortcut_path, "wb") as f:
        f.write(data)

    log("    Enabled 'Run as administrator' on shortcut")


def set_system_symbol_path():
    key = winreg.OpenKey(
        winreg.HKEY_LOCAL_MACHINE,
        r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
        0,
        winreg.KEY_SET_VALUE
    )
    try:
        winreg.SetValueEx(key, "_NT_SYMBOL_PATH", 0, winreg.REG_EXPAND_SZ, SYMBOL_PATH)
    finally:
        winreg.CloseKey(key)
    broadcast_environment_change()


def copy_file_checked(src, dst_dir):
    if not os.path.isfile(src):
        raise RuntimeError("Source file does not exist: {0}".format(src))
    if not os.path.isdir(dst_dir):
        raise RuntimeError("Destination folder does not exist: {0}".format(dst_dir))
    shutil.copy2(src, os.path.join(dst_dir, os.path.basename(src)))


def install_mona_and_windbglib(windbg_root):
    x86 = os.path.join(windbg_root, "x86")
    x64 = os.path.join(windbg_root, "x64")
    mona = os.path.join(TEMP_FOLDER, MONA_FILE)
    windbglib = os.path.join(TEMP_FOLDER, WINDBGLIB_FILE)

    copied_any = False

    if os.path.isdir(x86):
        copy_file_checked(mona, x86)
        copy_file_checked(windbglib, x86)
        copied_any = True
    if os.path.isdir(x64):
        copy_file_checked(mona, x64)
        copy_file_checked(windbglib, x64)
        copied_any = True

    if not copied_any:
        raise RuntimeError("No WinDBG x86/x64 folder was found to copy mona.py and windbglib.py into")


def find_file_recursive(root, wanted_name):
    for base, dirs, files in os.walk(root):
        for name in files:
            if name.lower() == wanted_name.lower():
                return os.path.join(base, name)
    return None


def remove_old_pykd_pyd(windbg_root):
    targets = [
        os.path.join(windbg_root, "x86", "winext", "pykd.pyd"),
        os.path.join(windbg_root, "x64", "winext", "pykd.pyd"),
    ]
    for target in targets:
        try:
            if os.path.isfile(target):
                os.remove(target)
                log("    Removed old file: {0}".format(target))
        except Exception:
            pass


def install_pykd_extensions(windbg_root):
    x86_zip = os.path.join(TEMP_FOLDER, PYKD_EXT_X86_ZIP)
    x64_zip = os.path.join(TEMP_FOLDER, PYKD_EXT_X64_ZIP)
    x86_extract = os.path.join(TEMP_FOLDER, "pykd-ext-x86")
    x64_extract = os.path.join(TEMP_FOLDER, "pykd-ext-x64")

    extract_zip(x86_zip, x86_extract)
    extract_zip(x64_zip, x64_extract)

    dll_x86 = find_file_recursive(x86_extract, "pykd.dll")
    dll_x64 = find_file_recursive(x64_extract, "pykd.dll")

    if not dll_x86:
        raise RuntimeError("Unable to locate pykd.dll in x86 PyKD-Ext archive")
    if not dll_x64:
        raise RuntimeError("Unable to locate pykd.dll in x64 PyKD-Ext archive")

    dst_x86 = os.path.join(windbg_root, "x86", "winext")
    dst_x64 = os.path.join(windbg_root, "x64", "winext")

    ensure_dir(dst_x86)
    ensure_dir(dst_x64)

    shutil.copy2(dll_x86, os.path.join(dst_x86, "pykd.dll"))
    shutil.copy2(dll_x64, os.path.join(dst_x64, "pykd.dll"))


def register_dll_silent(dll_path, bitness):
    if not os.path.isfile(dll_path):
        return False
    regsvr = REGSVR32_64 if bitness == "x64" else REGSVR32_32
    if not os.path.isfile(regsvr):
        raise RuntimeError("regsvr32 was not found for {0}".format(bitness))
    run_checked([regsvr, "/s", dll_path], "Registering {0} ({1})".format(os.path.basename(dll_path), bitness))
    return True


def register_msdia_files(windbg_root):
    search_roots = [
        windbg_root,
        os.path.dirname(windbg_root),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "10"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.1"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.0"),
        os.path.join(PROGRAM_FILES_X86, "Common Files", "Microsoft Shared", "VC"),
        os.path.join(PROGRAM_FILES, "Common Files", "Microsoft Shared", "VC"),
        PYTHON27_X86_ROOT,
        PYTHON27_X64_ROOT,
        PYTHON38_X86_ROOT,
        PYTHON38_X64_ROOT,
    ]

    found_any = False

    preferred = [
        ("msdia120.dll", "x86"),
        ("msdia100.dll", "x64"),
        ("msdia140.dll", "x86"),
        ("msdia140.dll", "x64"),
    ]

    seen = set()
    for wanted_name, bitness in preferred:
        for root in search_roots:
            key = (root, wanted_name, bitness)
            if key in seen:
                continue
            seen.add(key)
            if not os.path.isdir(root):
                continue
            hit = find_file_recursive(root, wanted_name)
            if hit:
                if register_dll_silent(hit, bitness):
                    found_any = True
                    break

    if not found_any:
        log("    No msdia*.dll files were found to register. Continuing.")


def run_pip_for_python(python_exe):
    if not os.path.isfile(python_exe):
        log("    Python not found: {0}".format(python_exe))
        return

    run_checked([python_exe, "-m", "pip", "install", "--upgrade", "pip"],
                "Updating pip for {0}".format(python_exe))
    run_checked([python_exe, "-m", "pip", "install", "pykd"],
                "Installing pykd for {0}".format(python_exe))


def upgrade_pip_and_install_pykd():
    targets = [
        os.path.join(PYTHON27_X86_ROOT, "python.exe"),
        os.path.join(PYTHON27_X64_ROOT, "python.exe"),
        os.path.join(PYTHON38_X86_ROOT, "python.exe"),
        os.path.join(PYTHON38_X64_ROOT, "python.exe"),
    ]
    for python_exe in targets:
        run_pip_for_python(python_exe)


def write_text_file(path, content):
    f = open(path, "wb")
    try:
        if not isinstance(content, bytes):
            content = content.encode("ascii")
        f.write(content)
    finally:
        f.close()


def create_wpy2_bat_files(windbg_root):
    x86_dir = os.path.join(windbg_root, "x86")
    x64_dir = os.path.join(windbg_root, "x64")

    content_x86 = (
        "@echo off\r\n"
        "REM ==========================================\r\n"
        "REM Run WinDBG with optional arguments\r\n"
        "REM Corelan Stack / Heap Training\r\n"
        "REM www.corelan-training.com\r\n"
        "REM ==========================================\r\n"
        "\r\n"
        "set PATH=C:\\Python27;%PATH%\r\n"
        "set PYTHONHOME=C:\\Python27\r\n"
        "set PYTHONPATH=C:\\Python27\\Lib\r\n"
        "\r\n"
        "REM Define base command (adjust path to wew file as needed)\r\n"
        "set \"WINDBG_CMD=windbg.exe -hd -c '!load pykd.pyd;as mona !py --global mona.py'\"\r\n"
        "\r\n"
        "%WINDBG_CMD% %*\r\n"
    )

    content_x64 = (
        "@echo off\r\n"
        "REM ==========================================\r\n"
        "REM Run WinDBG with optional arguments\r\n"
        "REM Corelan Stack / Heap Training\r\n"
        "REM www.corelan-training.com\r\n"
        "REM ==========================================\r\n"
        "\r\n"
        "set PATH=C:\\Python27-64;%PATH%\r\n"
        "set PYTHONHOME=C:\\Python27-64\r\n"
        "set PYTHONPATH=C:\\Python27-64\\Lib\r\n"
        "\r\n"
        "REM Define base command (adjust path to wew file as needed)\r\n"
        "set \"WINDBG_CMD=windbg.exe -hd -c '!load pykd.pyd;as mona !py --global mona.py'\"\r\n"
        "\r\n"
        "%WINDBG_CMD% %*\r\n"
    )

    if os.path.isdir(x86_dir):
        write_text_file(os.path.join(x86_dir, "wpy2.bat"), content_x86)
    if os.path.isdir(x64_dir):
        write_text_file(os.path.join(x64_dir, "wpy2.bat"), content_x64)


def create_wpy3_bat_files(windbg_root):
    x86_dir = os.path.join(windbg_root, "x86")
    x64_dir = os.path.join(windbg_root, "x64")

    content_x86 = (
        "@echo off\r\n"
        "REM ==========================================\r\n"
        "REM Run WinDBG with optional arguments\r\n"
        "REM Corelan Stack / Heap Training\r\n"
        "REM www.corelan-training.com\r\n"
        "REM ==========================================\r\n"
        "\r\n"
        "set PATH=%LOCALAPPDATA%\\Python\\Python38-32;%PATH%\r\n"
        "set PYTHONHOME=%LOCALAPPDATA%\\Python\\Python38-32\r\n"
        "set PYTHONPATH=%LOCALAPPDATA%\\Python\\Python38-32\\Lib\r\n"
        "\r\n"
        "REM Define base command (adjust path to wew file as needed)\r\n"
        "set \"WINDBG_CMD=windbg.exe -hd -c '!load pykd;as mona !py --global mona.py' \"\r\n"
        "\r\n"
        "%WINDBG_CMD% %*\r\n"
    )

    content_x64 = (
        "@echo off\r\n"
        "REM ==========================================\r\n"
        "REM Run WinDBG with optional arguments\r\n"
        "REM Corelan Stack / Heap Training\r\n"
        "REM www.corelan-training.com\r\n"
        "REM ==========================================\r\n"
        "\r\n"
        "set PATH=%LOCALAPPDATA%\\Python\\Python38;%PATH%\r\n"
        "set PYTHONHOME=%LOCALAPPDATA%\\Python\\Python38\r\n"
        "set PYTHONPATH=%LOCALAPPDATA%\\Python\\Python38\\Lib\r\n"
        "\r\n"
        "REM Define base command (adjust path to wew file as needed)\r\n"
        "set \"WINDBG_CMD=windbg.exe -hd -c '!load pykd;as mona !py --global mona.py\"\r\n"
        "\r\n"
        "%WINDBG_CMD% %*\r\n"
    )

    if os.path.isdir(x86_dir):
        write_text_file(os.path.join(x86_dir, "wpy3.bat"), content_x86)
    if os.path.isdir(x64_dir):
        write_text_file(os.path.join(x64_dir, "wpy3.bat"), content_x64)



def download_everything():
    download_file(PYTHON2_X86_URL, os.path.join(TEMP_FOLDER, PYTHON2_X86_INSTALLER), "1. Python 2.7.18 32-bit")
    download_file(PYTHON2_X64_URL, os.path.join(TEMP_FOLDER, PYTHON2_X64_INSTALLER), "2. Python 2.7.18 64-bit")
    download_file(PYTHON32_URL, os.path.join(TEMP_FOLDER, PYTHON32_INSTALLER), "3. Python 3.8.10 32-bit")
    download_file(PYTHON64_URL, os.path.join(TEMP_FOLDER, PYTHON64_INSTALLER), "4. Python 3.8.10 64-bit")
    download_file(WINDBG_URL, os.path.join(TEMP_FOLDER, WINDBG_INSTALLER), "5. WinDBG Classic sdksetup.exe")
    download_file(MONA_URL, os.path.join(TEMP_FOLDER, MONA_FILE), "6. mona.py")
    download_file(WINDBGLIB_URL, os.path.join(TEMP_FOLDER, WINDBGLIB_FILE), "7. windbglib.py")
    download_file(PYKD_EXT_X86_URL, os.path.join(TEMP_FOLDER, PYKD_EXT_X86_ZIP), "8. PyKD-Ext x86")
    download_file(PYKD_EXT_X64_URL, os.path.join(TEMP_FOLDER, PYKD_EXT_X64_ZIP), "9. PyKD-Ext x64")
    download_file(VCREDIST_X86_URL, os.path.join(TEMP_FOLDER, VCREDIST_X86_FILE), "10. VC++ 2010 SP1 x86")
    download_file(VCREDIST_X64_URL, os.path.join(TEMP_FOLDER, VCREDIST_X64_FILE), "11. VC++ 2010 SP1 x64")
    download_file(DOTNET_URL, os.path.join(TEMP_FOLDER, DOTNET_FILE), "12. .Net Framework 4.8")
    download_file(SYMBOLS_ZIP_URL, os.path.join(TEMP_FOLDER, SYMBOLS_ZIP_FILE), "13. Win7 local symbols")


def cleanup():
    log("[+] Removing temporary folder again")
    remove_tree(TEMP_FOLDER)


def main():
    if os.name != "nt":
        abort("This script must be run on Windows.")

    if not is_admin():
        abort("This script must be run as Administrator.")

    log("*** -->> Make sure you have an active internet connection before proceeding! <<-- ***")
    confirm_continue("Ready to continue?")

    safe_step("Checking internet connectivity", check_internet)

    log("\n[+] Creating temp folder {0}".format(TEMP_FOLDER))
    ensure_dir(TEMP_FOLDER)
    ensure_dir(ENGINE_EXT_32)
    ensure_dir(ENGINE_EXT_64)

    safe_step("Downloading installers and support files", download_everything)
    safe_step("Installing local Win7 symbols into C:\\symbols", install_local_symbols)
    safe_step("Installing .Net Framework", install_dotnetframework_48)
    safe_step("Installing Python 2.7.18 x86/x64", install_python27)
    safe_step("Installing Python 3.8.10 x86/x64", install_python38)
    safe_step("Updating pip and installing pykd", upgrade_pip_and_install_pykd)

    windbg_root = detect_windbg_root()
    if windbg_root:
        log("[+] WinDBG debugger folder found at: {0}".format(windbg_root))
    else:
        safe_step("Installing WinDBG Classic", install_windbg_classic)
        windbg_root = detect_windbg_root()

    safe_step("Creating system environment variable _NT_SYMBOL_PATH", set_system_symbol_path)

    if windbg_root:
        safe_step("Copying mona.py and windbglib.py into WinDBG x86/x64 folders", install_mona_and_windbglib, windbg_root)
        safe_step("Removing old pykd.pyd files from WinDBG winext folders", remove_old_pykd_pyd, windbg_root)
        safe_step("Installing pykd.dll into WinDBG x86/x64 winext folders", install_pykd_extensions, windbg_root)
        safe_step("Installing VC++ 2010 SP1 Redistributables", install_vcredist_2010)
        safe_step("Registering msdia files", register_msdia_files, windbg_root)
        safe_step("Creating wpy2.bat launchers", create_wpy2_bat_files, windbg_root)
        safe_step("Creating wpy3.bat launchers", create_wpy3_bat_files, windbg_root)
    else:
        log("[!] WinDBG debugger folder was not found. Skipping WinDBG-dependent steps.")
        safe_step("Installing VC++ 2010 SP1 Redistributables", install_vcredist_2010)

    safe_step("Creating elevated command prompt shortcut on desktop", create_admin_cmd_shortcut_hack)
    safe_step("Cleaning up temporary folder", cleanup)

    log("")
    log("[+] Script completed")
    log("[+] Review the log above for any failed steps")


if __name__ == "__main__":
    main()