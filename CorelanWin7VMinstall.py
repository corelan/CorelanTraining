# -*- coding: utf-8 -*-
from __future__ import print_function

"""
corelanWin7VMinstall.py

(c) Corelan Consulting bv - 2026
www.corelan-consulting.com
www.corelan-training.com
www.corelan-certified.com
www.corelan.be

Python 2 / Python 3 compatible installer helper for older Windows 7 SP1 VMs.

IMPORTANT:
- Run this script as Administrator.
- This script uses only the Python standard library for downloads/extraction.
- The script will attempt to install Python 3.8.10 directly, because that is the right fit for legacy Windows 7 SP1 environments.
- WinDBG Classic is installed through sdksetup.exe.
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

PYTHON2_URL = "https://www.python.org/ftp/python/2.7.18/python-2.7.18.msi"
PYTHON32_URL = "https://www.python.org/ftp/python/3.8.10/python-3.8.10.exe"
PYTHON64_URL = "https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe"
WINDBG_URL = "https://go.microsoft.com/fwlink/p/?LinkId=323507"
MONA_URL = "https://github.com/corelan/mona/raw/master/mona.py"
WINDBGLIB_URL = "https://github.com/corelan/windbglib/raw/master/windbglib.py"
PYKD_EXT_X86_URL = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x86.zip"
PYKD_EXT_X64_URL = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x64.zip"
PYKD_PYD_URL = "https://github.com/corelan/windbglib/raw/master/pykd/pykd.zip"
VCREDIST_X86_URL = "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe"
VCREDIST_X64_URL = "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe"
DOTNET_URL = "https://go.microsoft.com/fwlink/?linkid=2088631"
SYMBOLS_ZIP_URL = "https://www.corelan-training.com/downloads/win7symbols.zip"

PYTHON2_INSTALLER = "python-2.7.18.msi"
PYTHON32_INSTALLER = "python-3.8.10.exe"
PYTHON64_INSTALLER = "python-3.8.10-amd64.exe"
WINDBG_INSTALLER = "sdksetup.exe"
MONA_FILE = "mona.py"
WINDBGLIB_FILE = "windbglib.py"
PYKD_EXT_X86_ZIP = "pykd-ext-x86.zip"
PYKD_EXT_X64_ZIP = "pykd-ext-x64.zip"
PYKD_PYD_ZIP = "pykd.zip"
VCREDIST_X86_FILE = "vcredist_x86.exe"
VCREDIST_X64_FILE = "vcredist_x64.exe"
DOTNET_FILE = "NPD48-x86-x64-AllOS-ENU.exe"
SYMBOLS_ZIP_FILE = "win7symbols.zip"


ENGINE_EXT_64 = os.path.join(os.environ.get("LOCALAPPDATA", r"C:\Users\Default\AppData\Local"), "DBG", "EngineExtensions")
ENGINE_EXT_32 = os.path.join(os.environ.get("LOCALAPPDATA", r"C:\Users\Default\AppData\Local"), "DBG", "EngineExtensions32")

PYTHON32_ROOT = os.path.join(os.environ.get("LOCALAPPDATA", r"C:\Users\Default\AppData\Local"), "Programs", "Python", "Python38-32")
PYTHON64_ROOT = os.path.join(os.environ.get("LOCALAPPDATA", r"C:\Users\Default\AppData\Local"), "Programs", "Python", "Python38")

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
            if hasattr(ssl, "PROTOCOL_TLSv1_2"):
                ctx = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
            else:
                ctx = ssl.create_default_context()
            try:
                ctx.verify_mode = ssl.CERT_NONE
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
        response = urlopen(req, context=ctx, timeout=60)
    else:
        response = urlopen(req, timeout=60)

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
        ("pypi.org", 443),
        ("github.com", 443),
        ("go.microsoft.com", 443),
    ]
    for host, port in hosts:
        s = socket.create_connection((host, port), 20)
        s.close()
        log("    OK   {0}:{1}".format(host, port))

def install_local_symbols():
    """
    Downloaded win7symbols.zip contains a 'symbols' folder.
    Extract its contents into C:\symbols
    """

    zip_path = os.path.join(TEMP_FOLDER, SYMBOLS_ZIP_FILE)
    extract_path = os.path.join(TEMP_FOLDER, "symbols_extract")
    target_root = r"C:\symbols"

    log("    Preparing local symbol store at {0}".format(target_root))

    # extract zip
    extract_zip(zip_path, extract_path)

    # locate "symbols" folder inside extracted content
    symbols_folder = None
    for root, dirs, files in os.walk(extract_path):
        for d in dirs:
            if d.lower() == "symbols":
                symbols_folder = os.path.join(root, d)
                break
        if symbols_folder:
            break

    if not symbols_folder:
        raise RuntimeError("Could not find 'symbols' folder inside win7symbols.zip")

    # create C:\symbols
    ensure_dir(target_root)

    # copy contents (not the parent folder itself)
    for item in os.listdir(symbols_folder):
        src = os.path.join(symbols_folder, item)
        dst = os.path.join(target_root, item)

        if os.path.isdir(src):
            if os.path.exists(dst):
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)

    log("    Symbols installed into {0}".format(target_root))

def install_python38():
    py32 = os.path.join(TEMP_FOLDER, PYTHON32_INSTALLER)
    py64 = os.path.join(TEMP_FOLDER, PYTHON64_INSTALLER)

    args32 = [
        py32,
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_pip=1",
        "Include_test=0",
        "Include_launcher=1",
        "SimpleInstall=1",
        "TargetDir={0}".format(PYTHON32_ROOT),
    ]
    args64 = [
        py64,
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_pip=1",
        "Include_test=0",
        "Include_launcher=1",
        "SimpleInstall=1",
        "TargetDir={0}".format(PYTHON64_ROOT),
    ]

    run_checked(args32, "Installing Python 3.8.10 32-bit")
    run_checked(args64, "Installing Python 3.8.10 64-bit")


def install_python27():
    py27 = os.path.join(TEMP_FOLDER, PYTHON2_INSTALLER)
    run_checked(
        ["msiexec.exe", "/i", py27, "/qn", "ALLUSERS=0"],
        "Installing Python 2.7.18 32-bit"
    )


def find_py_launcher():
    candidates = [
        os.path.join(WINDIR, "py.exe"),
        os.path.join(PYTHON64_ROOT, "py.exe"),
        os.path.join(PYTHON32_ROOT, "py.exe"),
        "py",
    ]
    for c in candidates:
        if c == "py":
            rc, out = run_capture([c, "-0p"])
            if rc == 0:
                return c
        else:
            if os.path.isfile(c):
                return c
    return None


def parse_py_tags(py_launcher):
    rc, out = run_capture([py_launcher, "-0p"])
    if rc != 0:
        raise RuntimeError("Unable to enumerate Python versions via py launcher")

    tags = []
    for line in out.splitlines():
        line = line.strip()
        if not line:
            continue
        if not line.startswith("-"):
            continue
        parts = line.split()
        tag = parts[0]
        if tag not in tags:
            tags.append(tag)
    return tags


def upgrade_pip_and_install_pykd():
    py_launcher = find_py_launcher()
    if not py_launcher:
        raise RuntimeError("py launcher was not found")

    tags = parse_py_tags(py_launcher)
    if not tags:
        raise RuntimeError("No Python versions were found via py launcher")

    for tag in tags:
        run_checked([py_launcher, tag, "-m", "pip", "install", "--upgrade", "pip"],
                    "Updating pip for {0}".format(tag))
        run_checked([py_launcher, tag, "-m", "pip", "install", "pykd"],
                    "Installing pykd for {0}".format(tag))


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
    run_checked([installer, "/features", "OptionId.WindowsDesktopDebuggers", "/ceip", "off", "/q"],
                "Installing WinDBG Classic via sdksetup.exe")


def detect_windbg_root():
    candidates = [
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.1", "Debuggers"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.0", "Debuggers"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "10", "Debuggers"),
    ]
    for c in candidates:
        if os.path.isdir(c):
            x86 = os.path.join(c, "x86")
            x64 = os.path.join(c, "x64")
            if os.path.isdir(x86) or os.path.isdir(x64):
                return c
    return None


def create_admin_cmd_shortcut_hack():
    desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.path.join(r"C:\Users\Public", "Desktop")

    shortcut_path = os.path.join(desktop, "Corelan Admin Command Prompt.lnk")

    windir = os.environ.get("WINDIR", r"C:\Windows")
    target = os.path.join(windir, "System32", "cmd.exe")

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
        workdir=os.path.join(windir, "System32"),
        icon=target + ",0"
    )

    vbs_path = os.path.join(desktop, "create_shortcut.vbs")

    try:
        f = open(vbs_path, "w")
        f.write(vbs)
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
    log("    Shortcut created at: {0}".format(shortcut_path))


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


def extract_zip(zip_path, dst_folder):
    remove_tree(dst_folder)
    ensure_dir(dst_folder)
    zf = zipfile.ZipFile(zip_path, "r")
    try:
        zf.extractall(dst_folder)
    finally:
        zf.close()


def find_file_recursive(root, wanted_name):
    for base, dirs, files in os.walk(root):
        for name in files:
            if name.lower() == wanted_name.lower():
                return os.path.join(base, name)
    return None


def find_all_files_recursive(root, wanted_name):
    hits = []
    for base, dirs, files in os.walk(root):
        for name in files:
            if name.lower() == wanted_name.lower():
                hits.append(os.path.join(base, name))
    return hits


def choose_pykd_pyd_for_arch(candidates, arch_tag):
    arch_tag = arch_tag.lower()
    for candidate in candidates:
        low = candidate.lower()
        if arch_tag in low:
            return candidate

    if arch_tag == "x86":
        for candidate in candidates:
            low = candidate.lower()
            if "32" in low:
                return candidate

    if arch_tag == "x64":
        for candidate in candidates:
            low = candidate.lower()
            if "64" in low or "amd64" in low:
                return candidate

    if candidates:
        return candidates[0]

    return None


def install_pykd_extensions():
    ensure_dir(ENGINE_EXT_32)
    ensure_dir(ENGINE_EXT_64)

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

    shutil.copy2(dll_x86, os.path.join(ENGINE_EXT_32, "pykd.dll"))
    shutil.copy2(dll_x64, os.path.join(ENGINE_EXT_64, "pykd.dll"))


def install_pykd_pyd_winext(windbg_root):
    """
    Extract pykd.zip, install VC++ 2010 x86, then copy pykd.pyd into WinDBG winext folders.
    """

    pykd_zip = os.path.join(TEMP_FOLDER, PYKD_PYD_ZIP)
    pykd_extract = os.path.join(TEMP_FOLDER, "pykd-pyd")

    extract_zip(pykd_zip, pykd_extract)

    pyd_candidates = find_all_files_recursive(pykd_extract, "pykd.pyd")
    if not pyd_candidates:
        raise RuntimeError("Unable to locate pykd.pyd in pykd.zip")

    vcredist_x86 = os.path.join(TEMP_FOLDER, VCREDIST_X86_FILE)
    if not os.path.isfile(vcredist_x86):
        raise RuntimeError("VC++ 2010 x86 installer was not found: {0}".format(vcredist_x86))

    run_checked([vcredist_x86, "/quiet", "/norestart"],
                "Installing VC++ 2010 SP1 x86 for pykd.pyd")

    copied_any = False

    x86_winext = os.path.join(windbg_root, "x86", "winext")
    x64_winext = os.path.join(windbg_root, "x64", "winext")

    pyd_x86 = choose_pykd_pyd_for_arch(pyd_candidates, "x86")
    pyd_x64 = choose_pykd_pyd_for_arch(pyd_candidates, "x64")

    if os.path.isdir(x86_winext) and pyd_x86:
        shutil.copy2(pyd_x86, os.path.join(x86_winext, "pykd.pyd"))
        copied_any = True

    if os.path.isdir(x64_winext) and pyd_x64:
        shutil.copy2(pyd_x64, os.path.join(x64_winext, "pykd.pyd"))
        copied_any = True

    if not copied_any:
        raise RuntimeError("No WinDBG winext folders were found to copy pykd.pyd into")


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
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.1"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "8.0"),
        os.path.join(PROGRAM_FILES_X86, "Windows Kits", "10"),
        os.path.join(PROGRAM_FILES_X86, "Common Files", "Microsoft Shared", "VC"),
        os.path.join(PROGRAM_FILES, "Common Files", "Microsoft Shared", "VC"),
        PYTHON32_ROOT,
        PYTHON64_ROOT,
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


def download_everything():
    download_file(PYTHON2_URL, os.path.join(TEMP_FOLDER, PYTHON2_INSTALLER), "1. Python 2.7.18 32-bit")
    download_file(PYTHON32_URL, os.path.join(TEMP_FOLDER, PYTHON32_INSTALLER), "2. Python 3.8.10 32-bit")
    download_file(PYTHON64_URL, os.path.join(TEMP_FOLDER, PYTHON64_INSTALLER), "3. Python 3.8.10 64-bit")
    download_file(WINDBG_URL, os.path.join(TEMP_FOLDER, WINDBG_INSTALLER), "4. WinDBG Classic sdksetup.exe")
    download_file(MONA_URL, os.path.join(TEMP_FOLDER, MONA_FILE), "5. mona.py")
    download_file(WINDBGLIB_URL, os.path.join(TEMP_FOLDER, WINDBGLIB_FILE), "6. windbglib.py")
    download_file(PYKD_EXT_X86_URL, os.path.join(TEMP_FOLDER, PYKD_EXT_X86_ZIP), "7. PyKD-Ext x86")
    download_file(PYKD_EXT_X64_URL, os.path.join(TEMP_FOLDER, PYKD_EXT_X64_ZIP), "8. PyKD-Ext x64")
    download_file(PYKD_PYD_URL, os.path.join(TEMP_FOLDER, PYKD_PYD_ZIP), "9. pykd.zip (pykd.pyd)")
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
    safe_step("Installing Python 2.7.18", install_python27)
    safe_step("Installing Python 3.8.10", install_python38)
    safe_step("Updating pip and installing pykd", upgrade_pip_and_install_pykd)
    safe_step("Installing WinDBG Classic", install_windbg_classic)

    safe_step("Creating system environment variable _NT_SYMBOL_PATH", set_system_symbol_path)


    windbg_root = safe_step("Detecting WinDBG installation folder", detect_windbg_root)

    if windbg_root and windbg_root is not True:
        log("[+] WinDBG debugger folder found at: {0}".format(windbg_root))
        safe_step("Copying mona.py and windbglib.py into WinDBG x86/x64 folders", install_mona_and_windbglib, windbg_root)
        safe_step("Extracting pykd.pyd and copying it into WinDBG winext folders", install_pykd_pyd_winext, windbg_root)
        safe_step("Obtaining and copying pykd.dll into DBG EngineExtensions folders", install_pykd_extensions)
        safe_step("Installing VC++ 2010 SP1 Redistributables", install_vcredist_2010)
        safe_step("Registering msdia files", register_msdia_files, windbg_root)
    else:
        log("[!] WinDBG debugger folder was not found. Skipping WinDBG-dependent steps.")
        safe_step("Installing VC++ 2010 SP1 Redistributables", install_vcredist_2010)
        safe_step("Obtaining and copying pykd.dll into DBG EngineExtensions folders", install_pykd_extensions)

    safe_step("Creating elevated command prompt shortcut on desktop", create_admin_cmd_shortcut_hack)
    safe_step("Cleaning up temporary folder", cleanup)

    log("")
    log("[+] Script completed")
    log("[+] Review the log above for any failed steps")
    log("[+] You may want to reboot the VM before launching WinDBG")


if __name__ == "__main__":
    main()