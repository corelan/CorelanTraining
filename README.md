This repository contains a few PowerShell and Python scripts:

1. [CorelanWin11VMInstall.ps1](CorelanWin11VMInstall.ps1): Helps set up and configure a Windows 11 VM in preparation for [Corelan Windows Exploit Development Training](https://www.corelan-training.com).
2. [CorelanPyKDInstall.ps1](CorelanPyKDInstall.ps1): Installs and configures the necessary components to run Python 3 / PyKD / PyKD-Ext in WinDBG(x), for both 32-bit and 64-bit, on Windows 10/11.
3. [CorelanWin7VMInstall.py](CorelanWin7VMInstall.py): Installs the necessary components to run PyKD/PyKD-Ext scripts on a Windows 7 VM. This script requires a working instance of Python 2.7.18.

## Table of Contents

- [Windows 11 VM Setup](#windows-11-vm-setup)
- [PyKD for Windows 10/11](#pykd-for-windows-1011)
- [PyKD for Windows 7](#pykd-for-windows-7)
- [FAQ](#faq)

Prior to running one of the PowerShell scripts on Windows 11, please update your Windows VM first.
Some parts of the scripts rely on `winget`, so make sure that is present on your system already.

Make sure to run scripts from an administrator prompt.

## Windows 11 VM Setup

1. Download [CorelanWin11VMInstall.ps1](CorelanWin11VMInstall.ps1) to your Windows 11/10 VM.
2. Open an administrator command prompt and go to the folder that contains `CorelanWin11VMInstall.ps1`.
3. Verify that you have a working internet connection.
4. Run `powershell ./CorelanWin11VMInstall.ps1`.
5. If all goes well, the script will:
   - Download installers for Python 2.7.18, Python 3.9.13, WinDBG, WinDBGX, PyKD, `mona.py`, `windbglib.py`, and Visual Studio 2017 Desktop Express.
   - Install `winget` if needed.
   - Install the required prerequisites, libraries and applications.
   - Install WinDBGX, Visual Studio Code, and 7-Zip via `winget`.
   - Set up the `PATH` environment variable.
   - Create a system environment variable named `_NT_SYMBOL_PATH`.
   - Create an administrator Command Prompt shortcut on the desktop.

## PyKD for Windows 10/11

This script installs Python 3.9 and a PyKD + PyKD-Ext version that is compatible with both Python 3 and Python 2.
It also installs the `keystone-engine` library for Python 3.9 specifically.

Please keep in mind that this script will break existing legacy `mona.py` installations based on PyKD 0.2.0.x and Python 2.

Do not use this script unless you know what you're doing.

It allows you to run `!load pykd` in WinDBG(x), loading the `pykd-ext` extension (`pykd.dll`).
This allows you to run PyKD/Python scripts via `!py`.

## PyKD for Windows 7

This script simulates what [CorelanPyKDInstall.ps1](CorelanPyKDInstall.ps1) does, but it is written in Python and targets Windows 7 SP1 and later specifically.
Install Python 2.7.18 yourself first, then run [CorelanWin7VMInstall.py](CorelanWin7VMInstall.py) from an administrator command prompt.

## FAQ

### Help, the PowerShell script refuses to run

For example:

```powershell
./CorelanWin11VMInstall.ps1 : File CorelanWin11VMInstall.ps1 cannot be loaded because running scripts is disabled on
 this system. For more information, see about_Execution_Policies at https:/go.microsoft.com/fwlink/?LinkID=135170.
```

Solution:

- Open a PowerShell window as administrator.
- Run `Set-ExecutionPolicy RemoteSigned` and press `Y` when prompted.
- Try running the PowerShell script again.

If that doesn't work, try `Set-ExecutionPolicy Unrestricted` instead.
