This repository contains a few Powershell and Python scripts:

1. `CorelanVMInstall.ps1`: Will help you set up a Windows 11 VM, in preparation for [Corelan Windows Exploit Development Training](https://www.corelan-training.com)
2. `CorelanPyKDInstall.ps1`: This script sets up your machine to run support Python3 / PyKD / PyKD-Ext in WinDBG(x), for both 32bit and 64bit 
3. `CorelanWin7VMinstall.py`: This python script will install the necessary components to run PyKD/PyKD-Ext scripts on a Windows 7 VM.

Prior to running one of the Powershell script(s) on Windows 11,  please update your Windows VM first.
Some parts of the script rely on `winget`, so make sure that is present on your system already.

Make sure to run scripts from an administrator prompt!!


# CorelanVMInstall.ps1

Usage
-----

1. Download `CorelanVMInstall.ps1` to your Windows 11/10 VM
2. Open an administrator command prompt and go to the folder that contains the `CorelanVMInstall.ps1` file
3. Verify/confirm that you have a working internet connection
4. run `powershell ./CorelanVMInstall.ps1`
5. If all goes well, the script will:
  - download installers for Python 2.7.18, WinDBG, WinDBGX, PyKD, mona.py, windbglib.py and Visual Studio 2017 Desktop Express
  - install winget if needed
  - install the required prerequisites and applications
  - install WinDBGX, Visual Studio Code and 7Zip via winget
  - set up the PATH environment variable
  - create a system environment variable `_NT_SYMBOL_PATH`
  - create an administrator Command Prompt shortcut on the Desktop


# CorelanPyKDInstall.ps1

This second script will install Python3.9 and a PyKD + pykd-ext version that is compatible with Python3 and Python2
Please keep in mind that this script will break existing legacy mona.py installations that are based on PyKD 0.2.0.x and Python2
Do not use this script unless you know what you're doing ;-)

It allows you to run `.load pykd` in WinDBG(x), loading the `pykd-ext` extension (`pykd.dll`).
This allows you to run pykd/python scripts via `!py`.


# CorelanWin7VMInstall.py

This script kind of simulates what `CorelanPyKDInstall.ps1` does, but in Python and for Windows 7 specifically.




# FAQ 

## Help, the powershell script refuses to run.  For example:

```
./CorelanVMInstall.ps1 : File CorelanVMInstall.ps1 cannot be loaded because running scripts is disabled on
 this system. For more information, see about_Execution_Policies at https:/go.microsoft.com/fwlink/?LinkID=135170.
 ```

Solution:
* Open a PowerShell window (as administrator)
* Run `Set-ExecutionPolicy RemoteSigned` and press "Y" when prompted
* Try running the powershell script again.

If that doesn't work, try `Set-ExecutionPolicy Unrestricted` instead


