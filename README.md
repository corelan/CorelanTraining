This repository contains script(s) to assist with the setup of Virtual Machines, in preparation for Corelan Windows Exploit Development Training.

For info on classes, see https://www.corelan-training.com/


Usage
-----

1. Download `CorelanVMInstall.ps1` to your Windows 10 VM
2. Open an administrator command prompt and go to the folder that contains the `CorelanVMInstall.ps1` file
3. Verify/confirm that you have a working internet connection
4. run `powershell ./CorelanVMInstall.ps1`
5. If all goes well, the script will:
  - download Python 2.7.17, WinDBG, PyKD, mona.py and windbglib.py
  - install the required prerequisites
  - set up the PATH environment variable
  - create a system environment variable `_NT_SYMBOL_PATH`


FAQ 
----

1. Help! My freshly installed Windows 10 doesn't allow me to run your powershell script. It produces the following error message:

```
./CorelanVMInstall.ps1 : File CorelanVMInstall.ps1 cannot be loaded because running scripts is disabled on
 this system. For more information, see about_Execution_Policies at https:/go.microsoft.com/fwlink/?LinkID=135170.
 ```

Solution:
* Open a PowerShell window (as administrator)
* Run `Set-ExecutionPolicy RemoteSigned` and press "Y" when prompted
* Try running the powershell script again.


Enjoy!

