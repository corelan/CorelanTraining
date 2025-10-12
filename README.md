This repository contains a Powershell script to assist with the setup of Virtual Machines, in preparation for Corelan Windows Exploit Development Training.

For info on classes, see https://www.corelan-training.com/


Usage
-----

1. Download `CorelanVMInstall.ps1` to your Windows 11/10 VM
2. Open an administrator command prompt and go to the folder that contains the `CorelanVMInstall.ps1` file
3. Verify/confirm that you have a working internet connection
4. run `powershell ./CorelanVMInstall.ps1`
5. If all goes well, the script will:
  - download installers for Python 2.7.17, WinDBG, PyKD, mona.py, windbglib.py and Visual Studio 2017 Desktop Express
  - install the required prerequisites and applications
  - set up the PATH environment variable
  - create a system environment variable `_NT_SYMBOL_PATH`


FAQ 
----

## File CorelanVMInstall.ps1 cannot be loaded because running scripts is disabled on this system

My freshly installed Windows 11/10 doesn't allow me to run your powershell script. It produces the following error message:

```
./CorelanVMInstall.ps1 : File CorelanVMInstall.ps1 cannot be loaded because running scripts is disabled on
 this system. For more information, see about_Execution_Policies at https:/go.microsoft.com/fwlink/?LinkID=135170.
 ```

Solution:
* Open a PowerShell window (as administrator)
* Run `Set-ExecutionPolicy RemoteSigned` and press "Y" when prompted
* Try running the powershell script again.



##  !peb produces 'error 3 InitTypeRead' on Windows 10 1903/1909

On Windows 10 (1903/1909), WinDBG throws an error when running `!peb` or when trying to run mona.py commands:

```
0:000> !peb
PEB at xxxxxxxx
error 3 InitTypeRead
``` 

It looks like MS may have removed(?) type information from the latest symbol files associated with ntdll.dll.
As a workaround, you can try the following procedure:

1. Open folder `c:\symbols\wntdll.pdb` and delete all subfolders
2. Open an administrator command prompt
3. Run `C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\windbg.exe -o c:\windows\system32\calc.exe`
4. In WinDBG, run `!peb` and confirm that it is still broken
5. Close WinDBG and open folder `c:\symbols\wntdll.pdb`.  There should be only one subfolder, for instance `D85FCE08D56038E2C69B69F29E11B5EE1`(the actual name could be different). Open the folder and remove wntdll.pdb from that folder. We'll call this the `original` folder.  Leave this original `D85FCE08D56038E2C69B69F29E11B5EE1` folder open.
6. Download wntdllsymbolfix.zip file from this repository
7. Extract the zipfile directly into the `c:\symbols\wntdll.pdb` folder. You should get an additional folder and a file:
- Folder: `6BFA8EAE64E07F11AD6B27F575C7BDC21`
- File: `ChkMatch.exe`
8. From inside the new `6BFA8EAE64E07F11AD6B27F575C7BDC21` folder, copy wntdll.pdb and paste it into the `original` folder (the one where you just removed wntdll.pdb)
3. Open an administrator command prompt and go to the `c:\symbols\wntdll.pdb` folder
4. Run the following command to forcibly match ntdll.dll with the older symbol file (replace <foldername> with the name of the `original` folder):

```
ChkMatch.exe -m c:\Windows\SysWOW64\ntdll.dll c:\symbols\wntdll.pdb\<foldername>\wntdll.pdb
```



Example output:

```
C:\symbols\wntdll.pdb>ChkMatch.exe -m c:\Windows\SysWOW64\ntdll.dll c:\symbols\wntdll.pdb\D85FCE08D56038E2C69B69F29E11B5EE1\wntdll.pdb
ChkMatch - version 1.0
Copyright (C) 2004 Oleg Starodumov
http://www.debuginfo.com/


Executable: c:\Windows\SysWOW64\ntdll.dll
Debug info file: c:\symbols\wntdll.pdb\D85FCE08D56038E2C69B69F29E11B5EE1\wntdll.pdb

Executable:
TimeDateStamp: a4208572
Debug info: 2 ( CodeView )
TimeStamp: a4208572  Characteristics: 0  MajorVer: 0  MinorVer: 0
Size: 35  RVA: 000255e8  FileOffset: 000249e8
CodeView format: RSDS
Signature: {d85fce08-d560-38e2-c69b-69f29e11b5ee}  Age: 1
PdbFile: wntdll.pdb
Debug info: 13 ( Unknown )
TimeStamp: a4208572  Characteristics: 0  MajorVer: 0  MinorVer: 0
Size: 1252  RVA: 0002560c  FileOffset: 00024a0c
Debug info: 16 ( Unknown )
TimeStamp: a4208572  Characteristics: 0  MajorVer: 0  MinorVer: 0
Size: 36  RVA: 00025af0  FileOffset: 00024ef0

Debug information file:
Format: PDB 7.00
Signature: {6bfa8eae-64e0-7f11-ad6b-27f575c7bdc2}  Age: 2

Writing to the debug information file...
Result: Success.

```

5. Open WinDBG again (`C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\windbg.exe -o c:\windows\system32\calc.exe`), run `!peb` and verify that the issue has been resolved


Enjoy!

