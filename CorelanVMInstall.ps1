#Requires -RunAsAdministrator
#
# (c) Corelan Consulting bv - 2020
# www.corelan-consulting.com
# www.corelan-training.com
#
# PowerShell script to prepare a Windows 10 machine for use in the  
# Corelan 'Advanced' Windows Exploit Development training
#

$env:tempfolder = "c:\corelantemp"
$env:pythonfile = "python-2.7.17.msi"
$env:windbgfile = "winsdksetup.exe"
$env:vscommunityfile = "vs_WDExpress.exe"
$env:monafile = "mona.py"
$env:windbglibfile = "windbglib.py"
$env:pykdfile = "pykd.zip"

# helper functions

function pause()
{
	Write-Host -NoNewLine 'Press any key to continue...';
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
	Write-Host ""
}



# main stuff

Write-Host "Make sure you are running as admin, and that you have an active internet connection before proceeding!"
pause


Write-Output "[+] Creating temp folder $env:tempfolder"
New-Item -Path "c:\" -Name "corelantemp" -ItemType "directory" *>$null

if (Test-Path $env:tempfolder -PathType Container)
{ 
	Write-Output "[+] Downloading packages to temp folder"
	Write-Output "    1. Python 2.7.17"
	Invoke-WebRequest -Uri "https://www.python.org/ftp/python/2.7.17/python-2.7.17.msi" -OutFile "$env:tempfolder\$env:pythonfile" *>$null
	Write-Output "    2. WinDBG"
	Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?linkid=2083338&clcid=0x409" -OutFile "$env:tempfolder\$env:windbgfile" *>$null
	Write-Output "    3. PyKD"
	Invoke-WebRequest -Uri "https://github.com/corelan/windbglib/raw/master/pykd/pykd.zip" -OutFile "$env:tempfolder\$env:pykdfile" *>$null
	Expand-Archive -Path "$env:tempfolder\$env:pykdfile" -DestinationPath "$env:tempfolder\" -Force
	Write-Output "    4. mona.py"
	Invoke-WebRequest -Uri "https://github.com/corelan/mona/raw/master/mona.py" -OutFile "$env:tempfolder\$env:monafile" *>$null
	Write-Output "    5. windbglib.py"
	Invoke-WebRequest -Uri "https://github.com/corelan/windbglib/raw/master/windbglib.py" -OutFile "$env:tempfolder\$env:windbglibfile" *>$null
	Write-Output "    6. Visual Studio 2017 Desktop Express"
	Invoke-WebRequest -Uri "https://aka.ms/vs/15/release/vs_WDExpress.exe" -OutFile "$env:tempfolder\$env:vscommunityfile" *>$null


	Write-Output "[+] Creating System Environment variable _NT_SYMBOL_PATH"
	[Environment]::SetEnvironmentVariable("_NT_SYMBOL_PATH", "srv*c:\symbols*http://msdl.microsoft.com/download/symbols", "Machine")

	Write-Output "[+] Adding c:\Python27 to PATH"
	$oldpath = (Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment' -Name PATH).path
	if ($oldpath -like '*c:\Python27*') 
	{ 
		Write-Output "    PATH already contains entry for Python 2.7"
	}
	else
	{ 
		$newPath = "$oldpath;c:\Python27"
		Set-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment' -Name PATH -Value $newPath
	}
	
	Write-Output "[+] Installing software"
	Write-Output "    1. Python 2.7"
	Start-Process "$env:tempfolder\$env:pythonfile"  -Wait -ArgumentList '/quiet /passive'
	Write-Output "    2. VC++ Redistributable"
	Start-Process "$env:tempfolder\vcredist_x86.exe" -Wait -ArgumentList "/q"
	if (Test-Path "C:\Program Files (x86)\Common Files\microsoft shared\VC\msdia90.dll" -PathType Leaf)
	{
	Write-Output "    3. Register msdia90.dll"
	Start-Process regsvr32 -Wait -ArgumentList '"C:\Program Files (x86)\Common Files\microsoft shared\VC\msdia90.dll" /s'
	}
	else
	{
	Write-Output "    3. *** Error registering msdia90.dll, file does not exist ***"
	}	
	Write-Output "    4. WinDBG"
	Write-Output "       Hold on, this may take a while..."
	Start-Process "$env:tempfolder\$env:windbgfile" -Wait -ArgumentList '/features OptionId.WindowsDesktopDebuggers /ceip off /q'
	Write-Output "    5. PyKD, windbglib & mona"
	Copy-Item -Path "$env:tempfolder\$env:monafile" -Destination "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\"
	Copy-Item -Path "$env:tempfolder\$env:windbglibfile" -Destination "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\"
	Copy-Item -Path "$env:tempfolder\pykd.pyd" -Destination "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\winext\"
	
	Write-Output "    5. Visual Studio 2017 Desktop Express - manual install"
	Start-Process "$env:tempfolder\$env:vscommunityfile" -Wait

	Write-Output "[+] Launching WinDBG to check if everything is ok"
	Start-Process "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\windbg" -ArgumentList '-c ".load pykd.pyd; !peb; !py mona" -o "c:\windows\system32\calc.exe"'
	
	Write-Output "[+] Removing temporary folder again"
	Remove-Item -Path "$env:tempfolder" -recurse -force
	Write-Output "[+] All set"
	Write-Output ""
	Write-Output "==> Please check the WinDBG log window and confirm that:"
	Write-Output "    - the !peb command didn't produce an error message"
	Write-Output "    - the !py mona command resulted in producing a list of available mona commands"
	
}
else
{ 
	Write-Output "*** Oops, folder " + $env:tempfolder + " does not exist"
}

