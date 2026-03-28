# Requires -RunAsAdministrator
#
# (c) Corelan Consulting bv - 2020
# www.corelan-consulting.com
# www.corelan-training.com
# www.corelan-certified.com
# www.corelan.be
#
#
# PowerShell script to prepare a Windows 11/10 machine for use in the  
# Corelan Stack & Heap Exploit Development classes
#
#

$env:tempfolder = "c:\corelantemp"
$env:pythonfile = "python-2.7.18.msi"
$env:windbgfile = "winsdksetup.exe"
$env:vscommunityfile = "vs_WDExpress.exe"
$env:monafile = "mona.py"
$env:windbglibfile = "windbglib.py"
$env:pykdfile = "pykd.zip"
$env:immunityprogramfolder = "C:\Program Files (x86)\Immunity Inc\Immunity Debugger"
$env:immunitypycommandsfolder = Join-Path $env:immunityprogramfolder "PyCommands"

$cmdPath = "$env:WINDIR\System32\cmd.exe"
$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "Corelan CMD Prompt.lnk"
$cmdArguments = '/K "cd /d ""C:\Program Files (x86)\Windows Kits\10\Debuggers\x86"""'

function Confirm-Continue
{
    param(
        [string]$Message = "Ready to continue?"
    )

    while ($true)
    {
        $response = Read-Host "$Message (Y/N)"

        switch ($response.ToLower())
        {
            "y" { return }
            "yes" { return }
            "n" 
            { 
                Write-Output "Aborted by user."
                exit 1 
            }
            "no"
            {
                Write-Output "Aborted by user."
                exit 1
            }
            default
            {
                Write-Output "Please enter Y or N."
            }
        }
    }
}

function Ensure-Folder($path)
{
    if (-not (Test-Path $path -PathType Container))
    {
        New-Item -Path $path -ItemType Directory -Force *>$null
    }
}

function Ensure-Admin
{
    Write-Output "[+] Testing if we have admin privileges"
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal   = New-Object Security.Principal.WindowsPrincipal($currentUser)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        Write-Output "**********************************************"
        Write-Output "!! This script must be run as Administrator !!"
        Write-Output "**********************************************"
        exit 1
    }
    else
    {
        Write-Output "    OK: Administrator privileges detected."
    }
}

function Download-File($uri, $outFile, $label)
{
    Write-Output "    $label"
    Invoke-WebRequest -Uri $uri -OutFile $outFile -UseBasicParsing *>$null
}

function Ensure-Winget
{
    param(
        [string]$TempFolder
    )

    Write-Output "[+] Checking for winget"

    if (Get-Command winget -ErrorAction SilentlyContinue)
    {
        Write-Output "    winget found"
        return
    }

    Write-Output "*** winget was not found on this system."
    Write-Output "***"
    Write-Output "*** This script can continue without installing it,"
    Write-Output "*** but features that rely on winget will not work."
    Write-Output ""

    while ($true)
    {
        $response = Read-Host "Do you want to install winget now? (Y/N)"

        switch ($response.ToLower())
        {
            "y" { break }
            "yes" { break }
            "n"
            {
                Write-Output "Aborted by user."
                exit 1
            }
            "no"
            {
                Write-Output "Aborted by user."
                exit 1
            }
            default
            {
                Write-Output "Please enter Y or N."
            }
        }
    }

    Ensure-Folder $TempFolder

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $appInstallerFile = Join-Path $TempFolder "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    $appInstallerUrl  = "https://aka.ms/getwinget"

    try
    {
        Write-Output "    Downloading App Installer package"

        if (Test-Path $appInstallerFile)
        {
            Remove-Item $appInstallerFile -Force
        }

        Invoke-WebRequest -Uri $appInstallerUrl -OutFile $appInstallerFile -UseBasicParsing *>$null
    }
    catch
    {
        Write-Output "*** Failed to download App Installer package"
        Write-Output "*** Please install winget manually and run this script again."
        exit 1
    }

    try
    {
        Write-Output "    Installing App Installer / winget"
        Add-AppxPackage -Path $appInstallerFile
    }
    catch
    {
        Write-Output "*** Failed to install App Installer package"
        Write-Output "*** Please install winget manually and run this script again."
        exit 1
    }

    Start-Sleep -Seconds 5

    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath    = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path    = "$machinePath;$userPath"

    if (-not (Get-Command winget -ErrorAction SilentlyContinue))
    {
        Write-Output "*** winget still not available after installation"
        Write-Output "*** A reboot or sign-out/sign-in may be required."
        exit 1
    }

    Write-Output "    winget installed successfully"
}



function Test-InternetConnectivity
{
    Write-Output "[+] Checking internet connectivity"

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $testUris = @(
        "https://www.python.org/",
        "https://pypi.org/",
        "https://github.com/",
        "https://aka.ms/"
    )

    foreach ($uri in $testUris)
    {
        try
        {
            Invoke-WebRequest -Uri $uri -Method Head -UseBasicParsing -TimeoutSec 20 *>$null
            Write-Output "    OK   $uri"
        }
        catch
        {
            Write-Output "*** Unable to reach $uri"
            Write-Output "*** Please verify internet connectivity and try again."
            exit 1
        }
    }
}


function Test-PendingWindowsUpdates
{
    Write-Output "[+] Checking for pending Windows Updates"

    try
    {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()

        # Search for updates that are not installed
        $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software'")

        if ($searchResult.Updates.Count -gt 0)
        {
            Write-Output "    There are pending Windows Updates:"
            
            for ($i = 0; $i -lt $searchResult.Updates.Count; $i++)
            {
                $update = $searchResult.Updates.Item($i)
                Write-Output "      - $($update.Title)"
            }

            return $true
        }
        else
        {
            Write-Output "    No pending updates found"
            return $false
        }
    }
    catch
    {
        Write-Output "*** Failed to query Windows Update"
        return $true   # fail-safe: assume updates are pending
    }
}


Ensure-Admin

Write-Host "*** -->> Make sure you have an active internet connection before proceeding! <<-- ***"
Confirm-Continue

Test-InternetConnectivity


Write-Output "[+] Creating temp folder $env:tempfolder"
New-Item -Path "c:\" -Name "corelantemp" -ItemType "directory" *>$null

Ensure-Winget -TempFolder $env:tempfolder


Write-Output "[+] Creating shortcut to cmd.exe on desktop (set to 'Run As Administrator')."

# Create shortcut
$WshShell = New-Object -ComObject WScript.Shell
$shortcut = $WshShell.CreateShortcut($shortcutPath)

$shortcut.TargetPath = $cmdPath
$shortcut.Arguments = $cmdArguments
$shortcut.WorkingDirectory = "$env:WINDIR\System32"
$shortcut.WindowStyle = 1
$shortcut.IconLocation = "$cmdPath,0"
$shortcut.Save()

# --- Enable "Run as Administrator" flag ---
$bytes = [System.IO.File]::ReadAllBytes($shortcutPath)
$bytes[0x15] = $bytes[0x15] -bor 0x20
[System.IO.File]::WriteAllBytes($shortcutPath, $bytes)


# Since all the URLs used for downloading packages uses TLS 1.2 so we need to make sure TLS 1.2 is being used while files been requested otherwise we get SSL error
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Test-Path $env:tempfolder -PathType Container)
{ 
	Write-Output "[+] Downloading packages to temp folder"
	Write-Output "    1. Python 2.7.18"
	Invoke-WebRequest -Uri "https://www.python.org/ftp/python/2.7.18/python-2.7.18.msi" -OutFile "$env:tempfolder\$env:pythonfile" *>$null
	Write-Output "    2. Classic WinDBG"
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

	Write-Output "    5. WinDBGX"
	winget install --id Microsoft.WinDbg -e --source winget --silent --accept-package-agreements --accept-source-agreements

	Write-Output "    6. Visual Studio Code"
	winget install --id Microsoft.VisualStudioCode -e --source winget --silent --accept-package-agreements --accept-source-agreements

	Write-Output "    7. PyKD, windbglib and mona"
	Write-Output "       a. Installing mona.py in WinDBG"
	Copy-Item -Path "$env:tempfolder\$env:monafile" -Destination "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\"
	Copy-Item -Path "$env:tempfolder\$env:windbglibfile" -Destination "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\"
	Copy-Item -Path "$env:tempfolder\pykd.pyd" -Destination "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\winext\"
	
	# if Immunity Debugger was already installed, copy mona.py into the PyCommands folder
	if (Test-Path $env:immunitypycommandsfolder -PathType Container)
	{
		Write-Output "       b. Installing mona.py in Immunity Debugger"
		Copy-Item -Path "$env:tempfolder\$env:monafile" -Destination $env:immunitypycommandsfolder -Force

		# Check for mona.ini
		$monaIniPath = Join-Path $env:immunityprogramfolder "mona.ini"

		if (-not (Test-Path $monaIniPath -PathType Leaf))
		{
			Write-Output "       c. Creating mona.ini"
			"workingfolder=c:\logs\%p" | Out-File -FilePath $monaIniPath -Encoding ASCII
		}
		else
		{
			Write-Output "       c. mona.ini already exists"
		}
	}

	Write-Output "    8. 7Zip"
	winget install --id 7zip.7zip -e --source winget --silent --accept-package-agreements --accept-source-agreements
	
	Write-Output "    9. Visual Studio 2017 Desktop Express - manual install"
	Start-Process "$env:tempfolder\$env:vscommunityfile" -Wait

	Write-Output "[+] Launching WinDBG to check if everything is ok"
	Write-Output "    ==> Please check the WinDBG log window and confirm that:"
	Write-Output "        - the !peb command didn't produce an error message"
	Write-Output "        - the !py mona command resulted in producing a list of available mona commands"
	Start-Process "C:\Program Files (x86)\Windows Kits\10\Debuggers\x86\windbg" -ArgumentList '-c ".load pykd.pyd; !py mona config -set workingfolder c:\logs\%p; !peb; !py mona" -o "c:\windows\system32\calc.exe"'
	
	Write-Output "[+] Removing temporary folder again"
	Remove-Item -Path "$env:tempfolder" -recurse -force
	Write-Output ""
	Write-Output "[+] All set"
	Write-Output ""
	Write-Output "[+] Reboot your VM, and wait for updates to be installed if needed"
}
else
{ 
	Write-Output "*** Oops, folder '$env:tempfolder' does not exist"
}