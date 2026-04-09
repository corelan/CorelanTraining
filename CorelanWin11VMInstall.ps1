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
$env:vcredistfile = "vcredist_x86.exe"
$env:monafile = "mona.py"
$env:windbglibfile = "windbglib.py"
$env:pykdExtX86File = "pykd-ext-x86.zip"
$env:pykdExtX64File = "pykd-ext-x64.zip"
$env:pythonUrl = "https://www.python.org/ftp/python/2.7.18/python-2.7.18.msi"
$env:windbgUrl = "https://go.microsoft.com/fwlink/p/?linkid=2083338&clcid=0x409"
$env:vscommunityUrl = "https://aka.ms/vs/15/release/vs_WDExpress.exe"
$env:vcredistUrl = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/runtimes/vcredist_x86.exe"
$env:monaUrl = "https://github.com/corelan/mona/raw/master/mona.py"
$env:windbglibUrl = "https://github.com/corelan/windbglib/raw/master/windbglib.py"
$env:pykdExtX86Url = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x86.zip"
$env:pykdExtX64Url = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x64.zip"
$env:immunityprogramfolder = "C:\Program Files (x86)\Immunity Inc\Immunity Debugger"
$env:immunitypycommandsfolder = Join-Path $env:immunityprogramfolder "PyCommands"

$classicDbgBase = "C:\Program Files (x86)\Windows Kits\10\Debuggers"
# Debugger extension search path for both WinDBG Classic and modern WinDbg.
# %LOCALAPPDATA%\dbg\UserExtensions is the extension gallery location and requires a manifest.
$engineExt32 = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions32"
$engineExt64 = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions"
$vcShared32 = "C:\Program Files (x86)\Common Files\Microsoft Shared\VC"
$msdia140Target = Join-Path $vcShared32 "msdia140.dll"
$msdia100_64 = "C:\Program Files\Common Files\microsoft shared\VC\msdia100.dll"
$msdia120_32 = "C:\Program Files (x86)\Windows Kits\10\App Certification Kit\msdia120.dll"
$regsvr32_64 = "$env:WINDIR\System32\regsvr32.exe"
$regsvr32_32 = "$env:WINDIR\SysWOW64\regsvr32.exe"

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

function Download-File
{
    param(
        [string]$Uri,
        [string]$OutFile,
        [string]$Label
    )

    Write-Output "    $Label"

    if (Test-Path $OutFile)
    {
        Remove-Item $OutFile -Force
    }

    $curl = "$env:SystemRoot\System32\curl.exe"

    if (Test-Path $curl)
    {
        & $curl -L --fail -o $OutFile $Uri
        if ($LASTEXITCODE -eq 0)
        {
            return
        }
    }

    try
    {
        Start-BitsTransfer -Source $Uri -Destination $OutFile
        return
    }
    catch
    {
        Write-Output "*** Download failed with curl and BITS"
        exit 1
    }
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
        Download-File -Uri $appInstallerUrl -OutFile $appInstallerFile -Label "Downloading App Installer package"
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
        "https://files.pythonhosted.org/",
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

function Find-PyKDPyd
{
    param(
        [string]$SitePackagesPath
    )

    $candidates = @(
        (Join-Path $SitePackagesPath "pykd.pyd"),
        (Join-Path $SitePackagesPath "pykd\pykd.pyd")
    )

    foreach ($candidate in $candidates)
    {
        if (Test-Path $candidate -PathType Leaf)
        {
            return $candidate
        }
    }

    $found = Get-ChildItem -Path $SitePackagesPath -Recurse -Filter "pykd.pyd" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found)
    {
        return $found.FullName
    }

    return $null
}

function Find-Msdia140
{
    param(
        [string]$SearchRoot
    )

    $found = Get-ChildItem -Path $SearchRoot -Recurse -Filter "msdia140.dll" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found)
    {
        return $found.FullName
    }

    return $null
}

function Register-DllSilent
{
    param(
        [string]$DllPath,
        [ValidateSet("x86","x64")]
        [string]$Bitness,
        [switch]$ContinueOnMissing
    )

    if (-not (Test-Path $DllPath -PathType Leaf))
    {
        Write-Output "    File not found, continuing: $DllPath"
        return
    }

    $regsvr = if ($Bitness -eq "x86") { $regsvr32_32 } else { $regsvr32_64 }

    try
    {
        $proc = Start-Process -FilePath $regsvr -ArgumentList "`"$DllPath`" /s" -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -ne 0)
        {
            Write-Output "    regsvr32 failed for $DllPath (exit code $($proc.ExitCode)), continuing"
        }
    }
    catch
    {
        Write-Output "    regsvr32 failed for $DllPath, continuing"
    }
}

### MAIN ROUTINE ###

Ensure-Admin

Write-Host "*** -->> Make sure you have an active internet connection before proceeding! <<-- ***"
Confirm-Continue

Test-InternetConnectivity

Write-Output "[+] Creating temp folder $env:tempfolder"
Ensure-Folder $env:tempfolder
Ensure-Folder $engineExt32
Ensure-Folder $engineExt64
Ensure-Folder $vcShared32

Ensure-Winget -TempFolder $env:tempfolder

Write-Output "[+] Creating shortcut to cmd.exe on desktop (set to 'Run As Administrator')."

$WshShell = New-Object -ComObject WScript.Shell
$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $cmdPath
$shortcut.Arguments = $cmdArguments
$shortcut.WorkingDirectory = "$env:WINDIR\System32"
$shortcut.WindowStyle = 1
$shortcut.IconLocation = "$cmdPath,0"
$shortcut.Save()

$bytes = [System.IO.File]::ReadAllBytes($shortcutPath)
$bytes[0x15] = $bytes[0x15] -bor 0x20
[System.IO.File]::WriteAllBytes($shortcutPath, $bytes)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Test-Path $env:tempfolder -PathType Container)
{
    Write-Output "[+] Downloading packages to temp folder"
    Download-File -Uri $env:pythonUrl      -OutFile (Join-Path $env:tempfolder $env:pythonfile)      -Label "1. Python 2.7.18"
    Download-File -Uri $env:windbgUrl      -OutFile (Join-Path $env:tempfolder $env:windbgfile)      -Label "2. Classic WinDBG"
    Download-File -Uri $env:pykdExtX86Url  -OutFile (Join-Path $env:tempfolder $env:pykdExtX86File)  -Label "3. PyKD extension package (x86)"
    Download-File -Uri $env:pykdExtX64Url  -OutFile (Join-Path $env:tempfolder $env:pykdExtX64File)  -Label "4. PyKD extension package (x64)"
    Download-File -Uri $env:monaUrl        -OutFile (Join-Path $env:tempfolder $env:monafile)        -Label "5. mona.py"
    Download-File -Uri $env:windbglibUrl   -OutFile (Join-Path $env:tempfolder $env:windbglibfile)   -Label "6. windbglib.py"
    Download-File -Uri $env:vscommunityUrl -OutFile (Join-Path $env:tempfolder $env:vscommunityfile) -Label "7. Visual Studio 2017 Desktop Express"
    Download-File -Uri $env:vcredistUrl    -OutFile (Join-Path $env:tempfolder $env:vcredistfile)    -Label "8. VC++ 2010 SP1 Redistributable (x86)"

    Remove-Item -Path "$env:tempfolder\pykd-ext-x86" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:tempfolder\pykd-ext-x64" -Recurse -Force -ErrorAction SilentlyContinue
    Expand-Archive -Path (Join-Path $env:tempfolder $env:pykdExtX86File) -DestinationPath "$env:tempfolder\pykd-ext-x86" -Force
    Expand-Archive -Path (Join-Path $env:tempfolder $env:pykdExtX64File) -DestinationPath "$env:tempfolder\pykd-ext-x64" -Force

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
    Start-Process (Join-Path $env:tempfolder $env:pythonfile) -Wait -ArgumentList '/quiet /passive'

    Write-Output "       Updating pip in Python 2.7.18"
    Start-Process "C:\Python27\python.exe" -Wait -ArgumentList '-m ensurepip --default-pip'
    Start-Process "C:\Python27\python.exe" -Wait -ArgumentList '-m pip install --upgrade pip'

    Write-Output "       Installing PyKD via pip in Python 2.7.18"
    Start-Process "C:\Python27\python.exe" -Wait -ArgumentList '-m pip install --upgrade pykd'

    Write-Output "    2. VC++ 2010 SP1 Redistributable (x86)"
    Start-Process (Join-Path $env:tempfolder $env:vcredistfile) -Wait -ArgumentList '/quiet /norestart'

    $pykdPydPath = Find-PyKDPyd -SitePackagesPath "C:\Python27\Lib\site-packages"
    if ($pykdPydPath)
    {
        $pykdSitePackagesFolder = Split-Path -Parent $pykdPydPath
        $msdia140Source = Join-Path $pykdSitePackagesFolder "msdia140.dll"
        if (-not (Test-Path $msdia140Source -PathType Leaf))
        {
            $msdia140Source = Find-Msdia140 -SearchRoot $pykdSitePackagesFolder
        }
    }
    else
    {
        $pykdSitePackagesFolder = "C:\Python27\Lib\site-packages"
        $msdia140Source = Find-Msdia140 -SearchRoot $pykdSitePackagesFolder
    }

    if ($msdia140Source -and (Test-Path $msdia140Source -PathType Leaf))
    {
        Write-Output "    3. Copying msdia140.dll to $vcShared32"
        try
        {
            Ensure-Folder $vcShared32
            Copy-Item -Path $msdia140Source -Destination $msdia140Target -Force -ErrorAction Stop
            Write-Output "       Registering msdia140.dll"
            Register-DllSilent -DllPath $msdia140Target -Bitness x86 -ContinueOnMissing
        }
        catch
        {
            Write-Output "       Failed to copy/register msdia140.dll, continuing"
        }
    }
    else
    {
        Write-Output "    3. msdia140.dll not found in site-packages, skipping"
    }

    Write-Output "    4. Registering msdia100.dll (continue if missing)"
    Register-DllSilent -DllPath $msdia100_64 -Bitness x64 -ContinueOnMissing

    Write-Output "    5. Registering msdia120.dll (continue if missing)"
    Register-DllSilent -DllPath $msdia120_32 -Bitness x86 -ContinueOnMissing

    Write-Output "    6. WinDBG"
    Write-Output "       Hold on, this may take a while..."
    Start-Process (Join-Path $env:tempfolder $env:windbgfile) -Wait -ArgumentList '/features OptionId.WindowsDesktopDebuggers /ceip off /q'

    Write-Output "    7. WinDBGX"
    winget install --id Microsoft.WinDbg -e --source winget --silent --accept-package-agreements --accept-source-agreements

    Write-Output "    8. Visual Studio Code"
    winget install --id Microsoft.VisualStudioCode -e --source winget --silent --accept-package-agreements --accept-source-agreements

    Write-Output "    9. PyKD, windbglib and mona"
    Write-Output "       a. Installing mona.py and windbglib.py in WinDBG"
    Copy-Item -Path (Join-Path $env:tempfolder $env:monafile) -Destination (Join-Path $classicDbgBase 'x86') -Force
    Copy-Item -Path (Join-Path $env:tempfolder $env:windbglibfile) -Destination (Join-Path $classicDbgBase 'x86') -Force

    Write-Output "       b. Installing pykd.dll in debugger extension search path x86 (EngineExtensions32)"
    Copy-Item -Path "$env:tempfolder\pykd-ext-x86\Release\pykd.dll" -Destination (Join-Path $engineExt32 'pykd.dll') -Force

    Write-Output "       c. Installing pykd.dll in debugger extension search path x64 (EngineExtensions)"
    Copy-Item -Path "$env:tempfolder\pykd-ext-x64\Release\pykd.dll" -Destination (Join-Path $engineExt64 'pykd.dll') -Force

    if (Test-Path $env:immunitypycommandsfolder -PathType Container)
    {
        Write-Output "       d. Installing mona.py in Immunity Debugger"
        Copy-Item -Path (Join-Path $env:tempfolder $env:monafile) -Destination $env:immunitypycommandsfolder -Force

        $monaIniPath = Join-Path $env:immunityprogramfolder "mona.ini"

        if (-not (Test-Path $monaIniPath -PathType Leaf))
        {
            Write-Output "       e. Creating mona.ini"
            "workingfolder=c:\logs\%p" | Out-File -FilePath $monaIniPath -Encoding ASCII
        }
        else
        {
            Write-Output "       e. mona.ini already exists"
        }
    }

    Write-Output "    10. 7Zip"
    winget install --id 7zip.7zip -e --source winget --silent --accept-package-agreements --accept-source-agreements

    Write-Output "    11. Visual Studio 2017 Desktop Express - manual install"
    Start-Process (Join-Path $env:tempfolder $env:vscommunityfile) -Wait

    Write-Output "[+] Launching WinDBG to check if everything is ok"
    Write-Output "    ==> Please check the WinDBG log window and confirm that:"
    Write-Output "        - the !peb command didn't produce an error message"
    Write-Output "        - the !py -2 mona command resulted in producing a list of available mona commands"
    Start-Process (Join-Path $classicDbgBase 'x86\windbg') -ArgumentList '-c ".load pykd; !py -2 mona config -set workingfolder c:\logs\%p; !peb; !py -2 mona" -o "c:\windows\system32\calc.exe"'

    Write-Output "[+] Removing temporary folder again"
    Remove-Item -Path $env:tempfolder -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output ""
    Write-Output "[+] All set"
    Write-Output ""
    Write-Output "[+] Reboot your VM, and wait for updates to be installed if needed"
}
else
{
    Write-Output "*** Oops, folder '$env:tempfolder' does not exist"
}
