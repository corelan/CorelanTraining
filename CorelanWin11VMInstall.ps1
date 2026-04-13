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
$env:vc2010redistfile = "vc2010_runtime_redist_x86.exe"
$env:monafile = "mona.py"
$env:windbglibfile = "windbglib.py"
$env:pykdExtX86File = "pykd-ext-x86.zip"
$env:pykdExtX64File = "pykd-ext-x64.zip"
$env:pythonUrl = "https://www.python.org/ftp/python/2.7.18/python-2.7.18.msi"
$env:windbgUrl = "https://go.microsoft.com/fwlink/p/?linkid=2083338&clcid=0x409"
$env:vscommunityUrl = "https://aka.ms/vs/15/release/vs_WDExpress.exe"
$env:vcredistUrl = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/runtimes/vcredist_x86.exe"
$env:vc2010redistUrl = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/runtimes/vc2010_runtime_redist_x86.exe"
$env:monaUrl = "https://github.com/corelan/mona/raw/master/mona.py"
$env:windbglibUrl = "https://github.com/corelan/windbglib/raw/master/windbglib.py"
$env:pykdExtX86Url = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x86.zip"
$env:pykdExtX64Url = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.24/x64.zip"
$env:immunityprogramfolder = "C:\Program Files (x86)\Immunity Inc\Immunity Debugger"
$env:immunitypycommandsfolder = Join-Path $env:immunityprogramfolder "PyCommands"

$classicDbgBase = "C:\Program Files (x86)\Windows Kits\10\Debuggers"
$engineExt32 = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions32"
$engineExt64 = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions"
$vcShared32 = "C:\Program Files (x86)\Common Files\Microsoft Shared\VC"
$msdia140Target = Join-Path $vcShared32 "msdia140.dll"
$msdia100_32 = Join-Path $vcShared32 "msdia100.dll"
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
                Write-Output "Skipping winget installation."
                return $false
            }
            "no"
            {
                Write-Output "Skipping winget installation."
                return $false
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

function Get-WinDbgProgramFolder
{
    $programFilesX86 = ${env:ProgramFiles(x86)}
    if (-not $programFilesX86)
    {
        return $null
    }

    $kitsRoot = Join-Path $programFilesX86 "Windows Kits"
    if (-not (Test-Path $kitsRoot -PathType Container))
    {
        return $null
    }

    $kitRoots = Get-ChildItem -Path $kitsRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName "Debuggers") }

    if (-not $kitRoots)
    {
        return $null
    }

    $kitRoots = $kitRoots | Sort-Object {
        $versionText = $_.Name -replace '[^0-9\.]', ''
        try { [version]$versionText } catch { [version]'0.0' }
    } -Descending

    foreach ($kit in $kitRoots)
    {
        $debuggers = Join-Path $kit.FullName "Debuggers"
        if (Test-Path $debuggers -PathType Container)
        {
            return $debuggers
        }
    }

    return $null
}

function Create-SymbolicLinkDirectory
{
    param(
        [string]$LinkPath,
        [string]$TargetPath
    )

    if (Test-Path $LinkPath)
    {
        Remove-Item -Path $LinkPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    Ensure-Folder (Split-Path -Parent $LinkPath)
    Ensure-Folder $TargetPath

    $arguments = "/c mklink /D `"$LinkPath`" `"$TargetPath`""
    Start-Process -FilePath $cmdPath -ArgumentList $arguments -NoNewWindow -Wait -ErrorAction Stop
}

function Create-SymbolicLinkFile
{
    param(
        [string]$LinkPath,
        [string]$TargetPath
    )

    if (Test-Path $LinkPath)
    {
        $item = Get-Item $LinkPath
        if ($item.LinkType -ne "SymbolicLink")
        {
            Remove-Item -Path $LinkPath -Force -ErrorAction SilentlyContinue
        }
        else
        {
            Remove-Item -Path $LinkPath -Force -ErrorAction SilentlyContinue
        }
    }

    Ensure-Folder (Split-Path -Parent $LinkPath)

    $arguments = "/c mklink `"$LinkPath`" `"$TargetPath`""
    Start-Process -FilePath $cmdPath -ArgumentList $arguments -NoNewWindow -Wait -ErrorAction Stop
}

function Ensure-Admin
{
    Write-Output "[+] Testing if we have admin privileges"
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal   = New-Object Security.Principal.WindowsPrincipal($currentUser)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        Write-Output "**********************************************"
        Write-Output "!! This script is not running as Administrator !!"
        Write-Output "!! Some install/registration steps may fail.  !!"
        Write-Output "!! Continuing anyway.                         !!"
        Write-Output "**********************************************"
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
            return $true
        }
    }

    try
    {
        Start-BitsTransfer -Source $Uri -Destination $OutFile
        return $true
    }
    catch
    {
        Write-Output "*** Download failed with curl and BITS, continuing"
        return $false
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
        return $true
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
                return $false
            }
            "no"
            {
                Write-Output "Aborted by user."
                return $false
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
        Write-Output "*** Failed to download App Installer package, continuing without winget"
        return $false
    }

    try
    {
        Write-Output "    Installing App Installer / winget"
        Add-AppxPackage -Path $appInstallerFile
    }
    catch
    {
        Write-Output "*** Failed to install App Installer package, continuing without winget"
        return $false
    }

    Start-Sleep -Seconds 5

    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath    = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path    = "$machinePath;$userPath"

    if (-not (Get-Command winget -ErrorAction SilentlyContinue))
    {
        Write-Output "*** winget still not available after installation"
        Write-Output "*** A reboot or sign-out/sign-in may be required."
        return $false
    }

    Write-Output "    winget installed successfully"
    return $true
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
            Write-Output "*** Unable to reach $uri, continuing anyway." 
        }
    }
}


function Find-PyKDPyd
{
    param(
        [string]$SitePackagesPath = "C:\Python27\Lib\site-packages"
    )

    if (-not (Test-Path $SitePackagesPath -PathType Container))
    {
        return $null
    }

    $directPath = Join-Path $SitePackagesPath "pykd.pyd"
    if (Test-Path $directPath -PathType Leaf)
    {
        return $directPath
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
        if ($ContinueOnMissing)
        {
            Write-Output "    File not found, continuing: $DllPath"
            return
        }
        else
        {
            Write-Output "    File not found, continuing: $DllPath"
            return
        }
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

function Invoke-NonFatalStep
{
    param(
        [string]$Description,
        [scriptblock]$Action
    )

    try
    {
        & $Action
    }
    catch
    {
        Write-Output "*** Step failed: $Description"
        if ($_.Exception -and $_.Exception.Message)
        {
            Write-Output "*** $($_.Exception.Message)"
        }
        Write-Output "*** Continuing"
    }
}

### MAIN ROUTINE ###

# Check if system is Windows 10 or later

$ver = [System.Environment]::OSVersion.Version

if ($ver.Major -lt 10) {
    Write-Output "*** This script is designed to run on Windows 10 and later ***"
    exit 1
}


Ensure-Admin

Write-Host "*** -->> Make sure you have an active internet connection before proceeding! <<-- ***"
Confirm-Continue

Test-InternetConnectivity

Write-Output "[+] Creating temp folder $env:tempfolder"
Ensure-Folder $env:tempfolder
Ensure-Folder $engineExt32
Ensure-Folder $engineExt64
Ensure-Folder $vcShared32

$wingetAvailable = Ensure-Winget -TempFolder $env:tempfolder

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

    Write-Output "    1. Python 2.7.18"
    [void](Download-File -Uri $env:pythonUrl -OutFile (Join-Path $env:tempfolder $env:pythonfile) -Label "Downloading Python 2.7.18")

    Write-Output "    2. Classic WinDBG"
    [void](Download-File -Uri $env:windbgUrl -OutFile (Join-Path $env:tempfolder $env:windbgfile) -Label "Downloading Classic WinDBG")

    Write-Output "    3. PyKD extension package (x86)"
    [void](Download-File -Uri $env:pykdExtX86Url -OutFile (Join-Path $env:tempfolder $env:pykdExtX86File) -Label "Downloading PyKD extension package (x86)")

    Write-Output "    4. PyKD extension package (x64)"
    [void](Download-File -Uri $env:pykdExtX64Url -OutFile (Join-Path $env:tempfolder $env:pykdExtX64File) -Label "Downloading PyKD extension package (x64)")

    Write-Output "    5. mona.py"
    [void](Download-File -Uri $env:monaUrl -OutFile (Join-Path $env:tempfolder $env:monafile) -Label "Downloading mona.py")

    Write-Output "    6. windbglib.py"
    [void](Download-File -Uri $env:windbglibUrl -OutFile (Join-Path $env:tempfolder $env:windbglibfile) -Label "Downloading windbglib.py")

    Write-Output "    7. Visual Studio 2017 Desktop Express"
    [void](Download-File -Uri $env:vscommunityUrl -OutFile (Join-Path $env:tempfolder $env:vscommunityfile) -Label "Downloading Visual Studio 2017 Desktop Express")

    Write-Output "    8. VC++ Redistributable (x86)"
    [void](Download-File -Uri $env:vcredistUrl -OutFile (Join-Path $env:tempfolder $env:vcredistfile) -Label "Downloading VC++ Redistributable (x86)")

    Write-Output "    9. VC++ 2010 SP1 Redistributable (x86)"
    [void](Download-File -Uri $env:vc2010redistUrl -OutFile (Join-Path $env:tempfolder $env:vc2010redistfile) -Label "Downloading VC++ 2010 SP1 Redistributable (x86)")


    Remove-Item -Path "$env:tempfolder\pykd-ext-x86" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:tempfolder\pykd-ext-x64" -Recurse -Force -ErrorAction SilentlyContinue
    Invoke-NonFatalStep "Expand PyKD extension package (x86)" {
        Expand-Archive -Path (Join-Path $env:tempfolder $env:pykdExtX86File) -DestinationPath "$env:tempfolder\pykd-ext-x86" -Force -ErrorAction Stop
    }
    Invoke-NonFatalStep "Expand PyKD extension package (x64)" {
        Expand-Archive -Path (Join-Path $env:tempfolder $env:pykdExtX64File) -DestinationPath "$env:tempfolder\pykd-ext-x64" -Force -ErrorAction Stop
    }

    Write-Output "[+] Creating System Environment variable _NT_SYMBOL_PATH"
    Invoke-NonFatalStep "Set _NT_SYMBOL_PATH" {
        [Environment]::SetEnvironmentVariable("_NT_SYMBOL_PATH", "srv*c:\symbols*http://msdl.microsoft.com/download/symbols", "Machine")
    }

    Write-Output "[+] Adding c:\Python27 to PATH"
    $oldpath = (Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment' -Name PATH).path
    if ($oldpath -like '*c:\Python27*')
    {
        Write-Output "    PATH already contains entry for Python 2.7"
    }
    else
    {
        $newPath = "$oldpath;c:\Python27"
        Invoke-NonFatalStep "Add c:\Python27 to PATH" {
            Set-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment' -Name PATH -Value $newPath -ErrorAction Stop
        }
    }

    Write-Output "[+] Installing software"
    Write-Output "    1. Python 2.7"
    Invoke-NonFatalStep "Install Python 2.7" {
        Start-Process (Join-Path $env:tempfolder $env:pythonfile) -Wait -ArgumentList '/quiet /passive' -ErrorAction Stop
    }

    Write-Output "       Updating pip in Python 2.7.18"
    Invoke-NonFatalStep "Bootstrap pip in Python 2.7.18" {
        Start-Process "C:\Python27\python.exe" -Wait -ArgumentList '-m ensurepip --default-pip' -ErrorAction Stop
    }
    Invoke-NonFatalStep "Upgrade pip in Python 2.7.18" {
        Start-Process "C:\Python27\python.exe" -Wait -ArgumentList '-m pip install --upgrade pip' -ErrorAction Stop
    }

    Write-Output "       Installing PyKD via pip in Python 2.7.18"
    Invoke-NonFatalStep "Install PyKD via pip in Python 2.7.18" {
        Start-Process "C:\Python27\python.exe" -Wait -ArgumentList '-m pip install --upgrade pykd' -ErrorAction Stop
    }

    Write-Output "    2. VC++ Redistributable (x86)"
    Invoke-NonFatalStep "Install VC++ Redistributable (x86)" {
        Start-Process (Join-Path $env:tempfolder $env:vcredistfile) -Wait -ArgumentList '/quiet /norestart' -ErrorAction Stop
    }
    Invoke-NonFatalStep "Install VC++ 2010 Redistributable (x86)" {
        Start-Process (Join-Path $env:tempfolder $env:vc2010redistfile) -Wait -ArgumentList '/quiet /norestart' -ErrorAction Stop
    }


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
    Register-DllSilent -DllPath $msdia100_32 -Bitness x86 -ContinueOnMissing

    Write-Output "    5. Registering msdia120.dll (continue if missing)"
    Register-DllSilent -DllPath $msdia120_32 -Bitness x86 -ContinueOnMissing

    Write-Output "    6. WinDBG"
    Write-Output "       Hold on, this may take a while..."
    Invoke-NonFatalStep "Install WinDBG" {
        Start-Process (Join-Path $env:tempfolder $env:windbgfile) -Wait -ArgumentList '/features OptionId.WindowsDesktopDebuggers /ceip off /q' -ErrorAction Stop
    }

    Write-Output "    7. WinDBGX"
    if ($wingetAvailable)
    {
        Invoke-NonFatalStep "Install WinDBGX" {
            winget install --id Microsoft.WinDbg -e --source winget --silent --accept-package-agreements --accept-source-agreements
        }
    }
    else
    {
        Write-Output "    winget not available, skipping WinDBGX"
    }

    Write-Output "    8. Visual Studio Code"
    if ($wingetAvailable)
    {
        Invoke-NonFatalStep "Install Visual Studio Code" {
            winget install --id Microsoft.VisualStudioCode -e --source winget --silent --accept-package-agreements --accept-source-agreements
        }
    }
    else
    {
        Write-Output "    winget not available, skipping Visual Studio Code"
    }

    Write-Output "    9. PyKD, windbglib and mona"
    Write-Output "       a. Installing mona.py and windbglib.py in C:\Tools\mona"
    Invoke-NonFatalStep "Install mona.py and windbglib.py in C:\Tools\mona" {
        $toolsMonaFolder = "C:\Tools\mona"
        Ensure-Folder $toolsMonaFolder
        Copy-Item -Path (Join-Path $env:tempfolder $env:monafile) -Destination $toolsMonaFolder -Force -ErrorAction Stop
        Copy-Item -Path (Join-Path $env:tempfolder $env:windbglibfile) -Destination $toolsMonaFolder -Force -ErrorAction Stop
    }

    Write-Output "       b. Installing pykd.dll in WinDBG engine/extensions x86 folder"
    Invoke-NonFatalStep "Install pykd.dll in debugger extension search path x86 (EngineExtensions32)" {
        Copy-Item -Path "$env:tempfolder\pykd-ext-x86\Release\pykd.dll" -Destination (Join-Path $engineExt32 'pykd.dll') -Force -ErrorAction Stop
    }

    Write-Output "       c. Installing pykd.dll in WinDBG engine/extensions x64 folder"
    Invoke-NonFatalStep "Install pykd.dll in debugger extension search path x64 (EngineExtensions)" {
        Copy-Item -Path "$env:tempfolder\pykd-ext-x64\Release\pykd.dll" -Destination (Join-Path $engineExt64 'pykd.dll') -Force -ErrorAction Stop
    }

    if (Test-Path $env:immunitypycommandsfolder -PathType Container)
    {
        Write-Output "       d. Creating mona.py symlink in Immunity Debugger"
        Invoke-NonFatalStep "Create mona.py symlink in Immunity Debugger" {
            $monaPyPath = Join-Path $env:immunitypycommandsfolder "mona.py"
            $toolsMonaPy = "C:\Tools\mona\mona.py"
            Create-SymbolicLinkFile -LinkPath $monaPyPath -TargetPath $toolsMonaPy
        }

        $monaIniPath = Join-Path $env:immunityprogramfolder "mona.ini"

        if (-not (Test-Path $monaIniPath -PathType Leaf))
        {
            Write-Output "       e. Creating mona.ini"
            Invoke-NonFatalStep "Create mona.ini" {
                "workingfolder=c:\logs\%p" | Out-File -FilePath $monaIniPath -Encoding ASCII -ErrorAction Stop
            }
        }
        else
        {
            Write-Output "       e. mona.ini already exists"
        }
    }

    Write-Output "    10. 7Zip"
    if ($wingetAvailable)
    {
        Invoke-NonFatalStep "Install 7Zip" {
            winget install --id 7zip.7zip -e --source winget --silent --accept-package-agreements --accept-source-agreements
        }
    }
    else
    {
        Write-Output "    winget not available, skipping 7Zip"
    }

    Write-Output "    11. Visual Studio 2017 Desktop Express - manual install"
    Invoke-NonFatalStep "Launch Visual Studio 2017 Desktop Express installer" {
        Start-Process (Join-Path $env:tempfolder $env:vscommunityfile) -Wait -ErrorAction Stop
    }

    Write-Output "[+] Launching WinDBG to check if everything is ok"
    Write-Output "    ==> Please check the WinDBG log window and confirm that:"
    Write-Output "        - the !peb command didn't produce an error message"
    Write-Output "        - the !py -2 C:\Tools\mona\mona.py command resulted in a list of available mona commands"
    Invoke-NonFatalStep "Launch WinDBG validation session" {
        Start-Process (Join-Path $classicDbgBase 'x86\windbg') -ArgumentList '-c ".load pykd; !py -2 C:\Tools\mona\mona.py config -set workingfolder c:\logs\%p; !peb; !py -2 C:\Tools\mona\mona.py" -o "c:\windows\system32\calc.exe"' -ErrorAction Stop
    }

    Write-Output "[+] Removing temporary folder again"
    Remove-Item -Path $env:tempfolder -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output ""
    Write-Output "[+] All set"
    Write-Output ""
    Write-Output "[+] Mona usage:"
    Write-Output "    In WinDBG(X) run:"
    Write-Output ""
    Write-Output "      !load pykd"
    Write-Output "      as !mona !py -2 C:\Tools\mona\mona.py"
    Write-Output ""
    Write-Output '    Or run windbg.exe (windbgx.exe) with argument -c "!load pykd; as !mona !py -2 C:\Tools\mona\mona.py"'
    Write-Output "    After that you can simply run '!mona' at the WinDBG(X) command line."
    Write-Output ""
    Write-Output "[+] Reboot your VM, and wait for updates to be installed if needed"
}
else
{
    Write-Output "*** Oops, folder '$env:tempfolder' does not exist"
}
