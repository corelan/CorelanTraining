# Requires -RunAsAdministrator
#
# (c) Corelan Consulting bv - 2026
# www.corelan-consulting.com
# www.corelan-training.com
# www.corelan-certified.com
# www.corelan.be
#
# PowerShell script to install Python3.9 and a PyKD version that is compatible with Python3
# for WinDBG Classic and WinDBGX on Windows 10/11


$ErrorActionPreference = "Stop"

$env:tempfolder                = "C:\corelantemp"
$env:python32installer         = "python-3.9.13.exe"
$env:python64installer         = "python-3.9.13-amd64.exe"

$python32Url                   = "https://www.python.org/ftp/python/3.9.13/python-3.9.13.exe"
$python64Url                   = "https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe"

$classicDbgBase                = "C:\Program Files (x86)\Windows Kits\10\Debuggers"
$engineExt64                   = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions"
$engineExt32                   = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions32"
$userExtensions                = Join-Path $env:USERPROFILE "AppData\dbg\UserExtensions"

$python32Root                  = Join-Path $env:LOCALAPPDATA "Programs\Python\Python39-32"
$python64Root                  = Join-Path $env:LOCALAPPDATA "Programs\Python\Python39"

$python32SitePackages          = Join-Path $python32Root "Lib\site-packages"
$python64SitePackages          = Join-Path $python64Root "Lib\site-packages"

$python32Dll                   = Join-Path $python32Root "python39.dll"
$python64Dll                   = Join-Path $python64Root "python39.dll"

$vcShared32                    = "C:\Program Files (x86)\Common Files\Microsoft Shared\VC"
$msdia140Target                = Join-Path $vcShared32 "msdia140.dll"

$msdia100_64                   = "C:\Program Files\Common Files\microsoft shared\VC\msdia100.dll"
$msdia120_32                   = "C:\Program Files (x86)\Windows Kits\10\App Certification Kit\msdia120.dll"

$regsvr32_64                   = "$env:WINDIR\System32\regsvr32.exe"
$regsvr32_32                   = "$env:WINDIR\SysWOW64\regsvr32.exe"


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

    # Try curl first
    $curl = "$env:SystemRoot\System32\curl.exe"

    if (Test-Path $curl)
    {
        & $curl -L --fail -o $OutFile $Uri
        if ($LASTEXITCODE -eq 0)
        {
            return
        }
    }

    # Fallback to BITS
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


function Remove-FileIfExists($path)
{
    if (Test-Path $path -PathType Leaf)
    {
        Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
    }
}

function Remove-ExistingPyKD
{
    Write-Output "[+] Removing existing pykd.pyd files"

    $paths = @(
        (Join-Path $classicDbgBase "x86\winext\pykd.pyd"),
        (Join-Path $classicDbgBase "x64\winext\pykd.pyd"),
        (Join-Path $engineExt64 "pykd.pyd"),
        (Join-Path $engineExt32 "pykd.pyd"),
        (Join-Path $userExtensions "pykd.pyd")
    )

    foreach ($path in $paths)
    {
        Remove-FileIfExists $path
    }
}

function Get-WingetPythonLines
{
    $wingetOutput = winget list python 2>&1
    if ($LASTEXITCODE -ne 0)
    {
        Write-Output "*** Failed to run 'winget list python'"
        exit 1
    }

    return @($wingetOutput | Where-Object { $_ -match '^\s*Python' })
}

function Validate-WingetPythonSources
{
    param(
        [string]$StageDescription
    )

    Write-Output "[+] Checking Python packages via winget ($StageDescription)"

    $pythonLines = Get-WingetPythonLines

    if ($pythonLines.Count -eq 0)
    {
        Write-Output "    No Python packages reported by 'winget list python'"
        return
    }

    $invalidLines = @()
    foreach ($line in $pythonLines)
    {
        if ($line -notmatch 'winget\s*$')
        {
            $invalidLines += $line
        }
    }

    if ($invalidLines.Count -gt 0)
    {
        Write-Output "*** One or more Python packages do not have source 'winget':"
        foreach ($line in $invalidLines)
        {
            Write-Output "    $line"
        }
        Write-Output ""
        Write-Output "*** Please remove the Python packages above that were not installed from winget,"
        Write-Output "*** then run this script again."
        exit 1
    }
    else
    {
        Write-Output "    All Python entries reported by winget have source 'winget'"
    }
}


function Run-ProcessChecked
{
    param(
        [string]$FilePath,
        [string]$Arguments,
        [string]$Description
    )

    Write-Output "    $Description"
    $proc = Start-Process -FilePath $FilePath -ArgumentList $Arguments -Wait -PassThru
    if ($proc.ExitCode -ne 0)
    {
        Write-Output "*** $Description failed with exit code $($proc.ExitCode)"
        exit 1
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
        if ($ContinueOnMissing)
        {
            Write-Output "    File not found, continuing: $DllPath"
            return
        }
        else
        {
            Write-Output "*** File not found: $DllPath"
            exit 1
        }
    }

    $regsvr = if ($Bitness -eq "x86") { $regsvr32_32 } else { $regsvr32_64 }

    $proc = Start-Process -FilePath $regsvr -ArgumentList "`"$DllPath`" /s" -Wait -PassThru
    if ($proc.ExitCode -ne 0)
    {
        Write-Output "*** regsvr32 failed for $DllPath (exit code $($proc.ExitCode))"
        exit 1
    }
}

function Install-Python39
{
    Write-Output "[+] Installing Python 3.9.13 (32-bit and 64-bit)"

    $python32Args = '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0 Include_launcher=1 SimpleInstall=1 TargetDir="' + $python32Root + '"'
    $python64Args = '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0 Include_launcher=1 SimpleInstall=1 TargetDir="' + $python64Root + '"'

    Run-ProcessChecked -FilePath (Join-Path $env:tempfolder $env:python32installer) -Arguments $python32Args -Description "Installing Python 3.9.13 32-bit"
    Run-ProcessChecked -FilePath (Join-Path $env:tempfolder $env:python64installer) -Arguments $python64Args -Description "Installing Python 3.9.13 64-bit"
}

function Upgrade-Pip
{
    Write-Output "[+] Updating pip"

    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-32 -m pip install --upgrade pip" -Description "Updating pip for Python 3.9 32-bit"
    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-64 -m pip install --upgrade pip" -Description "Updating pip for Python 3.9 64-bit"
}

function Install-PyKD32
{
    Write-Output "[+] Installing PyKD 32-bit"

    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-32 -m pip install pykd" -Description "Installing PyKD with pip for Python 3.9 32-bit"

    Ensure-Folder $engineExt32
    Ensure-Folder $vcShared32

    $pykdPyd32 = Find-PyKDPyd -SitePackagesPath $python32SitePackages
    if (-not $pykdPyd32)
    {
        Write-Output "*** Unable to locate pykd.pyd for Python 3.9 32-bit"
        exit 1
    }

    if (-not (Test-Path $python32Dll -PathType Leaf))
    {
        Write-Output "*** Unable to locate python39.dll for Python 3.9 32-bit"
        exit 1
    }

    $msdia140Source = Find-Msdia140 -SearchRoot $python32Root
    if (-not $msdia140Source)
    {
        Write-Output "*** Unable to locate msdia140.dll for Python 3.9 32-bit"
        exit 1
    }

    Write-Output "    Copying pykd.pyd to $engineExt32"
    Copy-Item -Path $pykdPyd32 -Destination (Join-Path $engineExt32 "pykd.pyd") -Force

    Write-Output "    Copying python39.dll to $engineExt32"
    Copy-Item -Path $python32Dll -Destination (Join-Path $engineExt32 "python39.dll") -Force

    Write-Output "    Copying msdia140.dll to $vcShared32"
    Copy-Item -Path $msdia140Source -Destination $msdia140Target -Force

    Write-Output "    Registering msdia140.dll"
    Register-DllSilent -DllPath $msdia140Target -Bitness x86
}

function Install-VCRuntime2010
{
    Write-Output "[+] Installing Microsoft Visual C++ 2010 SP1 Redistributables (x86 and x64)"

    Run-ProcessChecked -FilePath "winget" `
        -Arguments "install --id Microsoft.VCRedist.2010.x86 -e --source winget --silent --accept-package-agreements --accept-source-agreements" `
        -Description "Installing VC++ 2010 SP1 x86"

    Run-ProcessChecked -FilePath "winget" `
        -Arguments "install --id Microsoft.VCRedist.2010.x64 -e --source winget --silent --accept-package-agreements --accept-source-agreements" `
        -Description "Installing VC++ 2010 SP1 x64"
}

function Install-PyKD64
{
    Write-Output "[+] Installing PyKD 64-bit"

    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-64 -m pip install pykd" -Description "Installing PyKD with pip for Python 3.9 64-bit"

    Ensure-Folder $engineExt64

    $pykdPyd64 = Find-PyKDPyd -SitePackagesPath $python64SitePackages
    if (-not $pykdPyd64)
    {
        Write-Output "*** Unable to locate pykd.pyd for Python 3.9 64-bit"
        exit 1
    }

    if (-not (Test-Path $python64Dll -PathType Leaf))
    {
        Write-Output "*** Unable to locate python39.dll for Python 3.9 64-bit"
        exit 1
    }

    Write-Output "    Copying pykd.pyd to $engineExt64"
    Copy-Item -Path $pykdPyd64 -Destination (Join-Path $engineExt64 "pykd.pyd") -Force

    Write-Output "    Copying python39.dll to $engineExt64"
    Copy-Item -Path $python64Dll -Destination (Join-Path $engineExt64 "python39.dll") -Force

    Write-Output "    Registering msdia100.dll (continue if missing)"
    Register-DllSilent -DllPath $msdia100_64 -Bitness x64 -ContinueOnMissing

    Write-Output "    Registering msdia120.dll"
    Register-DllSilent -DllPath $msdia120_32 -Bitness x86
}


# main stuff

Ensure-Admin


Write-Host "*** -->> Make sure you have an active internet connection before proceeding! <<-- ***"
Confirm-Continue

Test-InternetConnectivity

Write-Output "[+] Creating temp folder $env:tempfolder"
Ensure-Folder $env:tempfolder
Ensure-Folder $engineExt64
Ensure-Folder $engineExt32
Ensure-Folder $userExtensions

Ensure-Winget -TempFolder $env:tempfolder


Remove-ExistingPyKD



Validate-WingetPythonSources -StageDescription "before Python install"

Write-Output "[+] Downloading installers"
Download-File -Uri $python32Url      -OutFile (Join-Path $env:tempfolder $env:python32installer)     -Label "1. Python 3.9.13 32-bit"
Download-File -Uri $python64Url      -OutFile (Join-Path $env:tempfolder $env:python64installer)     -Label "2. Python 3.9.13 64-bit"

Install-Python39
Install-VCRuntime2010

Validate-WingetPythonSources -StageDescription "after Python install"

Upgrade-Pip
Install-PyKD32
Install-PyKD64

Write-Output "[+] Removing temporary folder again"
Remove-Item -Path $env:tempfolder -Recurse -Force -ErrorAction SilentlyContinue

Write-Output ""
Write-Output "[+] All set"
Write-Output ""
Write-Output "[+] You may want to restart WinDBG / WinDBGX if they were already open"