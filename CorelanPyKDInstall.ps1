# Requires -RunAsAdministrator
#
# (c) Corelan Consulting bv - 2026
# www.corelan-consulting.com
# www.corelan-training.com
# www.corelan-certified.com
# www.corelan.be
#
# PowerShell script to install Python3.9, Python3.14.4 and PyKD versions that are compatible with Python3
# for WinDBG Classic and WinDBGX on Windows 10/11

# DO NOT RUN THIS ON WINDOWS 7

$ErrorActionPreference = "Stop"

$env:tempfolder                = "C:\corelantemp"
$env:python32installer         = "python-3.9.13.exe"
$env:python64installer         = "python-3.9.13-amd64.exe"
$env:python31432installer      = "python-3.14.4.exe"
$env:python31464installer      = "python-3.14.4-amd64.exe"
$env:vc2010redistfile          = "vc2010_runtime_redist_x86.exe"

$python32Url                   = "https://www.python.org/ftp/python/3.9.13/python-3.9.13.exe"
$python64Url                   = "https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe"
$python31432Url                = "https://www.python.org/ftp/python/3.14.4/python-3.14.4.exe"
$python31464Url                = "https://www.python.org/ftp/python/3.14.4/python-3.14.4-amd64.exe"
$vc2010redistUrl               = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/runtimes/vc2010_runtime_redist_x86.exe"
$pykdExtX86Url                 = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.25/x86.zip"
$pykdExtX64Url                 = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd-ext/2.0.0.25/x64.zip"
$pykd314X86ZipUrl              = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd/pykd-0.3.4.15-cp314-win32.zip"
$pykd314X64ZipUrl              = "https://github.com/corelan/CorelanTraining/raw/refs/heads/master/pykd/pykd-0.3.4.15-cp314-amd64.zip"

$classicDbgBase                = "C:\Program Files (x86)\Windows Kits\10\Debuggers"
$engineExt64                   = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions"
$engineExt32                   = Join-Path $env:LOCALAPPDATA "DBG\EngineExtensions32"
$userExtensions                = Join-Path $env:USERPROFILE "AppData\dbg\UserExtensions"

$python32Root                  = Join-Path $env:LOCALAPPDATA "Programs\Python\Python39-32"
$python64Root                  = Join-Path $env:LOCALAPPDATA "Programs\Python\Python39"
$python31432Root               = Join-Path $env:LOCALAPPDATA "Programs\Python\Python314-32"
$python31464Root               = Join-Path $env:LOCALAPPDATA "Programs\Python\Python314"

$python32SitePackages          = Join-Path $python32Root "Lib\site-packages"
$python64SitePackages          = Join-Path $python64Root "Lib\site-packages"

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

    function Test-DownloadedFile([string]$Path)
    {
        if (-not (Test-Path $Path -PathType Leaf))
        {
            return $false
        }
        return ((Get-Item $Path).Length -gt 0)
    }

    function Invoke-WebRequestCompat([string]$RequestUri, [string]$RequestOutFile)
    {
        $iwr = Get-Command Invoke-WebRequest -ErrorAction Stop
        if ($iwr.Parameters.ContainsKey('UseBasicParsing'))
        {
            Invoke-WebRequest -Uri $RequestUri -OutFile $RequestOutFile -UseBasicParsing -ErrorAction Stop
        }
        else
        {
            Invoke-WebRequest -Uri $RequestUri -OutFile $RequestOutFile -ErrorAction Stop
        }
    }

    # Ensure TLS 1.2 where applicable (older Windows / .NET defaults can be too old for modern HTTPS)
    try
    {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }
    catch { }

    $errors = @()

    # 1) Try curl.exe (real binary, not the PowerShell alias)
    $curl = "$env:SystemRoot\System32\curl.exe"
    if (Test-Path $curl)
    {
        $curlArgs = @('-L', '--fail', '--silent', '--show-error', '--retry', '3', '--retry-delay', '2', '-o', $OutFile, $Uri)

        $curlOutput = & $curl @curlArgs 2>&1
        $curlExit = $LASTEXITCODE
        if ($curlExit -eq 0 -and (Test-DownloadedFile $OutFile))
        {
            return
        }

        $curlText = ($curlOutput | Out-String)
        $revocationError =
            ($curlText -match 'revocation') -or
            ($curlText -match '0x80092012') -or
            ($curlText -match 'CERT_TRUST_REVOCATION_STATUS_UNKNOWN') -or
            ($curlText -match 'schannel')

        if ($revocationError)
        {
            if (Test-Path $OutFile) { Remove-Item $OutFile -Force -ErrorAction SilentlyContinue }
            $curlNoRevokeArgs = @('-L', '--fail', '--silent', '--show-error', '--retry', '3', '--retry-delay', '2', '--ssl-no-revoke', '-o', $OutFile, $Uri)
            $curlNoRevokeOutput = & $curl @curlNoRevokeArgs 2>&1
            $curlNoRevokeExit = $LASTEXITCODE
            if ($curlNoRevokeExit -eq 0 -and (Test-DownloadedFile $OutFile))
            {
                return
            }

            $curlNoRevokeText = ($curlNoRevokeOutput | Out-String)
            if ($curlNoRevokeText -match 'unknown option' -or $curlNoRevokeText -match 'is unknown')
            {
                $errors += "curl.exe failed (exit $curlExit): certificate revocation check failed; --ssl-no-revoke unsupported by this curl.exe"
            }
            else
            {
                $errors += "curl.exe failed (exit $curlNoRevokeExit): certificate revocation check failed; retry with --ssl-no-revoke did not succeed"
            }
        }
        else
        {
            $errors += "curl.exe failed (exit $curlExit)"
        }

        if (Test-Path $OutFile) { Remove-Item $OutFile -Force -ErrorAction SilentlyContinue }
    }

    # 2) Fallback to Invoke-WebRequest
    try
    {
        Invoke-WebRequestCompat -RequestUri $Uri -RequestOutFile $OutFile
        if (Test-DownloadedFile $OutFile)
        {
            return
        }
        throw "Downloaded file is missing or empty"
    }
    catch
    {
        $errors += "Invoke-WebRequest failed: $($_.Exception.Message)"
        if (Test-Path $OutFile) { Remove-Item $OutFile -Force -ErrorAction SilentlyContinue }
    }

    # 3) Fallback to BITS (if available)
    try
    {
        if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue)
        {
            Start-BitsTransfer -Source $Uri -Destination $OutFile -ErrorAction Stop
            if (Test-DownloadedFile $OutFile)
            {
                return
            }
            throw "Downloaded file is missing or empty"
        }
        else
        {
            $errors += "BITS not available (Start-BitsTransfer missing)"
        }
    }
    catch
    {
        $errors += "BITS failed: $($_.Exception.Message)"
        if (Test-Path $OutFile) { Remove-Item $OutFile -Force -ErrorAction SilentlyContinue }
    }

    # 4) Last resort: .NET WebClient
    try
    {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($Uri, $OutFile)
        if (Test-DownloadedFile $OutFile)
        {
            return
        }
        throw "Downloaded file is missing or empty"
    }
    catch
    {
        $errors += "WebClient failed: $($_.Exception.Message)"
        if (Test-Path $OutFile) { Remove-Item $OutFile -Force -ErrorAction SilentlyContinue }
    }

    Write-Output "*** Download failed: $Uri"
    foreach ($err in $errors)
    {
        Write-Output "    $err"
    }
    exit 1
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
    Write-Output "[+] Removing existing PyKD files"

    $paths = @(
        (Join-Path $classicDbgBase "x86\winext\pykd.pyd"),
        (Join-Path $classicDbgBase "x64\winext\pykd.pyd"),
        (Join-Path $engineExt64 "pykd.pyd"),
        (Join-Path $engineExt32 "pykd.pyd"),
        (Join-Path $engineExt64 "pykd.dll"),
        (Join-Path $engineExt32 "pykd.dll"),
        (Join-Path $userExtensions "pykd.pyd"),
        (Join-Path $userExtensions "pykd.dll")
    )

    foreach ($path in $paths)
    {
        Remove-FileIfExists $path
    }
}

function Get-WingetPythonLines
{
    # When output is captured (e.g. `2>&1`), winget may prompt for source agreements and hang waiting for input.
    # Force non-interactive execution and auto-accept any required source agreements.
    $wingetArgs = @('list', 'python', '--accept-source-agreements', '--disable-interactivity')
    $wingetCommand = "winget $($wingetArgs -join ' ')"
    $wingetOutput = & winget @wingetArgs 2>&1

    # Older winget versions may not support these flags. Fall back gracefully so we don't break legacy installs.
    if ($LASTEXITCODE -ne 0)
    {
        $wingetText = ($wingetOutput | Out-String)
        if ($wingetText -match '(?i)(unknown argument|unrecognized option|is not recognized|invalid argument).*--disable-interactivity')
        {
            $wingetArgs = @('list', 'python', '--accept-source-agreements')
            $wingetCommand = "winget $($wingetArgs -join ' ')"
            $wingetOutput = & winget @wingetArgs 2>&1
        }
    }

    if ($LASTEXITCODE -ne 0)
    {
        $wingetText = ($wingetOutput | Out-String)
        if ($wingetText -match '(?i)(unknown argument|unrecognized option|is not recognized|invalid argument).*--accept-source-agreements')
        {
            $wingetArgs = @('list', 'python')
            $wingetCommand = "winget $($wingetArgs -join ' ')"
            $wingetOutput = & winget @wingetArgs 2>&1
        }
    }

    if ($LASTEXITCODE -ne 0)
    {
        Write-Output "*** Failed to run '$wingetCommand'"
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

function Test-PythonVersionInstalled
{
    param(
        [string]$PythonRoot,
        [string]$ExpectedVersionPrefix,
        [string]$Label
    )

    $pythonExe = Join-Path $PythonRoot "python.exe"

    if (-not (Test-Path $pythonExe -PathType Leaf))
    {
        Write-Output "    $Label not found at $PythonRoot"
        return $false
    }

    $detectedVersion = $null

    try
    {
        $detectedVersion = (& $pythonExe -c "import sys; print(sys.version.split()[0])" 2>$null | Select-Object -First 1)
        if ($detectedVersion)
        {
            $detectedVersion = $detectedVersion.Trim()
        }
    }
    catch
    {
        $detectedVersion = $null
    }

    if ($detectedVersion -and $detectedVersion -like "$ExpectedVersionPrefix*")
    {
        Write-Output "    $Label already installed ($detectedVersion)"
        return $true
    }

    if ($detectedVersion)
    {
        Write-Output "    $Label found at $PythonRoot, but version is $detectedVersion (expected $ExpectedVersionPrefix)"
    }
    else
    {
        Write-Output "    $Label found at $PythonRoot, but the version could not be verified"
    }

    return $false
}

function Get-ExtractedWheelPath
{
    param(
        [string]$ExtractPath,
        [string]$WheelNamePattern,
        [string]$Label
    )

    $allWheels = @(Get-ChildItem -Path $ExtractPath -Recurse -Filter "*.whl" -File -ErrorAction SilentlyContinue)

    if ($allWheels.Count -eq 0)
    {
        Write-Output "*** Unable to locate a wheel file in $Label archive"
        exit 1
    }

    $matchingWheels = @($allWheels | Where-Object { $_.Name -like $WheelNamePattern })

    if ($matchingWheels.Count -eq 1)
    {
        return $matchingWheels[0].FullName
    }

    if ($allWheels.Count -eq 1)
    {
        return $allWheels[0].FullName
    }

    Write-Output "*** Unable to determine the correct wheel file in $Label archive"
    foreach ($wheel in $allWheels)
    {
        Write-Output "    Found wheel: $($wheel.Name)"
    }
    exit 1
}

function Install-Python39
{
    Write-Output "[+] Installing Python 3.9.13 (32-bit and 64-bit)"

    $python32Args = '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0 Include_launcher=1 SimpleInstall=1 TargetDir="' + $python32Root + '"'
    $python64Args = '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0 Include_launcher=1 SimpleInstall=1 TargetDir="' + $python64Root + '"'

    if (-not (Test-PythonVersionInstalled -PythonRoot $python32Root -ExpectedVersionPrefix "3.9.13" -Label "Python 3.9.13 32-bit"))
    {
        Download-File -Uri $python32Url -OutFile (Join-Path $env:tempfolder $env:python32installer) -Label "1. Python 3.9.13 32-bit"
        Run-ProcessChecked -FilePath (Join-Path $env:tempfolder $env:python32installer) -Arguments $python32Args -Description "Installing Python 3.9.13 32-bit"
    }

    if (-not (Test-PythonVersionInstalled -PythonRoot $python64Root -ExpectedVersionPrefix "3.9.13" -Label "Python 3.9.13 64-bit"))
    {
        Download-File -Uri $python64Url -OutFile (Join-Path $env:tempfolder $env:python64installer) -Label "2. Python 3.9.13 64-bit"
        Run-ProcessChecked -FilePath (Join-Path $env:tempfolder $env:python64installer) -Arguments $python64Args -Description "Installing Python 3.9.13 64-bit"
    }
}

function Upgrade-Pip
{
    Write-Output "[+] Updating pip"

    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-32 -m pip install --upgrade pip" -Description "Updating pip for Python 3.9 32-bit"
    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-64 -m pip install --upgrade pip" -Description "Updating pip for Python 3.9 64-bit"
}

function Install-Keystone-engine
{
    Write-Output "[+] Installing Keystone-Engine"
    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-32 -m pip install keystone-engine" -Description "Installing keystone-engine for Python 3.9 32-bit"
    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-64 -m pip install keystone-engine" -Description "Installing keystone-engine for Python 3.9 64-bit"
}

function Install-PyKD32
{
    Write-Output "[+] Installing PyKD 32-bit"

    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-32 -m pip install pykd" -Description "Installing PyKD with pip for Python 3.9 32-bit"

    Ensure-Folder $engineExt32
    Ensure-Folder $vcShared32

    $msdia140Source = Find-Msdia140 -SearchRoot $python32Root
    if (-not $msdia140Source)
    {
        Write-Output "*** Unable to locate msdia140.dll for Python 3.9 32-bit"
        exit 1
    }

    Write-Output "    Copying msdia140.dll to $vcShared32"
    Copy-Item -Path $msdia140Source -Destination $msdia140Target -Force

    Write-Output "    Registering msdia140.dll"
    Register-DllSilent -DllPath $msdia140Target -Bitness x86
}

function Install-VCRuntime2010
{
    Write-Output "[+] Installing VC++ 2010 SP1 Redistributable (x86)"
    
    Invoke-NonFatalStep "Install VC++ 2010 Redistributable (x86)" {
        Start-Process (Join-Path $env:tempfolder $env:vc2010redistfile) -Wait -ArgumentList '/quiet /norestart' -ErrorAction Stop
    }
}

function Install-PyKD64
{
    Write-Output "[+] Installing PyKD 64-bit"

    Run-ProcessChecked -FilePath "py" -Arguments "-3.9-64 -m pip install pykd" -Description "Installing PyKD with pip for Python 3.9 64-bit"

    Ensure-Folder $engineExt64

    Write-Output "    Registering msdia100.dll (continue if missing)"
    Register-DllSilent -DllPath $msdia100_64 -Bitness x64 -ContinueOnMissing

    Write-Output "    Registering msdia120.dll"
    Register-DllSilent -DllPath $msdia120_32 -Bitness x86
}

function Install-Python314PyKD
{
    Write-Output "[+] Installing Python 3.14.4 and matching PyKD wheels"

    $python31432Args = '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0 Include_launcher=1 SimpleInstall=1 TargetDir="' + $python31432Root + '"'
    $python31464Args = '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0 Include_launcher=1 SimpleInstall=1 TargetDir="' + $python31464Root + '"'

    if (-not (Test-PythonVersionInstalled -PythonRoot $python31432Root -ExpectedVersionPrefix "3.14.4" -Label "Python 3.14.4 32-bit"))
    {
        Download-File -Uri $python31432Url -OutFile (Join-Path $env:tempfolder $env:python31432installer) -Label "4. Python 3.14.4 32-bit"
        Run-ProcessChecked -FilePath (Join-Path $env:tempfolder $env:python31432installer) -Arguments $python31432Args -Description "Installing Python 3.14.4 32-bit"
    }

    if (-not (Test-PythonVersionInstalled -PythonRoot $python31464Root -ExpectedVersionPrefix "3.14.4" -Label "Python 3.14.4 64-bit"))
    {
        Download-File -Uri $python31464Url -OutFile (Join-Path $env:tempfolder $env:python31464installer) -Label "5. Python 3.14.4 64-bit"
        Run-ProcessChecked -FilePath (Join-Path $env:tempfolder $env:python31464installer) -Arguments $python31464Args -Description "Installing Python 3.14.4 64-bit"
    }

    Run-ProcessChecked -FilePath "py" -Arguments "-3.14-32 -m pip install --upgrade pip" -Description "Updating pip for Python 3.14.4 32-bit"
    Run-ProcessChecked -FilePath "py" -Arguments "-3.14 -m pip install --upgrade pip" -Description "Updating pip for Python 3.14.4 64-bit"

    Run-ProcessChecked -FilePath "py" -Arguments "-3.14-32 -m pip install keystone-engine" -Description "Installing keystone-engine for Python 3.14.4 32-bit"
    Run-ProcessChecked -FilePath "py" -Arguments "-3.14 -m pip install keystone-engine" -Description "Installing keystone-engine for Python 3.14.4 64-bit"

    $pykd314X86Zip     = Join-Path $env:tempfolder "pykd-0.3.4.15-cp314-win32.zip"
    $pykd314X64Zip     = Join-Path $env:tempfolder "pykd-0.3.4.15-cp314-amd64.zip"
    $pykd314X86Extract = Join-Path $env:tempfolder "pykd-0.3.4.15-cp314-win32"
    $pykd314X64Extract = Join-Path $env:tempfolder "pykd-0.3.4.15-cp314-amd64"

    Download-File -Uri $pykd314X86ZipUrl -OutFile $pykd314X86Zip -Label "6. PyKD cp314 x86"
    Download-File -Uri $pykd314X64ZipUrl -OutFile $pykd314X64Zip -Label "7. PyKD cp314 x64"

    Remove-Item -Path $pykd314X86Extract -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $pykd314X64Extract -Recurse -Force -ErrorAction SilentlyContinue

    Ensure-Folder $pykd314X86Extract
    Ensure-Folder $pykd314X64Extract

    Write-Output "    Extracting PyKD cp314 x86"
    Expand-Archive -Path $pykd314X86Zip -DestinationPath $pykd314X86Extract -Force

    Write-Output "    Extracting PyKD cp314 x64"
    Expand-Archive -Path $pykd314X64Zip -DestinationPath $pykd314X64Extract -Force

    $pykd314Wheel32 = Get-ExtractedWheelPath -ExtractPath $pykd314X86Extract -WheelNamePattern "*cp314*win32*.whl" -Label "PyKD cp314 x86"
    $pykd314Wheel64 = Get-ExtractedWheelPath -ExtractPath $pykd314X64Extract -WheelNamePattern "*cp314*amd64*.whl" -Label "PyKD cp314 x64"

    Write-Output "    Using wheel: $(Split-Path $pykd314Wheel32 -Leaf)"
    Run-ProcessChecked -FilePath "py" -Arguments ('-3.14-32 -m pip install "' + $pykd314Wheel32 + '"') -Description "Installing PyKD for Python 3.14.4 32-bit"

    Write-Output "    Using wheel: $(Split-Path $pykd314Wheel64 -Leaf)"
    Run-ProcessChecked -FilePath "py" -Arguments ('-3.14 -m pip install "' + $pykd314Wheel64 + '"') -Description "Installing PyKD for Python 3.14.4 64-bit"
}


function Install-Python27PyKD
{
    $python27Root = "C:\Python27"
    $pythonExe = Join-Path $python27Root "python.exe"

    Write-Output "[+] Checking for Python 2.7 at $python27Root"

    if (-not (Test-Path $python27Root -PathType Container))
    {
        Write-Output "    Python 2.7 not found at $python27Root, skipping"
        return
    }

    if (-not (Test-Path $pythonExe -PathType Leaf))
    {
        Write-Output "    python.exe not found at $pythonExe, skipping"
        return
    }

    Write-Output "       Updating pip in Python 2.7"
    Invoke-NonFatalStep "Bootstrap pip in Python 2.7" {
        Start-Process $pythonExe -Wait -ArgumentList '-m ensurepip --default-pip' -ErrorAction Stop
    }
    Invoke-NonFatalStep "Upgrade pip in Python 2.7" {
        Start-Process $pythonExe -Wait -ArgumentList '-m pip install --upgrade pip' -ErrorAction Stop
    }

    Write-Output "       Installing PyKD via pip in Python 2.7"
    Invoke-NonFatalStep "Install PyKD via pip in Python 2.7" {
        Start-Process $pythonExe -Wait -ArgumentList '-m pip install --upgrade pykd' -ErrorAction Stop
    }

    Write-Output "       Installing keystone-engine via pip in Python 2.7"
    Invoke-NonFatalStep "Install keystone-engine via pip in Python 2.7" {
        Start-Process $pythonExe -Wait -ArgumentList '-m pip install --upgrade keystone-engine' -ErrorAction Stop
    }

}

function Install-PyKDExtensions
{
    Write-Output "[+] Installing PyKD-Ext into WinDBG EngineExtensions folders"

    $pykdExtX86Zip      = Join-Path $env:tempfolder "pykd-ext-x86.zip"
    $pykdExtX64Zip      = Join-Path $env:tempfolder "pykd-ext-x64.zip"
    $pykdExtX86Extract  = Join-Path $env:tempfolder "pykd-ext-x86"
    $pykdExtX64Extract  = Join-Path $env:tempfolder "pykd-ext-x64"

    Download-File -Uri $pykdExtX86Url -OutFile $pykdExtX86Zip -Label "3. PyKD-Ext x86"
    Download-File -Uri $pykdExtX64Url -OutFile $pykdExtX64Zip -Label "4. PyKD-Ext x64"

    Remove-Item -Path $pykdExtX86Extract -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $pykdExtX64Extract -Recurse -Force -ErrorAction SilentlyContinue

    Ensure-Folder $pykdExtX86Extract
    Ensure-Folder $pykdExtX64Extract

    Write-Output "    Extracting PyKD-Ext x86"
    Expand-Archive -Path $pykdExtX86Zip -DestinationPath $pykdExtX86Extract -Force

    Write-Output "    Extracting PyKD-Ext x64"
    Expand-Archive -Path $pykdExtX64Zip -DestinationPath $pykdExtX64Extract -Force

    $pykdDllX86 = Join-Path $pykdExtX86Extract "pykd.dll"
    $pykdDllX64 = Join-Path $pykdExtX64Extract "pykd.dll"

    if (-not (Test-Path $pykdDllX86 -PathType Leaf))
    {
        Write-Output "*** Unable to locate x86 pykd.dll in PyKD-Ext archive"
        exit 1
    }

    if (-not (Test-Path $pykdDllX64 -PathType Leaf))
    {
        Write-Output "*** Unable to locate x64 pykd.dll in PyKD-Ext archive"
        exit 1
    }

    Ensure-Folder $engineExt32
    Ensure-Folder $engineExt64

    Write-Output "    Copying x86 pykd.dll to $engineExt32"
    Copy-Item -Path $pykdDllX86 -Destination (Join-Path $engineExt32 "pykd.dll") -Force

    Write-Output "    Copying x64 pykd.dll to $engineExt64"
    Copy-Item -Path $pykdDllX64 -Destination (Join-Path $engineExt64 "pykd.dll") -Force
}

function Show-CorelanBanner
{
    Write-Output "=============================================="
    Write-Output " Corelan PyKD Install"
    Write-Output " www.corelan-training.com"
    Write-Output "=============================================="
    Write-Output ""
}

# main stuff

Ensure-Admin
Show-CorelanBanner

# Check if system is Windows 10 or later

$ver = [System.Environment]::OSVersion.Version

if ($ver.Major -lt 10) {
    Write-Output "*** This script is designed to run on Windows 10 and later ***"
    exit 1
}


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
Download-File -Uri $vc2010redistUrl  -OutFile (Join-Path $env:tempfolder $env:vc2010redistfile)    -Label "3. VC++ 2010 SP1 Redistributable (x86)"

Install-Python39
Install-VCRuntime2010

Validate-WingetPythonSources -StageDescription "after Python install"

Upgrade-Pip
Install-PyKD32
Install-PyKD64
Install-Python314PyKD
Install-Python27PyKD
Install-PyKDExtensions
Install-Keystone-engine

Write-Output "[+] Removing temporary folder again"
Remove-Item -Path $env:tempfolder -Recurse -Force -ErrorAction SilentlyContinue

Write-Output ""
Write-Output "[+] All set"
Write-Output ""
Write-Output "[+] You may want to restart WinDBG / WinDBGX if they were already open"
