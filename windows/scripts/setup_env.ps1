<#
    Africa_DN_Testing - environment provisioning script
    ---------------------------------------------------
    Ensures the machine can run domain_dns_test.py:
      1. a real Python interpreter (installs one if missing)
      2. pip dependencies openpyxl / requests (offline wheels supported)
      3. dig.exe (ISC BIND), installed through winget when missing
    Finally it writes the resolved interpreter path to cache\python_path.txt
    so that run_dns_test.bat can launch the program with the exact interpreter.

    That cache file is MACHINE SPECIFIC (it records this computer's name plus
    an absolute interpreter path such as <project>\.venv\Scripts\python.exe),
    therefore it must never be copied to another computer together with the
    project - run_dns_test.bat ignores it when the computer name differs.

    NOTE: all messages are ASCII on purpose. Windows PowerShell 5.1 reads
    .ps1 files as ANSI unless they carry a UTF-8 BOM, so non-ASCII literals
    in this file would come out garbled (the Chinese user-facing text lives in
    domain_dns_test.py and in the .md documents instead).

    Usage examples:
      powershell -NoProfile -ExecutionPolicy Bypass -File scripts\setup_env.ps1
      powershell ... -File scripts\setup_env.ps1 -DryRun
      powershell ... -File scripts\setup_env.ps1 -DryRun -FakeMissingPython -FakeMissingDig
      powershell ... -File scripts\setup_env.ps1 -NoInstall      (check only)
      powershell ... -File scripts\setup_env.ps1 -IndexUrl https://pypi.tuna.tsinghua.edu.cn/simple
#>
[CmdletBinding()]
param(
    # Only report what would be done, never change the machine.
    [switch]$DryRun,
    # Never install anything; just check and report.
    [switch]$NoInstall,
    # Pretend the component is missing (only useful together with -DryRun,
    # it lets you verify the automatic-installation branches safely).
    [switch]$FakeMissingPython,
    [switch]$FakeMissingDeps,
    [switch]$FakeMissingDig,
    # Where to create the virtual environment ('' = <project>\.venv).
    [string]$VenvPath = '',
    # Extra pip index (mirror) used when the default PyPI is unreachable.
    [string]$IndexUrl = ''
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$CacheDir = Join-Path $ProjectRoot 'cache'
$WheelDir = Join-Path $ProjectRoot 'wheels'
$Requirements = Join-Path $ProjectRoot 'requirements.txt'
$PythonPathFile = Join-Path $CacheDir 'python_path.txt'
if ([string]::IsNullOrWhiteSpace($VenvPath)) {
    $VenvPath = Join-Path $ProjectRoot '.venv'
} elseif (-not [System.IO.Path]::IsPathRooted($VenvPath)) {
    $VenvPath = Join-Path $ProjectRoot $VenvPath
}
$VenvPath = [System.IO.Path]::GetFullPath($VenvPath)

$script:NeedInstall = -not ($DryRun -or $NoInstall)
$script:Fatal = $false

function Write-Step([string]$text) { Write-Host ""; Write-Host "[*] $text" -ForegroundColor Cyan }
function Write-Ok([string]$text)   { Write-Host "    OK   $text" -ForegroundColor Green }
function Write-Warn2([string]$text){ Write-Host "    WARN $text" -ForegroundColor Yellow }
function Write-Err2([string]$text) { Write-Host "    FAIL $text" -ForegroundColor Red }

# Runs an external program and returns its exit code.
# Native commands that write to stderr would otherwise raise a
# NativeCommandError because $ErrorActionPreference is 'Stop'.
function Invoke-Native {
    param(
        [string]$File,
        [string[]]$Arguments = @(),
        [switch]$Quiet
    )
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $code = -1
    try {
        if ($Quiet) {
            & $File @Arguments 2>&1 | Out-Null
        } else {
            & $File @Arguments 2>&1 | ForEach-Object { Write-Host "    $_" }
        }
        $code = $LASTEXITCODE
    } catch {
        Write-Warn2 ("could not run " + $File + ": " + $_.Exception.Message)
        $code = -1
    } finally {
        $ErrorActionPreference = $previous
    }
    return $code
}

function Test-Simulated([string]$name) {
    switch ($name) {
        'python' { return [bool]$FakeMissingPython }
        'deps'   { return [bool]$FakeMissingDeps }
        'dig'    { return [bool]$FakeMissingDig }
        default  { return $false }
    }
}

function Test-Winget {
    return [bool](Get-Command winget.exe -ErrorAction SilentlyContinue)
}

function Test-RealPython([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return $false }
    if (-not (Test-Path $path)) { return $false }
    # The Microsoft Store alias in WindowsApps is a 0-byte stub -> reject it.
    try { if ((Get-Item $path).Length -le 0) { return $false } } catch { return $false }
    return ((Invoke-Native -File $path -Arguments @('-c', 'import sys') -Quiet) -eq 0)
}

function Find-Python {
    if (Test-Simulated 'python') { return $null }

    $candidates = New-Object System.Collections.Generic.List[string]

    $venvPy = Join-Path $VenvPath 'Scripts\python.exe'
    $candidates.Add($venvPy) | Out-Null

    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd) { $candidates.Add($cmd.Source) | Out-Null }

    $pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        $previous = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $out = & $pyLauncher.Source -3 -c "import sys;print(sys.executable)" 2>$null
            if ($out) { $candidates.Add(($out | Select-Object -First 1).ToString().Trim()) | Out-Null }
        } catch { } finally { $ErrorActionPreference = $previous }
    }

    $patterns = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python3*\python.exe'),
        (Join-Path $env:ProgramFiles 'Python3*\python.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Python3*\python.exe'),
        'C:\Python3*\python.exe'
    )
    foreach ($pattern in $patterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        try {
            Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
                ForEach-Object { $candidates.Add($_.FullName) | Out-Null }
        } catch { }
    }

    # Newest first (Python313 > Python311)
    foreach ($candidate in ($candidates | Where-Object { $_ } | Sort-Object -Descending)) {
        if (Test-RealPython $candidate) { return (Get-Item $candidate).FullName }
    }
    return $null
}

function Install-PythonViaWinget {
    if (-not (Test-Winget)) { Write-Warn2 'winget is not available'; return $false }
    foreach ($id in @('Python.Python.3.13', 'Python.Python.3.12', 'Python.Python.3.11')) {
        foreach ($extra in @(@('--scope', 'user'), @())) {
            $wingetArgs = @('install', '--id', $id, '--exact', '--source', 'winget',
                            '--silent', '--accept-package-agreements', '--accept-source-agreements') + $extra
            Write-Host ("    running: winget " + ($wingetArgs -join ' '))
            $code = Invoke-Native -File 'winget.exe' -Arguments $wingetArgs
            if ($code -eq 0) {
                Write-Ok "installed $id via winget"
                return $true
            }
        }
    }
    return $false
}

function Install-PythonFromPythonOrg {
    Write-Host '    resolving the newest stable CPython build from python.org ...'
    $index = Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/' -UseBasicParsing -TimeoutSec 60
    $versions = [regex]::Matches($index.Content, 'href="(3\.1[0-9]\.[0-9]+)/"') |
        ForEach-Object { $_.Groups[1].Value } |
        Sort-Object { [version]$_ } -Unique |
        Select-Object -Last 10 |
        Sort-Object { [version]$_ } -Descending

    foreach ($version in $versions) {
        $url = "https://www.python.org/ftp/python/$version/python-$version-amd64.exe"
        try {
            Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -TimeoutSec 30 | Out-Null
        } catch {
            Write-Host "    no amd64 installer for $version, trying older"
            continue
        }
        $installer = Join-Path $env:TEMP "python-$version-amd64.exe"
        Write-Host "    downloading $url"
        Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing -TimeoutSec 1800
        Write-Host "    running silent install (user scope, no admin needed) ..."
        $installArgs = @('/quiet', 'InstallAllUsers=0', 'PrependPath=1',
                         'Include_pip=1', 'Include_launcher=1', 'Shortcuts=0')
        Start-Process -FilePath $installer -ArgumentList $installArgs -Wait
        try { Remove-Item -Force $installer -ErrorAction SilentlyContinue } catch { }
        return $true
    }
    return $false
}

function Ensure-Python {
    Write-Step 'Step 1/4  Python interpreter'
    $python = Find-Python
    if ($python) {
        Write-Ok "found: $python"
        return $python
    }
    Write-Warn2 'no usable Python found'

    if (-not $script:NeedInstall) {
        if ($DryRun) {
            Write-Host '    (dry-run) would install Python 3 automatically:'
            Write-Host '      - winget install --id Python.Python.3.13 --exact --scope user --silent'
            Write-Host '      - fallback: download python-<ver>-amd64.exe from python.org and run'
            Write-Host '        /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1'
        } else {
            Write-Err2 'Python is missing and -NoInstall was given'
            $script:Fatal = $true
        }
        return $null
    }

    Write-Host '    installing Python 3 ...'
    $installed = $false
    try { $installed = Install-PythonViaWinget } catch { Write-Warn2 ("winget failed: " + $_.Exception.Message) }
    if (-not $installed) {
        try { $installed = Install-PythonFromPythonOrg } catch { Write-Err2 ("download install failed: " + $_.Exception.Message) }
    }
    if (-not $installed) {
        Write-Err2 'automatic Python installation failed'
        $script:Fatal = $true
        return $null
    }

    # PATH of this process is not refreshed, so search the well-known folders.
    $python = Find-Python
    if ($python) { Write-Ok "installed: $python" } else { Write-Err2 'Python was installed but could not be located'; $script:Fatal = $true }
    return $python
}

function Test-Dependencies([string]$python) {
    if (Test-Simulated 'deps') { return $false }
    return ((Invoke-Native -File $python -Arguments @('-c', 'import openpyxl, requests') -Quiet) -eq 0)
}

function Get-PipBaseArguments {
    $pipArgs = @('-m', 'pip', '--disable-pip-version-check')
    if (-not [string]::IsNullOrWhiteSpace($IndexUrl)) { $pipArgs += @('-i', $IndexUrl) }
    return $pipArgs
}

function Install-Dependencies([string]$python) {
    if (-not $script:NeedInstall) { return $false }

    $wheelFiles = @()
    if (Test-Path $WheelDir) {
        $wheelFiles = @(Get-ChildItem -Path $WheelDir -Filter '*.whl' -ErrorAction SilentlyContinue)
    }
    if ($wheelFiles.Count -gt 0) {
        Write-Host ("    installing from local wheels (offline): " + (Split-Path $WheelDir -Leaf) +
                    " (" + $wheelFiles.Count + " files)")
        $pipArgs = @('-m', 'pip', 'install', '--disable-pip-version-check',
                     '--no-index', '--find-links', $WheelDir, '-r', $Requirements)
        if ((Invoke-Native -File $python -Arguments $pipArgs) -eq 0) { return $true }
        Write-Warn2 'offline wheel install failed, falling back to the online indexes'
    }

    $attempts = @()
    $attempts += ,@((Get-PipBaseArguments))
    if ([string]::IsNullOrWhiteSpace($IndexUrl)) {
        $attempts += ,@(@('-m', 'pip', '--disable-pip-version-check', '-i',
                          'https://pypi.tuna.tsinghua.edu.cn/simple'))
        $attempts += ,@(@('-m', 'pip', '--disable-pip-version-check', '-i',
                          'https://mirrors.aliyun.com/pypi/simple/'))
    }
    foreach ($pipArgs in $attempts) {
        Write-Host ("    running: " + (Split-Path $python -Leaf) + " " + ($pipArgs -join ' ') +
                    " install -r requirements.txt")
        $code = Invoke-Native -File $python -Arguments ($pipArgs + @('install', '-r', $Requirements))
        if ($code -eq 0) { return $true }
        Write-Warn2 'pip failed, trying another index'
    }
    return $false
}

function Get-VenvPython([string]$basePython) {
    $venvPy = Join-Path $VenvPath 'Scripts\python.exe'
    if (Test-RealPython $venvPy) { return $venvPy }
    if (-not $script:NeedInstall) { return $null }

    $venvArgs = @('-m', 'venv')
    if (Test-Path $VenvPath) {
        # A .venv that exists but is not usable is almost always a folder
        # copied from another computer: its pyvenv.cfg points at a base
        # Python which does not exist here. Rebuild it from scratch.
        Write-Warn2 'the existing .venv is unusable here (copied from another computer?) - rebuilding it'
        $venvArgs += '--clear'
    }
    $venvArgs += $VenvPath

    Write-Host "    creating virtual environment: $VenvPath"
    try {
        Invoke-Native -File $basePython -Arguments $venvArgs | Out-Null
    } catch {
        Write-Warn2 ("venv creation failed: " + $_.Exception.Message)
    }
    if (Test-RealPython $venvPy) { return $venvPy }
    Write-Warn2 'falling back to the system interpreter'
    return $null
}

function Ensure-Dependencies([string]$python) {
    Write-Step 'Step 2/4  pip dependencies (openpyxl, requests)'
    if (Test-Dependencies $python) {
        Write-Ok 'already installed'
        return $true
    }
    if (-not $script:NeedInstall) {
        Write-Warn2 'dependencies missing (not installing because of -DryRun/-NoInstall)'
        return $false
    }
    Write-Host '    installing ...'
    if (Install-Dependencies $python) {
        if (Test-Dependencies $python) { Write-Ok 'installed'; return $true }
    }
    Write-Err2 'could not install dependencies (check the network / proxy, or ship a wheels folder)'
    return $false
}

function Find-Dig {
    if (Test-Simulated 'dig') { return $null }

    $cmd = Get-Command dig.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $direct = @(
        (Join-Path $ProjectRoot 'tools\bind\dig.exe'),
        (Join-Path $ProjectRoot 'dig.exe'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Links\dig.exe')
    )
    foreach ($path in $direct) {
        if (Test-Path $path) { return (Get-Item $path).FullName }
    }

    $roots = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'),
        (Join-Path $env:ProgramFiles 'ISC BIND*'),
        (Join-Path ${env:ProgramFiles(x86)} 'ISC BIND*')
    )
    foreach ($root in $roots) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        try {
            $hit = Get-ChildItem -Path $root -Recurse -Filter 'dig.exe' -ErrorAction SilentlyContinue |
                   Where-Object { $_.Length -gt 0 } | Select-Object -First 1
            if ($hit) { return $hit.FullName }
        } catch { }
    }
    return $null
}

function Install-DigViaWinget {
    if (-not (Test-Winget)) { return $false }
    $wingetArgs = @('install', '--id', 'ISC.Bind', '--exact', '--source', 'winget',
                    '--silent', '--accept-package-agreements', '--accept-source-agreements')
    Write-Host ("    running: winget " + ($wingetArgs -join ' '))
    Write-Host '    (a UAC prompt may appear; the BIND installer needs administrator rights)'
    $code = Invoke-Native -File 'winget.exe' -Arguments $wingetArgs
    return ($code -eq 0)
}

function Ensure-Dig {
    Write-Step 'Step 3/4  dig.exe (ISC BIND)'
    # Let the child process see a project-local portable BIND copy.
    $portableBind = Join-Path $ProjectRoot 'tools\bind'
    if (Test-Path $portableBind) {
        if ($env:PATH -notlike "*$portableBind*") { $env:PATH = "$portableBind;$env:PATH" }
    }
    $dig = Find-Dig
    if ($dig) { Write-Ok "found: $dig"; return $dig }
    Write-Warn2 'dig.exe not found'

    if ($script:NeedInstall) {
        Write-Host '    installing BIND tools (dig) via winget ...'
        try { [void](Install-DigViaWinget) } catch { Write-Warn2 ("winget failed: " + $_.Exception.Message) }
        $dig = Find-Dig
        if ($dig) { Write-Ok "installed: $dig"; return $dig }
        Write-Warn2 'dig is still missing after the installation attempt'
    } elseif ($DryRun) {
        Write-Host '    (dry-run) would install BIND (dig) via winget'
    }

    Write-Warn2 'you can also copy a working BIND folder to <project>\tools\bind\'
    return $null
}

# ---------------------------------------------------------------- main flow

Write-Host '============================================================'
Write-Host ' Africa_DN_Testing - environment check / auto setup'
Write-Host '============================================================'
Write-Host (" project : $ProjectRoot")
if ($DryRun) { Write-Host ' mode    : DRY-RUN (nothing will be installed)' }
if ($NoInstall) { Write-Host ' mode    : NoInstall (check only)' }
if ($FakeMissingPython -or $FakeMissingDeps -or $FakeMissingDig) {
    $fake = @()
    if ($FakeMissingPython) { $fake += 'python' }
    if ($FakeMissingDeps) { $fake += 'deps' }
    if ($FakeMissingDig) { $fake += 'dig' }
    Write-Host (" simulate: pretending " + ($fake -join ', ') + ' are missing')
}

$python = Ensure-Python
if (-not $python) {
    if ($DryRun) {
        Write-Host ''
        Write-Host ' [dry-run] plan for a machine WITHOUT Python:' -ForegroundColor Cyan
        Write-Host '   1) install Python : winget install --id Python.Python.3.13 --exact --scope user --silent'
        Write-Host '      (fallback)     : https://www.python.org/ftp/python/<ver>/python-<ver>-amd64.exe'
        Write-Host '                       /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1'
        Write-Host '   2) create venv    : python -m venv .venv'
        Write-Host '   3) install deps   : .venv\Scripts\python.exe -m pip install -r requirements.txt'
        Write-Host '   4) install dig    : winget install --id ISC.Bind --exact --silent'
        Write-Host '   5) start the test : run_dns_test.bat'
        Write-Host ''
        exit 0
    }
    Write-Host ''
    Write-Err2 'environment is NOT ready (Python)'
    exit 2
}

$venvPython = Get-VenvPython $python
if ($venvPython) { $python = $venvPython }
Write-Host ""
Write-Host (" [i] interpreter in use: $python")

$depsOk = Ensure-Dependencies $python
$dig = Ensure-Dig

Write-Step 'Step 4/4  summary'
if ($depsOk) { Write-Ok 'dependencies ready' } else { Write-Err2 'dependencies NOT ready' }
if ($dig) { Write-Ok "dig ready: $dig" } else { Write-Err2 'dig NOT ready' }

# The cache is written only when this run really provisioned the machine and
# the result is usable:
#   -DryRun / -NoInstall must not touch the disk at all, and a dependency
#   failure must not leave a record that run_dns_test.bat would reject anyway
#   (it re-checks "interpreter exists + can import openpyxl, requests").
if ($script:NeedInstall -and $depsOk) {
    if (-not (Test-Path $CacheDir)) { New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null }
    # 'Default' = the ANSI code page on Windows PowerShell 5.1, which is exactly
    # what cmd.exe expects when run_dns_test.bat reads this file.
    # (UTF-8 or ASCII would corrupt non-ASCII characters in the project path.)
    #
    # The file is deliberately MACHINE BOUND: run_dns_test.bat only reuses the
    # interpreter when the recorded computer name matches this computer, so a
    # cache folder copied from another PC (with that PC's absolute paths) is
    # ignored instead of being used. Keep both comment lines free of '=' and
    # keep the two 'key=value' lines as they are.
    $machineName = $env:COMPUTERNAME
    if ([string]::IsNullOrWhiteSpace($machineName)) { $machineName = [System.Net.Dns]::GetHostName() }
    $cacheLines = @(
        '# Africa_DN_Testing interpreter cache - generated by scripts\setup_env.ps1',
        '# machine specific: do not copy this cache folder to another computer',
        ('machine=' + $machineName),
        ('python=' + $python)
    )
    Set-Content -Path $PythonPathFile -Value $cacheLines -Encoding Default
    Write-Host ""
    Write-Host (" [i] interpreter path written to: $PythonPathFile")
    Write-Host ("     (machine: $machineName)")
}

Write-Host ""
if (-not $depsOk) {
    if ($DryRun) { exit 0 }   # a simulated/planned run is not a failure
    exit 3
}
if ($dig) {
    Write-Host ' [DONE] environment ready - starting the test program ...' -ForegroundColor Green
} else {
    Write-Host ' [DONE] Python and dependencies are ready, but dig is missing.'
    Write-Host '        domain_dns_test.py will stop and explain how to get dig.exe.'
}
exit 0
