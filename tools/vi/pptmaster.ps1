<#
.SYNOPSIS
  Trình khởi chạy bản Việt của PPT Master: cài đặt, kiểm tra, cập nhật.
.PARAMETER Action
  setup  - kiểm tra Python, cài thư viện, tạo .env, công cụ tuỳ chọn, chạy doctor
  check  - chạy doctor
  update - cập nhật bằng git (update_repo.py của upstream) rồi kiểm tra nhanh
.PARAMETER NonInteractive
  Không hỏi Y/N và không tự cài phần mềm (dùng khi kiểm thử).
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('setup', 'check', 'update')]
    [string]$Action,
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$env:PYTHONIOENCODING = 'utf-8'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$Doctor = Join-Path $RepoRoot 'tools\vi\doctor.py'
$FixDoc = Join-Path $RepoRoot 'docs\vi\xu-ly-loi.md'
$InstallDoc = Join-Path $RepoRoot 'docs\vi\cai-dat-windows.md'

function Write-Step([string]$Text) { Write-Host ''; Write-Host "==> $Text" -ForegroundColor Cyan }
function Write-Fail([string]$Text) { Write-Host "[LỖI] $Text" -ForegroundColor Red }
function Write-Ok([string]$Text) { Write-Host "[OK] $Text" -ForegroundColor Green }

function Test-Winget { return [bool](Get-Command winget -ErrorAction SilentlyContinue) }

function Confirm-Choice([string]$Question) {
    if ($NonInteractive) { return $false }
    $answer = Read-Host "$Question (Y/N)"
    return ($answer -match '^(y|yes|c|co|có)$')
}

function Get-PythonInfo {
    $cmd = Get-Command python -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $cmd) { return $null }
    $version = $null
    try {
        $raw = & $cmd.Source -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -eq 0 -and $raw) { $version = [version]("$raw".Trim()) }
    } catch { $version = $null }
    return [pscustomobject]@{
        Path         = $cmd.Source
        Version      = $version
        IsStoreAlias = ($cmd.Source -like '*\WindowsApps\*')
    }
}

function Get-LauncherPython {
    $launcher = Get-Command py -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $launcher) { return $null }
    try {
        $raw = @(& $launcher.Source -3 -c "import sys; print(sys.executable); print('%d.%d' % sys.version_info[:2])" 2>$null)
        if ($LASTEXITCODE -ne 0 -or $raw.Count -lt 2) { return $null }
        $version = [version]("$($raw[$raw.Count - 1])".Trim())
    } catch { return $null }
    if ($version -lt [version]'3.10') { return $null }
    return [pscustomobject]@{
        Path    = "$($raw[$raw.Count - 2])".Trim()
        Version = $version
    }
}

function Resolve-Python([bool]$OfferInstall) {
    $py = Get-PythonInfo
    if ($py -and $py.Version -and $py.Version -ge [version]'3.10') {
        Write-Ok "Python $($py.Version) tại $($py.Path)"
        return $py
    }
    $installed = Get-LauncherPython
    if ($installed) {
        Write-Fail "Đã cài Python $($installed.Version) tại $($installed.Path) nhưng bản này chưa có trong PATH."
        Write-Host 'Cách sửa: Settings → Apps → Installed apps → Python 3.x → Modify → tick "Add Python to environment variables" (hoặc chạy lại bộ cài tải từ python.org và chọn Modify).'
        Write-Host 'Sau đó đóng cửa sổ này, mở lại rồi bấm lại CAI-DAT.bat.'
        Write-Host "Hướng dẫn chi tiết: $FixDoc (mục Đã cài Python nhưng bộ cài báo không tìm thấy)"
        return $null
    }
    if (-not $py) {
        Write-Fail 'Chưa tìm thấy Python trong PATH.'
    } elseif ((-not $py.Version) -and $py.IsStoreAlias) {
        Write-Fail 'Lệnh python đang trỏ tới lối tắt của Microsoft Store, chưa phải Python thật.'
        Write-Host 'Tắt lối tắt: Settings → Apps → Advanced app settings → App execution aliases → tắt "python.exe" và "python3.exe".'
    } elseif (-not $py.Version) {
        Write-Fail "Không chạy được Python tại $($py.Path)."
    } else {
        Write-Fail "Python $($py.Version) quá cũ, cần 3.10 trở lên."
    }
    if ($OfferInstall -and (Test-Winget) -and (Confirm-Choice 'Cài Python 3.12 bằng winget?')) {
        winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements | Out-Host
        Write-Host ''
        Write-Host 'Đã chạy trình cài Python. Hãy ĐÓNG cửa sổ này rồi bấm lại CAI-DAT.bat để PATH mới có hiệu lực.' -ForegroundColor Yellow
    } else {
        Write-Host 'Cài Python 3.12 tại https://www.python.org/downloads/ và tick "Add python.exe to PATH".'
        Write-Host "Hướng dẫn chi tiết: $InstallDoc"
    }
    return $null
}

function Invoke-Doctor($Py, [string[]]$DoctorArgs) {
    & $Py.Path $Doctor @DoctorArgs | Out-Host
    return $LASTEXITCODE
}

function Invoke-Setup {
    Write-Step 'Bước 1/4: Kiểm tra Python'
    $py = Resolve-Python $true
    if (-not $py) { return 1 }

    Write-Step 'Bước 2/4: Cài thư viện Python (lần đầu có thể mất vài phút)'
    & $py.Path -m pip install --upgrade pip | Out-Host
    & $py.Path -m pip install -r (Join-Path $RepoRoot 'requirements.txt') | Out-Host
    if ($LASTEXITCODE -ne 0) {
        Write-Fail 'Cài thư viện thất bại. Xem thông báo phía trên.'
        Write-Host "Cách xử lý: $FixDoc (mục Cài thư viện thất bại)"
        return 1
    }
    Write-Ok 'Đã cài thư viện.'

    Write-Step 'Bước 3/4: Tạo file cấu hình .env'
    $envFile = Join-Path $RepoRoot '.env'
    if (Test-Path $envFile) {
        Write-Ok 'Đã có .env, giữ nguyên.'
    } else {
        Copy-Item (Join-Path $RepoRoot '.env.example') $envFile
        Write-Ok 'Đã tạo .env từ .env.example.'
    }

    Write-Step 'Bước 4/4: Công cụ tuỳ chọn'
    $optional = @(
        @{ Command = 'git'; Id = 'Git.Git'; Label = 'Git (để cập nhật bằng CAP-NHAT.bat)' },
        @{ Command = 'pandoc'; Id = 'JohnMacFarlane.Pandoc'; Label = 'Pandoc (chuyển tài liệu định dạng cũ)' },
        @{ Command = 'ffmpeg'; Id = 'Gyan.FFmpeg'; Label = 'FFmpeg (thuyết minh, video)' }
    )
    $installedAny = $false
    foreach ($tool in $optional) {
        if (Get-Command $tool.Command -ErrorAction SilentlyContinue) {
            Write-Ok "$($tool.Label): đã có"
            continue
        }
        if ((Test-Winget) -and (Confirm-Choice "Cài $($tool.Label) bằng winget?")) {
            winget install -e --id $tool.Id --accept-package-agreements --accept-source-agreements | Out-Host
            $installedAny = $true
        } else {
            Write-Host "Bỏ qua: $($tool.Label)"
        }
    }
    if ($installedAny) {
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
    }

    Write-Step 'Kiểm tra lại toàn bộ'
    return (Invoke-Doctor $py @())
}

function Invoke-Check {
    $py = Resolve-Python $false
    if (-not $py) { Write-Host 'Hãy bấm CAI-DAT.bat trước.'; return 1 }
    return (Invoke-Doctor $py @())
}

function Invoke-Update {
    if (-not (Test-Path (Join-Path $RepoRoot '.git'))) {
        Write-Fail 'Thư mục này được tải dạng ZIP nên không tự cập nhật được.'
        Write-Host 'Tải bản mới tại https://github.com/luonghaianh1208/PPTmaster rồi chép thư mục projects\ và file .env của bạn sang.'
        return 1
    }
    $py = Resolve-Python $false
    if (-not $py) { Write-Host 'Hãy bấm CAI-DAT.bat trước.'; return 1 }
    Write-Step 'Tải bản mới nhất'
    & $py.Path (Join-Path $RepoRoot 'skills\ppt-master\scripts\update_repo.py') | Out-Host
    if ($LASTEXITCODE -ne 0) {
        Write-Fail 'Cập nhật thất bại.'
        Write-Host "Nếu thông báo có 'Tracked local changes' là bạn đã sửa file của bộ công cụ. Cách xử lý: $FixDoc (mục Cập nhật thất bại)"
        return 1
    }
    Write-Step 'Kiểm tra nhanh sau cập nhật'
    return (Invoke-Doctor $py @('--no-smoke'))
}

Push-Location $RepoRoot
try {
    switch ($Action) {
        'setup' { $code = Invoke-Setup }
        'check' { $code = Invoke-Check }
        'update' { $code = Invoke-Update }
    }
} finally {
    Pop-Location
}
exit ([int]($code | Select-Object -Last 1))
