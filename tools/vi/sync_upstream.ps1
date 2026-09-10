<#
.SYNOPSIS
  (Chủ repo) Đồng bộ một tag của upstream hugohe3/ppt-master vào nhánh hiện tại.
.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File tools\vi\sync_upstream.ps1 -Tag v6.4.0 -Python C:\venv\Scripts\python.exe
#>
param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [string]$Python = 'python',
    [switch]$SkipSmoke
)

$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }
$env:PYTHONIOENCODING = 'utf-8'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$IndexKinds = @{
    'skills/ppt-master/templates/decks/decks_index.json'     = 'deck'
    'skills/ppt-master/templates/brands/brands_index.json'   = 'brand'
    'skills/ppt-master/templates/layouts/layouts_index.json' = 'layout'
    'skills/ppt-master/templates/styles/styles_index.json'   = 'style'
}

function Stop-Sync([string]$Message) {
    Write-Host "[DỪNG] $Message" -ForegroundColor Red
    Pop-Location
    exit 1
}

Push-Location $RepoRoot

if (-not (Get-Command git -ErrorAction SilentlyContinue)) { Stop-Sync 'Không tìm thấy git trong PATH.' }
if (-not (Get-Command $Python -ErrorAction SilentlyContinue)) { Stop-Sync "Không tìm thấy Python: $Python" }

$dirty = git status --porcelain --untracked-files=no
if ($dirty) { Stop-Sync 'Cây làm việc còn thay đổi chưa commit.' }

Write-Host "==> Lấy tag từ upstream" -ForegroundColor Cyan
git fetch upstream --tags | Out-Host
if ($LASTEXITCODE -ne 0) { Stop-Sync 'Không fetch được upstream. Chạy: git remote add upstream https://github.com/hugohe3/ppt-master.git' }
git rev-parse -q --verify "refs/tags/$Tag" | Out-Null
if ($LASTEXITCODE -ne 0) { Stop-Sync "Không có tag $Tag." }

git config merge.ours.driver true
if ($LASTEXITCODE -ne 0) { Stop-Sync 'Không bật được merge.ours.driver (git config thất bại).' }

Write-Host "==> Merge $Tag" -ForegroundColor Cyan
git merge --no-ff --no-edit -m "merge: sync upstream hugohe3/ppt-master $Tag" $Tag | Out-Host
if ($LASTEXITCODE -ne 0) {
    $conflicts = @(git diff --name-only --diff-filter=U)
    if ($conflicts.Count -eq 0) { Stop-Sync 'git merge thất bại nhưng không có xung đột (xem thông báo phía trên).' }
    $unresolved = @()
    foreach ($path in $conflicts) {
        if ($IndexKinds.ContainsKey($path)) {
            git checkout --theirs -- $path | Out-Host
            if ($LASTEXITCODE -ne 0) { $unresolved += $path; continue }
            & $Python 'skills/ppt-master/scripts/register_template.py' --rebuild-all --kind $IndexKinds[$path] | Out-Host
            if ($LASTEXITCODE -ne 0) { $unresolved += $path; continue }
            git add -- $path | Out-Host
        } else {
            $unresolved += $path
        }
    }
    if ($unresolved.Count -gt 0) {
        Write-Host 'Các file cần xử lý tay:' -ForegroundColor Yellow
        foreach ($path in $unresolved) { Write-Host "  - $path" }
        Stop-Sync 'Merge chưa hoàn tất. Sửa các file trên, git add, rồi git commit.'
    }
    git commit --no-edit | Out-Host
    if ($LASTEXITCODE -ne 0) { Stop-Sync 'Không commit được merge. Merge đang dang dở: sửa rồi git commit, hoặc hủy bằng: git merge --abort' }
}

$UndoHint = 'Merge đã được commit nhưng CHƯA push. Sửa lớp Việt rồi commit tiếp, hoặc hoàn tác merge bằng: git reset --keep ORIG_HEAD'

Write-Host '==> Kiểm tra toàn vẹn skill' -ForegroundColor Cyan
& $Python 'skills/ppt-master/scripts/attribution_guard.py' | Out-Host
if ($LASTEXITCODE -ne 0) { Stop-Sync "attribution_guard.py thất bại sau khi đồng bộ. $UndoHint" }

Write-Host '==> Test lớp Việt hoá' -ForegroundColor Cyan
& $Python -m unittest discover -s tools/vi/tests | Out-Host
if ($LASTEXITCODE -ne 0) { Stop-Sync "Test lớp Việt hoá thất bại. Cập nhật lớp Việt cho khớp upstream mới rồi commit. $UndoHint" }

Write-Host '==> Kiểm tra môi trường' -ForegroundColor Cyan
$doctorArgs = @()
if ($SkipSmoke) { $doctorArgs = @('--no-smoke') }
& $Python 'tools/vi/doctor.py' @doctorArgs | Out-Host
if ($LASTEXITCODE -ne 0) { Stop-Sync "doctor.py còn lỗi bắt buộc. $UndoHint" }

Write-Host "[XONG] Đã đồng bộ $Tag. Việc còn lại: cập nhật CHANGELOG-VI.md, gắn tag phiên bản, push, tạo Release." -ForegroundColor Green
Pop-Location
exit 0
