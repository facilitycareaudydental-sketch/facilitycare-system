$ErrorActionPreference = "Continue"
$GIT = "C:\Program Files\Git\cmd\git.exe"
$GH_USER = "facilitycareaudydental-sketch"
$REPO = "facilitycare-system"
$EMAIL = "facilitycare.audydental@gmail.com"
$NAME = "Facility Care Audy Dental"
$TOKEN = $env:GH_TOKEN
$REMOTE = "https://${TOKEN}@github.com/${GH_USER}/${REPO}.git"

Write-Host "[1] Reset history - buat commit bersih baru" -ForegroundColor Yellow
# Reset ke single clean commit (orphan branch)
& $GIT checkout --orphan clean_main 2>&1
& $GIT add --all 2>&1
& $GIT config user.email $EMAIL
& $GIT config user.name $NAME
& $GIT commit -m "feat: FM Operations Management System - complete build" 2>&1
Write-Host "Commit: $LASTEXITCODE"

Write-Host "[2] Hapus branch main lama" -ForegroundColor Yellow
& $GIT branch -D main 2>&1
& $GIT branch -M clean_main main 2>&1
Write-Host "Branch OK"

Write-Host "[3] Setup remote" -ForegroundColor Yellow
& $GIT remote remove origin 2>$null
& $GIT remote add origin $REMOTE
Write-Host "Remote OK"

Write-Host "[4] Push ke GitHub..." -ForegroundColor Yellow
& $GIT push -u origin main --force 2>&1
$code = $LASTEXITCODE
Write-Host "Push code: $code"
if ($code -eq 0) {
    Write-Host "BERHASIL: https://github.com/$GH_USER/$REPO" -ForegroundColor Green
} else {
    Write-Host "GAGAL code $code" -ForegroundColor Red
}
