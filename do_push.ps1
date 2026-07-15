# Push helper - token dibaca dari env var, tidak hardcode
$GIT      = "C:\Program Files\Git\cmd\git.exe"
$GH_USER  = "facilitycareaudydental-sketch"
$REPO     = "facilitycare-system"
$EMAIL    = "facilitycare.audydental@gmail.com"
$NAME     = "Facility Care Audy Dental"
$TOKEN    = $env:GH_TOKEN
$REMOTE   = "https://${TOKEN}@github.com/${GH_USER}/${REPO}.git"

Write-Host "=== Git config ===" -ForegroundColor Yellow
& $GIT config user.email $EMAIL
& $GIT config user.name $NAME
Write-Host "=== Remote ===" -ForegroundColor Yellow
& $GIT remote remove origin 2>$null
& $GIT remote add origin $REMOTE
Write-Host "=== Add ===" -ForegroundColor Yellow
& $GIT add --all
Write-Host "=== Commit ===" -ForegroundColor Yellow
& $GIT commit -m "feat: FM Operations Management System - complete build"
Write-Host "=== Branch main ===" -ForegroundColor Yellow
& $GIT branch -M main
Write-Host "=== Push ===" -ForegroundColor Yellow
& $GIT push -u origin main --force 2>&1
Write-Host "Push code: $LASTEXITCODE"
