# GitHub Repository Setup Script for MyLibraries
# PowerShell version

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    GITHUB REPOSITORY SETUP SCRIPT" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Current directory: $(Get-Location)" -ForegroundColor Yellow
Write-Host ""

Write-Host "Step 1: Checking git status..." -ForegroundColor Green
git status
Write-Host ""

Write-Host "Step 2: Please create the repository on GitHub first:" -ForegroundColor Green
Write-Host "   1. Go to: https://github.com/new" -ForegroundColor White
Write-Host "   2. Repository name: MyLibraries" -ForegroundColor White
Write-Host "   3. Description: A collection of .NET libraries featuring ExcelHelper.NET - advanced Excel manipulation with merge cell support" -ForegroundColor White
Write-Host "   4. Keep it PUBLIC" -ForegroundColor White
Write-Host "   5. Do NOT add README, .gitignore, or license (we have them)" -ForegroundColor White
Write-Host "   6. Click 'Create repository'" -ForegroundColor White
Write-Host ""

$username = Read-Host "Enter your GitHub username"
if ([string]::IsNullOrWhiteSpace($username)) {
    Write-Host "Error: Username cannot be empty!" -ForegroundColor Red
    Read-Host "Press any key to exit"
    exit 1
}

Write-Host ""
Write-Host "Step 3: Adding remote origin..." -ForegroundColor Green
git remote add origin "https://github.com/$username/MyLibraries.git"

Write-Host "Step 4: Renaming branch to main..." -ForegroundColor Green
git branch -M main

Write-Host "Step 5: Pushing to GitHub..." -ForegroundColor Green
$pushResult = git push -u origin main 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "     SUCCESS! Repository created!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Your repository is now available at:" -ForegroundColor Yellow
    Write-Host "https://github.com/$username/MyLibraries" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Features uploaded:" -ForegroundColor Yellow
    Write-Host "✅ ExcelHelper.NET library with merge cell support" -ForegroundColor Green
    Write-Host "✅ Complete documentation and examples" -ForegroundColor Green
    Write-Host "✅ Clean modular architecture" -ForegroundColor Green
    Write-Host "✅ MIT License and professional README" -ForegroundColor Green
    Write-Host "✅ 32 files, 6,137 lines of code" -ForegroundColor Green
    Write-Host ""
    
    # Open repository in browser
    $openBrowser = Read-Host "Open repository in browser? (y/n)"
    if ($openBrowser -eq "y" -or $openBrowser -eq "Y") {
        Start-Process "https://github.com/$username/MyLibraries"
    }
} else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "     ERROR: Push failed!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error details:" -ForegroundColor Red
    Write-Host $pushResult -ForegroundColor Red
    Write-Host ""
    Write-Host "Possible causes:" -ForegroundColor Yellow
    Write-Host "1. Repository not created on GitHub yet" -ForegroundColor White
    Write-Host "2. Wrong username" -ForegroundColor White
    Write-Host "3. Authentication issues (try GitHub Desktop or personal access token)" -ForegroundColor White
    Write-Host ""
    Write-Host "Manual commands to try:" -ForegroundColor Yellow
    Write-Host "git remote add origin https://github.com/$username/MyLibraries.git" -ForegroundColor Cyan
    Write-Host "git push -u origin main" -ForegroundColor Cyan
}

Write-Host ""
Read-Host "Press any key to exit"
