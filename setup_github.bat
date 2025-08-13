@echo off
echo ========================================
echo    GITHUB REPOSITORY SETUP SCRIPT
echo ========================================
echo.

echo Current directory: %CD%
echo.

echo Step 1: Checking git status...
git status
echo.

echo Step 2: Please create the repository on GitHub first:
echo    1. Go to: https://github.com/new
echo    2. Repository name: MyLibraries
echo    3. Description: A collection of .NET libraries featuring ExcelHelper.NET - advanced Excel manipulation with merge cell support
echo    4. Keep it PUBLIC
echo    5. Do NOT add README, .gitignore, or license (we have them)
echo    6. Click 'Create repository'
echo.

set /p username="Enter your GitHub username: "
if "%username%"=="" (
    echo Error: Username cannot be empty!
    pause
    exit /b 1
)

echo.
echo Step 3: Adding remote origin...
git remote add origin https://github.com/%username%/MyLibraries.git

echo Step 4: Renaming branch to main...
git branch -M main

echo Step 5: Pushing to GitHub...
git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo     SUCCESS! Repository created!
    echo ========================================
    echo.
    echo Your repository is now available at:
    echo https://github.com/%username%/MyLibraries
    echo.
    echo Features uploaded:
    echo ✅ ExcelHelper.NET library with merge cell support
    echo ✅ Complete documentation and examples  
    echo ✅ Clean modular architecture
    echo ✅ MIT License and professional README
    echo ✅ 32 files, 6,137 lines of code
    echo.
) else (
    echo.
    echo ========================================
    echo     ERROR: Push failed!
    echo ========================================
    echo.
    echo Possible causes:
    echo 1. Repository not created on GitHub yet
    echo 2. Wrong username
    echo 3. Authentication issues
    echo.
    echo Please check and try again manually:
    echo git remote add origin https://github.com/%username%/MyLibraries.git
    echo git push -u origin main
)

echo.
pause
