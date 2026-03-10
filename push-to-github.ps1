# Run this in PowerShell from the project folder (where Git is installed).
# Adds the GitHub remote and pushes your changes.

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "Git is not installed or not in PATH. Install from https://git-scm.com/" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path .git)) {
    Write-Host "Initializing git repository..."
    git init
    git branch -M main
}

git remote remove origin 2>$null
git remote add origin https://github.com/Brian-K05/a4-word-barcode-generator.git
Write-Host "Remote 'origin' set to: https://github.com/Brian-K05/a4-word-barcode-generator.git" -ForegroundColor Green

git add .
git status
# Commit if there are changes (new repo or uncommitted files)
$status = git status --porcelain
if ($status) {
    git commit -m "Add ZIP download with Word + matching barcode image"
}
Write-Host "Pushing to origin main..."
git push -u origin main
