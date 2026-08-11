param(
    [string]$InputHtml = "",
    [string]$OutputPptx = ""
)

$ErrorActionPreference = "Stop"
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Add-Type -AssemblyName System.Windows.Forms

function Show-Error([string]$message) {
    [System.Windows.Forms.MessageBox]::Show(
        $message,
        "HTML Slides to PPT",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

try {
    $nodeCommand = Get-Command node -ErrorAction Stop
} catch {
    Show-Error "Node.js was not found. Install Node.js 20 or later, then run setup-windows.cmd."
    exit 1
}

$playwrightPackage = Join-Path $appDir "node_modules\playwright\package.json"
$pptxgenPackage = Join-Path $appDir "node_modules\pptxgenjs\package.json"
if (-not (Test-Path -LiteralPath $playwrightPackage -PathType Leaf) -or
    -not (Test-Path -LiteralPath $pptxgenPackage -PathType Leaf)) {
    Write-Host "Required packages are missing or incomplete. Installing them now..." -ForegroundColor Yellow
    try {
        $npmCommand = Get-Command npm.cmd -ErrorAction Stop
        Push-Location $appDir
        try {
            & $npmCommand.Source install
            $npmExitCode = $LASTEXITCODE
        } finally {
            Pop-Location
        }
    } catch {
        Show-Error "npm was not found. Reinstall Node.js, then run setup-windows.cmd."
        exit 1
    }
    if ($npmExitCode -ne 0 -or
        -not (Test-Path -LiteralPath $playwrightPackage -PathType Leaf) -or
        -not (Test-Path -LiteralPath $pptxgenPackage -PathType Leaf)) {
        Show-Error "Dependency installation failed. Check your network, then run setup-windows.cmd."
        exit 1
    }
}

if ([string]::IsNullOrWhiteSpace($InputHtml)) {
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Title = "Select HTML Slides"
    $openDialog.Filter = "HTML files (*.html)|*.html"
    $openDialog.Multiselect = $false
    if ($openDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        exit 0
    }
    $InputHtml = $openDialog.FileName
}

$InputHtml = [System.IO.Path]::GetFullPath($InputHtml)
if (-not (Test-Path -LiteralPath $InputHtml -PathType Leaf)) {
    Show-Error "The input file does not exist:`n$InputHtml"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($OutputPptx)) {
    $defaultName = [System.IO.Path]::GetFileNameWithoutExtension($InputHtml) + "-objects.pptx"
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Title = "Save PowerPoint"
    $saveDialog.Filter = "PowerPoint files (*.pptx)|*.pptx"
    $saveDialog.FileName = $defaultName
    $saveDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($InputHtml)
    if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        exit 0
    }
    $OutputPptx = $saveDialog.FileName
}

$converter = Join-Path $appDir "convert.mjs"
& $nodeCommand.Source $converter --input $InputHtml --output $OutputPptx --mode objects
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "Conversion completed:`n$OutputPptx",
        "HTML Slides to PPT",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    Start-Process explorer.exe -ArgumentList "/select,`"$OutputPptx`""
} else {
    Show-Error "Conversion failed. Review the error shown in the command window."
}

exit $exitCode
