Param(
    [string]$ProjectFolder = $PSScriptRoot,
    [string]$Configuration = "Release",
    [string]$SolutionPath = "",
    [string]$MsbuildPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe",
    [string]$ReferencePath = "",
    [string]$OutputDir = "",
    [switch]$SkipIceValidation,
    [string]$SigningPfx = $env:IBP_SIGNING_PFX,
    [string]$SigningPassword = $env:IBP_SIGNING_PASSWORD,
    [string]$TimestampUrl = "http://timestamp.digicert.com",
    [switch]$SkipSigning
)

$ErrorActionPreference = "Stop"

function Find-SignTool {
    $windowsKitsBin = "C:\Program Files (x86)\Windows Kits\10\bin"
    if (-not (Test-Path $windowsKitsBin)) {
        return $null
    }

    return Get-ChildItem $windowsKitsBin -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '\\x64\\signtool\.exe$' } |
        Sort-Object {
            try {
                [version]$_.Directory.Parent.Name
            }
            catch {
                [version]"0.0"
            }
        } -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}

function Sign-IBPFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        throw "File to sign not found: $Path"
    }

    Write-Host "Signing: $Path"
    & $script:SignToolPath sign `
        /f $SigningPfx `
        /p $SigningPassword `
        /fd SHA256 `
        /tr $TimestampUrl `
        /td SHA256 `
        $Path | Out-Host

    if ($LASTEXITCODE -ne 0) {
        throw "SignTool failed while signing '$Path' with exit code $LASTEXITCODE."
    }

    Write-Host "Verifying signature: $Path"
    & $script:SignToolPath verify /pa /v $Path | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Signature verification failed for '$Path' with exit code $LASTEXITCODE."
    }
}

$ProjectFolder = (Resolve-Path $ProjectFolder).Path
if ([string]::IsNullOrWhiteSpace($SolutionPath)) {
    $SolutionPath = Join-Path $ProjectFolder "NcTalkOutlookAddIn.sln"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $ProjectFolder "dist"
}

$signingEnabled = -not $SkipSigning
if ($signingEnabled) {
    if ([string]::IsNullOrWhiteSpace($SigningPfx)) {
        Write-Warning "Signing skipped: no PFX configured. Set IBP_SIGNING_PFX or pass -SigningPfx."
        $signingEnabled = $false
    }
    elseif (-not (Test-Path $SigningPfx)) {
        throw "Signing PFX not found at $SigningPfx"
    }
    elseif ([string]::IsNullOrWhiteSpace($SigningPassword)) {
        throw "Signing password missing. Set IBP_SIGNING_PASSWORD or pass -SigningPassword."
    }
    else {
        $script:SignToolPath = Find-SignTool
        if ([string]::IsNullOrWhiteSpace($script:SignToolPath)) {
            throw "signtool.exe not found. Install the Windows 10/11 SDK."
        }
        Write-Host "Signing enabled. SignTool: $script:SignToolPath"
    }
}

if (-not (Test-Path $MsbuildPath)) {
    throw "MSBuild.exe not found at $MsbuildPath"
}
if (-not (Test-Path $SolutionPath)) {
    throw "Solution not found at $SolutionPath"
}

Write-Host "Building solution using $MsbuildPath ($Configuration)..."
$msbuildArgs = @(
    $SolutionPath,
    "/m",
    "/t:Rebuild",
    "/p:Configuration=$Configuration"
)
if (-not [string]::IsNullOrWhiteSpace($ReferencePath)) {
    $msbuildArgs += "/p:ReferencePath=$ReferencePath"
}
& $MsbuildPath @msbuildArgs | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild exited with code $LASTEXITCODE."
}

$buildOutputDir = Join-Path $ProjectFolder "src\NcTalkOutlookAddIn\bin\$Configuration"
$dllPath = Join-Path $buildOutputDir "NcTalkOutlookAddIn.dll"
if (-not (Test-Path $dllPath)) {
    throw "Build succeeded but assembly not found at $dllPath."
}

$assemblyInfo = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath).Version
$assemblyVersionFull = $assemblyInfo.ToString()
$assemblyVersionShort = "{0}.{1}.{2}" -f $assemblyInfo.Major, $assemblyInfo.Minor, $assemblyInfo.Build
Write-Host "Assembly version detected: $assemblyVersionFull"

# Sign our own add-in assembly before WiX packages it into the MSI.
# Third-party dependency DLLs are deliberately left untouched.
if ($signingEnabled) {
    Sign-IBPFile $dllPath
}

$wixProject = Join-Path $ProjectFolder "installer\NcConnectorOutlookInstaller.wixproj"
if (-not (Test-Path $wixProject)) {
    throw "WiX project not found at $wixProject."
}

Write-Host "Building MSI via WiX v6 SDK (dotnet build)..."
$wixArgs = @(
    "build",
    $wixProject,
    "-c",
    $Configuration,
    "--no-incremental",
    "/p:BuildOutputDir=$buildOutputDir\",
    "/p:ProductVersion=$assemblyVersionShort",
    "/p:AssemblyVersion=$assemblyVersionFull"
)
if ($SkipIceValidation) {
    # Useful on environments where Windows Installer ICE execution is unavailable (WIX0217).
    $wixArgs += "/p:SuppressValidation=true"
}
& dotnet @wixArgs | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "dotnet build (WiX) exited with code $LASTEXITCODE."
}

$builtMsiPath = Join-Path $ProjectFolder "installer\bin\$Configuration\NCConnectorForOutlook.msi"
if (-not (Test-Path $builtMsiPath)) {
    throw "MSI build succeeded but output not found at $builtMsiPath."
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$finalName = "IBP-NC-ConnectorForOutlook$assemblyVersionShort.msi"
$finalPath = Join-Path $OutputDir $finalName
Copy-Item -Force $builtMsiPath $finalPath

# Sign the final MSI after all packaging/copy operations are complete.
if ($signingEnabled) {
    Sign-IBPFile $finalPath
}

Write-Host "MSI erstellt: $finalPath"
if ($signingEnabled) {
    Write-Host "Signierung abgeschlossen: Add-in DLL und MSI sind mit Zeitstempel signiert."
}
