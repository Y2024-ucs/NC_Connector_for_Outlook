Param(
    [Parameter(Mandatory = $true)][string]$MsiPath,
    [Parameter(Mandatory = $true)][string]$TransformPath
)

# Creates an MSI transform (.mst) that sets ALLUSERS=1, so a GPO assignment installs the package for
# all users of the machine (group policy cannot pass command-line properties). The MSI stays unchanged.

$ErrorActionPreference = "Stop"
$MsiPath = (Resolve-Path $MsiPath).Path
$TransformPath = [IO.Path]::GetFullPath($TransformPath)

function Invoke-Com($target, [string]$member, [string]$kind, [object[]]$arguments) {
    # PowerShell wraps COM objects; Windows Installer's IDispatch needs the raw objects.
    $raw = if ($null -eq $arguments) { $null } else { [object[]]@($arguments | ForEach-Object { if ($_ -is [psobject]) { $_.psobject.BaseObject } else { $_ } }) }
    $self = if ($target -is [psobject]) { $target.psobject.BaseObject } else { $target }
    return $self.GetType().InvokeMember($member, $kind, $null, $self, $raw)
}

$installer = New-Object -ComObject WindowsInstaller.Installer
$workCopy = Join-Path ([IO.Path]::GetTempPath()) ("nc4ol-mst-" + [Guid]::NewGuid().ToString("N") + ".msi")
Copy-Item -LiteralPath $MsiPath -Destination $workCopy
try {
    $msiOpenDatabaseModeReadOnly = 0
    $msiOpenDatabaseModeTransact = 1
    $changed = Invoke-Com $installer "OpenDatabase" "InvokeMethod" @($workCopy, $msiOpenDatabaseModeTransact)
    $query = Invoke-Com $changed "OpenView" "InvokeMethod" @("SELECT ``Value`` FROM ``Property`` WHERE ``Property`` = 'ALLUSERS'")
    Invoke-Com $query "Execute" "InvokeMethod" $null | Out-Null
    $existing = Invoke-Com $query "Fetch" "InvokeMethod" $null
    Invoke-Com $query "Close" "InvokeMethod" $null | Out-Null
    $sql = if ($existing) {
        "UPDATE ``Property`` SET ``Value`` = '1' WHERE ``Property`` = 'ALLUSERS'"
    } else {
        "INSERT INTO ``Property`` (``Property``, ``Value``) VALUES ('ALLUSERS', '1')"
    }
    $update = Invoke-Com $changed "OpenView" "InvokeMethod" @($sql)
    Invoke-Com $update "Execute" "InvokeMethod" $null | Out-Null
    Invoke-Com $update "Close" "InvokeMethod" $null | Out-Null
    Invoke-Com $changed "Commit" "InvokeMethod" $null | Out-Null

    $reference = Invoke-Com $installer "OpenDatabase" "InvokeMethod" @($MsiPath, $msiOpenDatabaseModeReadOnly)
    New-Item -ItemType Directory -Force -Path (Split-Path $TransformPath) | Out-Null
    if (Test-Path -LiteralPath $TransformPath) { Remove-Item -LiteralPath $TransformPath -Force }
    $generated = Invoke-Com $changed "GenerateTransform" "InvokeMethod" @($reference, $TransformPath)
    if (-not $generated) { throw "No differences found; transform was not created." }
    Invoke-Com $changed "CreateTransformSummaryInfo" "InvokeMethod" @($reference, $TransformPath, 0, 0) | Out-Null

    [Runtime.InteropServices.Marshal]::ReleaseComObject($reference) | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($changed) | Out-Null
    Write-Host "Transform erstellt: $TransformPath"
}
finally {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $workCopy -Force -ErrorAction SilentlyContinue
}
