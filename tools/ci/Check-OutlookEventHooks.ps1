Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$SourceRoot = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn"

$runtimeFiles = @(
    Get-ChildItem -Path $SourceRoot -Filter "NextcloudTalkAddIn*.cs"
    Get-ChildItem -Path (Join-Path $SourceRoot "Controllers") -Filter "*Lifecycle*.cs"
    Get-ChildItem -Path (Join-Path $SourceRoot "Services") -Filter "*Lifecycle*.cs"
) | Sort-Object FullName -Unique
if ($runtimeFiles.Count -eq 0) {
    throw "No NextcloudTalkAddIn runtime files found."
}

$additions = New-Object System.Collections.Generic.List[object]
$removals = New-Object System.Collections.Generic.List[object]
$eventPattern = '(?<target>[A-Za-z_][\w\.]*)\.(?<event>\w+)\s*(?<op>\+=|-=)\s*(?<handler>[A-Za-z_][\w\.]*)\s*;'

foreach ($file in $runtimeFiles) {
    $relative = $file.FullName.Substring($ProjectRoot.Length + 1)
    $source = Get-Content -Path $file.FullName -Raw
    foreach ($match in [regex]::Matches($source, $eventPattern)) {
        $line = ([regex]::Matches($source.Substring(0, $match.Index), "`n")).Count + 1

        $entry = [PSCustomObject]@{
            File = $relative
            Line = $line
            Target = $match.Groups["target"].Value
            Event = $match.Groups["event"].Value
            Handler = $match.Groups["handler"].Value
            Key = $match.Groups["event"].Value + "::" + $match.Groups["handler"].Value
        }

        if ($match.Groups["op"].Value -eq "+=") {
            $additions.Add($entry)
        } else {
            $removals.Add($entry)
        }
    }
}

$removalKeys = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::Ordinal)
$removalEvents = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::Ordinal)
foreach ($removal in $removals) {
    [void]$removalKeys.Add($removal.Key)
    [void]$removalEvents.Add($removal.Event)
}

$failures = New-Object System.Collections.Generic.List[string]
foreach ($addition in $additions) {
    $hasExactRemoval = $removalKeys.Contains($addition.Key)
    $hasStoredHandlerRemoval = $addition.Handler -eq "handler" -and $removalEvents.Contains($addition.Event)
    if (-not $hasExactRemoval -and -not $hasStoredHandlerRemoval) {
        $failures.Add("$($addition.File):$($addition.Line) subscribes $($addition.Event) += $($addition.Handler), but no matching unsubscribe was found.")
    }
}

foreach ($requiredEvent in @("InlineResponse", "InlineResponseClose", "SelectionChange")) {
    $matchingAddition = $additions | Where-Object { $_.Event -eq $requiredEvent } | Select-Object -First 1
    if ($null -eq $matchingAddition) {
        $failures.Add("Outlook lifecycle does not subscribe Explorer.$requiredEvent.")
    }
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    throw "Outlook event hook symmetry check failed with $($failures.Count) issue(s)."
}

Write-Host "Outlook event hook symmetry OK: $($additions.Count) subscription(s), $($removals.Count) unsubscribe(s)."
