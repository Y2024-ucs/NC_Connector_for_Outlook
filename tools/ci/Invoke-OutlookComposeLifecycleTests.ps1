Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path

function Read-Source([string]$RelativePath) {
    return [IO.File]::ReadAllText(
        (Join-Path $ProjectRoot $RelativePath))
}

function Assert-True(
    [string]$Name,
    [bool]$Condition
) {
    if (-not $Condition) {
        throw "[FAIL] $Name"
    }
    Write-Host "[OK] $Name"
}

function Assert-Contains(
    [string]$Name,
    [string]$Source,
    [string]$Expected
) {
    Assert-True $Name $Source.Contains($Expected)
}

function Assert-NotContains(
    [string]$Name,
    [string]$Source,
    [string]$Unexpected
) {
    Assert-True $Name (-not $Source.Contains($Unexpected))
}

function Assert-Precedes(
    [string]$Name,
    [string]$Source,
    [string]$First,
    [string]$Second
) {
    $firstIndex = $Source.IndexOf($First, [StringComparison]::Ordinal)
    $secondIndex = $Source.IndexOf($Second, [StringComparison]::Ordinal)
    Assert-True $Name (
        $firstIndex -ge 0 `
            -and $secondIndex -gt $firstIndex)
}

function Get-MethodSlice(
    [string]$Source,
    [string]$Signature,
    [string]$NextSignature
) {
    $start = $Source.IndexOf($Signature, [StringComparison]::Ordinal)
    $end = $Source.IndexOf(
        $NextSignature,
        $start + $Signature.Length,
        [StringComparison]::Ordinal)
    if ($start -lt 0 -or $end -le $start) {
        throw "Could not isolate method '$Signature'."
    }
    return $Source.Substring($start, $end - $start)
}

$composePath = "src\NcTalkOutlookAddIn\Controllers\ComposeShareLifecycleController.cs"
$trackerPath = "src\NcTalkOutlookAddIn\Controllers\ComposeShareCleanupTracker.cs"
$sendPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.Send.cs"
$shareCleanupPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs"
$addinPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.cs"
$subscriptionPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.cs"
$attachmentFlowPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentFlow.cs"
$subscriptionRegistryPath = "src\NcTalkOutlookAddIn\Controllers\MailComposeSubscriptionRegistryController.cs"
$hooksPath = "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.Hooks.cs"
$fileLinkPath = "src\NcTalkOutlookAddIn\Controllers\FileLinkLaunchController.cs"
$fileLinkLaunchOptionsPath = "src\NcTalkOutlookAddIn\Models\FileLinkWizardLaunchOptions.cs"
$fileLinkWizardPath = "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.cs"
$fileLinkWizardFilesPath = "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.Files.cs"
$fileLinkWizardDragDropPath = "src\NcTalkOutlookAddIn\UI\FileLinkWizardForm.DragDrop.cs"
$projectPath = "src\NcTalkOutlookAddIn\NcTalkOutlookAddIn.csproj"

$compose = Read-Source $composePath
$tracker = Read-Source $trackerPath
$send = Read-Source $sendPath
$shareCleanup = Read-Source $shareCleanupPath
$addin = Read-Source $addinPath
$subscription = Read-Source $subscriptionPath
$attachmentFlow = Read-Source $attachmentFlowPath
$subscriptionRegistry = Read-Source $subscriptionRegistryPath
$hooks = Read-Source $hooksPath
$fileLink = Read-Source $fileLinkPath
$fileLinkLaunchOptions = Read-Source $fileLinkLaunchOptionsPath
$fileLinkWizard = Read-Source $fileLinkWizardPath
$fileLinkWizardFiles = Read-Source $fileLinkWizardFilesPath
$fileLinkWizardDragDrop = Read-Source $fileLinkWizardDragDropPath
$project = Read-Source $projectPath
$directPasswordDispatch = Get-MethodSlice `
    $compose `
    "internal void DispatchSeparatePasswordMailQueue(" `
    "private List<string> PopulatePasswordMail("
$passwordMailPopulation = Get-MethodSlice `
    $compose `
    "private List<string> PopulatePasswordMail(" `
    "private bool TryOpenSeparatePasswordFallback("
$manualPasswordFallback = Get-MethodSlice `
    $compose `
    "private bool TryOpenSeparatePasswordFallback(" `
    "private static bool ReadSubmittedOrAmbiguous("
$fileLinkWizardUi = Get-MethodSlice `
    $fileLink `
    "private bool RunFileLinkWizardOnUiThread(" `
    "private static FileLinkResult BuildSecretPlaceholderResult("

$attachmentAdd = Get-MethodSlice `
    $attachmentFlow `
    "private void OnAttachmentAdd(" `
    "private void OnBeforeAttachmentAdd("
$beforeAttachmentAdd = Get-MethodSlice `
    $attachmentFlow `
    "private void OnBeforeAttachmentAdd(" `
    "private void OnPropertyChange("
$snapshotAttachments = Get-MethodSlice `
    $attachmentFlow `
    "private List<AttachmentSnapshot> SnapshotAttachments()" `
    "private static long SumAttachmentBytes("
$hiddenAttachment = Get-MethodSlice `
    $attachmentFlow `
    "private static bool IsHiddenAttachment(" `
    "private static void ShowForcedAttachmentProcessingError("
$collectAttachments = Get-MethodSlice `
    $attachmentFlow `
    "private void CollectAttachmentSelectionsForShare(" `
    "private bool TryResolveAttachmentLocalPath("
$startAttachmentShareFlow = Get-MethodSlice `
    $attachmentFlow `
    "private async Task StartComposeAttachmentShareFlowAsync(" `
    "private bool TryBuildBeforeAddAttachmentCandidate("
$lastAddedBatch = Get-MethodSlice `
    $attachmentFlow `
    "private AttachmentBatchInfo BuildLastAddedBatchInfo(" `
    "private async Task StartComposeAttachmentShareFlowAsync("
$readAttachmentSettings = Get-MethodSlice `
    $attachmentFlow `
    "private AttachmentAutomationSettings ReadAttachmentAutomationSettings()" `
    "private async Task<AttachmentAutomationSettings> ReadAttachmentAutomationSettingsAsync()"
$attachmentSendGate = Get-MethodSlice `
    $attachmentFlow `
    "private bool TryValidateAttachmentPolicyBeforeSend(" `
    "private int CountPolicyRelevantAttachments()"
$removeSuppressedAttachment = Get-MethodSlice `
    $attachmentFlow `
    "private void RemoveSuppressedBeforeAddAttachmentByName(" `
    "private void RemoveAttachmentsByIndices("
$removeLastAttachmentBatch = Get-MethodSlice `
    $attachmentFlow `
    "private void RemoveLastAddedAttachmentBatch(" `
    "private void EndAttachmentSuppression("
$queueSelectionScan = Get-MethodSlice `
    $fileLinkWizardFiles `
    "private async Task AddSelectionsAsync(" `
    "private void AddInitialSelections("
$initialFileSelection = Get-MethodSlice `
    $fileLinkWizardDragDrop `
    "private bool TryAddInitialFileSelection(" `
    "private bool TryReserveSelection("

Assert-Precedes `
    "Hidden attachments are ignored before post-add batching" `
    $attachmentAdd `
    "IsHiddenAttachment(attachment)" `
    "_pendingAddedBatch.Add("
Assert-Precedes `
    "Hidden attachments are allowed before automation preflight" `
    $beforeAttachmentAdd `
    "IsHiddenAttachment(attachment)" `
    "ReadAttachmentAutomationSettings()"
Assert-Precedes `
    "Hidden attachments are excluded from threshold snapshots" `
    $snapshotAttachments `
    "IsHiddenAttachment(attachment)" `
    "snapshots.Add("
Assert-Contains `
    "Hidden attachment detection reads PR_ATTACHMENT_HIDDEN" `
    $hiddenAttachment `
    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
Assert-Contains `
    "Hidden attachment PropertyAccessor is released" `
    $hiddenAttachment `
    "ComInteropScope.TryRelease("
Assert-Precedes `
    "Hidden attachments are excluded before FileLink materialization" `
    $collectAttachments `
    "IsHiddenAttachment(attachment)" `
    "TryResolveAttachmentLocalPath("
Assert-Precedes `
    "Post-add attachments are detached only through queue adoption" `
    $startAttachmentShareFlow `
    "OnInitialQueueAdopted = () =>" `
    "bool wizardAccepted = await _owner.RunFileLinkWizardForMailAsync(_mail, launchOptions);"
Assert-Contains `
    "Queue adoption removes the original Outlook attachments" `
    $startAttachmentShareFlow `
    'RemoveAttachmentsByIndices('
Assert-Contains `
    "Attachment launch options expose the queue-adoption boundary" `
    $fileLinkLaunchOptions `
    "internal Action OnInitialQueueAdopted { get; set; }"
Assert-Contains `
    "The wizard reports its accepted queue size" `
    $fileLinkWizard `
    "internal int QueuedSelectionCount"
Assert-Precedes `
    "The complete attachment queue is checked before adoption" `
    $fileLinkWizardUi `
    "wizard.QueuedSelectionCount" `
    "launchOptions.OnInitialQueueAdopted();"
Assert-Precedes `
    "Attachment ownership transfers before the wizard opens" `
    $fileLinkWizardUi `
    "launchOptions.OnInitialQueueAdopted();" `
    "wizard.ShowDialog()"
Assert-Precedes `
    "Synchronous attachment events refresh an expired settings snapshot" `
    $readAttachmentSettings `
    "HasFreshAttachmentAutomationSettingsSnapshot()" `
    "BeginAttachmentAutomationSettingsRefresh();"
Assert-Contains `
    "The send gate waits for a current attachment policy snapshot" `
    $attachmentSendGate `
    "if (!HasFreshAttachmentAutomationSettingsSnapshot()"
Assert-Contains `
    "Settings changes invalidate open compose attachment caches" `
    $addin `
    ".RefreshAttachmentAutomationSettings();"
Assert-Contains `
    "The compose registry refreshes every open subscription" `
    $subscriptionRegistry `
    "current[i].RefreshAttachmentAutomationSettings();"
Assert-Contains `
    "Superseded attachment-policy requests cannot replace current settings" `
    $attachmentFlow `
    "== _attachmentAutomationSettingsRefreshGeneration"
Assert-Contains `
    "The threshold prompt uses the last attachment size" `
    $lastAddedBatch `
    "latestBatchEntry.SizeBytes)"
Assert-NotContains `
    "The threshold prompt does not label a batch total as the last file size" `
    $lastAddedBatch `
    "total +="
Assert-Contains `
    "Local queue snapshots are built off the wizard thread" `
    $queueSelectionScan `
    "FileLinkQueueNode snapshot = await Task.Run("
Assert-Contains `
    "Local queue snapshots use the cancellable scan token" `
    $queueSelectionScan `
    "selection,`r`n                                token)"
Assert-NotContains `
    "Interactive queue scans do not use an uncancellable token" `
    $queueSelectionScan `
    "CancellationToken.None"
Assert-Precedes `
    "Queue snapshots return to the wizard before UI state changes" `
    $queueSelectionScan `
    "FileLinkQueueNode snapshot = await Task.Run(" `
    "AddPreparedSelection(selection, snapshot);"
Assert-NotContains `
    "Queue scans retain the captured WinForms context" `
    $queueSelectionScan `
    "ConfigureAwait(false)"
Assert-Contains `
    "The wizard busy state includes local queue scans" `
    $fileLinkWizard `
    "|| _queueScanInProgress;"
Assert-Contains `
    "Drag and drop awaits the background queue scan" `
    $fileLinkWizardDragDrop `
    "await AddSelectionsAsync(selections);"
Assert-Precedes `
    "Only individual attachment files use synchronous initial capture" `
    $initialFileSelection `
    "!= FileLinkSelectionType.File" `
    "CancellationToken.None"
Assert-Precedes `
    "Suppressed host cleanup preserves hidden attachments" `
    $removeSuppressedAttachment `
    "IsHiddenAttachment(attachment)" `
    "ReadAttachmentName(attachment)"
Assert-Contains `
    "Last-batch removal starts from the visible attachment snapshot" `
    $removeLastAttachmentBatch `
    "List<AttachmentSnapshot> attachments = SnapshotAttachments();"
Assert-Contains `
    "Last-batch removal uses the filtered attachment indices" `
    $removeLastAttachmentBatch `
    'RemoveAttachmentsByIndices(removeIndices, "remove_last_batch");'
Assert-NotContains `
    "Last-batch removal does not delete the physical collection tail" `
    $removeLastAttachmentBatch `
    "attachments.Remove(attachments.Count)"

Assert-Precedes `
    "Primary Send captures recipients before direct password dispatch" `
    $send `
    "CapturePasswordDispatchRecipients();" `
    "_owner.DispatchSeparatePasswordMails("
Assert-Precedes `
    "Primary Send captures the sender before direct password dispatch" `
    $send `
    "CapturePasswordDispatchSender();" `
    "_owner.DispatchSeparatePasswordMails("
Assert-Precedes `
    "The password queue is consumed before direct dispatch" `
    $send `
    "_passwordDispatchQueue.Clear();" `
    "_owner.DispatchSeparatePasswordMails("
Assert-NotContains `
    "Primary Send no longer waits for Sent-folder confirmation" `
    $send `
    "TryArmPendingPasswordDrafts"
Assert-NotContains `
    "Primary Send does not cancel because password auto-send failed" `
    $send `
    "SharingPasswordMailPrepareFailed"

Assert-Precedes `
    "Final password dispatch is prepared before Outlook mail creation" `
    $directPasswordDispatch `
    "PrepareSeparatePasswordDispatch(" `
    "_owner.OutlookApplication.CreateItem("
Assert-Precedes `
    "Final sender identity is applied before the password body" `
    $passwordMailPopulation `
    "ApplyAndVerifySeparatePasswordSender(" `
    "ApplySeparatePasswordBody("
Assert-Precedes `
    "Password body is complete before backend signature insertion" `
    $passwordMailPopulation `
    "ApplySeparatePasswordBody(" `
    "ApplySeparatePasswordBackendSignature("
Assert-Precedes `
    "Recipients are resolved after final body and signature assembly" `
    $passwordMailPopulation `
    "ApplySeparatePasswordBackendSignature(" `
    "ApplySeparatePasswordRecipientsForSend("
Assert-Precedes `
    "Password mail is fully populated before direct Send" `
    $directPasswordDispatch `
    "PopulatePasswordMail(" `
    "((Outlook._MailItem)passwordMail).Send();"
Assert-NotContains `
    "Direct password delivery does not persist an intermediate draft" `
    $directPasswordDispatch `
    ".Save();"
Assert-Contains `
    "Automatic Send failure offers a fully prepared manual fallback" `
    $directPasswordDispatch `
    "TryOpenSeparatePasswordFallback("
Assert-Contains `
    "Ambiguous submission suppresses a duplicate manual send" `
    $directPasswordDispatch `
    "ReadSubmittedOrAmbiguous(passwordMail)"
Assert-Contains `
    "A definite Send and fallback failure is reported to the user" `
    $directPasswordDispatch `
    "ShowPasswordMailFailure(ex.Message);"
Assert-Contains `
    "Secrets fallback warning also covers a prepared manual fallback" `
    $directPasswordDispatch `
    "(sent > 0 || manual > 0)"
Assert-Contains `
    "Manual Send fallback displays the prepared mail" `
    $manualPasswordFallback `
    "fallback.Display(false);"
Assert-Contains `
    "Build fallback uses normalized To recipients without resolution" `
    $manualPasswordFallback `
    "fallback.To = toRecipients;"
Assert-Contains `
    "Build fallback uses normalized Cc recipients without resolution" `
    $manualPasswordFallback `
    "fallback.CC = ccRecipients;"
Assert-Contains `
    "Build fallback uses normalized Bcc recipients without resolution" `
    $manualPasswordFallback `
    "fallback.BCC = bccRecipients;"
Assert-NotContains `
    "Build fallback does not repeat automatic recipient resolution" `
    $manualPasswordFallback `
    "ApplySeparatePasswordRecipientsForSend("
Assert-Precedes `
    "Build fallback reconciles the managed signature after display" `
    $manualPasswordFallback `
    "fallback.Display(false);" `
    "ApplySeparatePasswordBackendSignatureToDisplayedFallback("
Assert-Contains `
    "Unexpected password dispatch failures do not escape the primary Send callback" `
    $send `
    "The primary send continues."
Assert-True `
    "Pending Sent-folder controller is removed" `
    (-not (Test-Path -LiteralPath (
        Join-Path $ProjectRoot `
            "src\NcTalkOutlookAddIn\Controllers\PendingPasswordDraftController.cs")))
Assert-NotContains `
    "Project no longer compiles the pending Sent-folder controller" `
    $project `
    "PendingPasswordDraftController.cs"

$insertedIndex = $fileLink.IndexOf(
    "bool inserted =",
    [StringComparison]::Ordinal)
$insertFailureIndex = $fileLink.IndexOf(
    "if (!inserted)",
    $insertedIndex,
    [StringComparison]::Ordinal)
$armIndex = $fileLink.IndexOf(
    "composeSubscription.ArmShareCleanup(",
    $insertFailureIndex,
    [StringComparison]::Ordinal)
$passwordRegistrationIndex = $fileLink.IndexOf(
    "if (registerSeparatePassword)",
    $armIndex,
    [StringComparison]::Ordinal)
Assert-True `
    "Successful FileLink insertion arms compose cleanup before follow-up handling" `
    ($insertedIndex -ge 0 `
        -and $insertFailureIndex -gt $insertedIndex `
        -and $armIndex -gt $insertFailureIndex `
        -and $passwordRegistrationIndex -gt $armIndex)

Assert-Contains `
    "Compose cleanup arms the focused tracker" `
    $shareCleanup `
    "_shareCleanupTracker.Arm(record)"
Assert-Contains `
    "Compose cleanup tracker exposes ReleaseAll" `
    $tracker `
    "internal int ReleaseAll()"
Assert-Contains `
    "Compose cleanup tracker exposes Drain" `
    $tracker `
    "internal List<ComposeShareCleanupRecord> Drain()"

$afterWriteIndex = $shareCleanup.IndexOf(
    "private void OnAfterWrite()",
    [StringComparison]::Ordinal)
$releaseIndex = $shareCleanup.IndexOf(
    "_shareCleanupTracker.ReleaseAll();",
    $afterWriteIndex,
    [StringComparison]::Ordinal)
$unloadIndex = $shareCleanup.IndexOf(
    "private void OnUnload()",
    [StringComparison]::Ordinal)
$inspectorCloseIndex = $shareCleanup.IndexOf(
    "private void OnInspectorClosed()",
    [StringComparison]::Ordinal)
$completionIndex = $shareCleanup.IndexOf(
    "private void CompleteComposeShareCleanup(",
    [StringComparison]::Ordinal)
$drainIndex = $shareCleanup.IndexOf(
    "_shareCleanupTracker.Drain();",
    $completionIndex,
    [StringComparison]::Ordinal)
$disposeIndex = $shareCleanup.IndexOf(
    "Dispose(detachItemEvents);",
    $drainIndex,
    [StringComparison]::Ordinal)
$cleanupQueueIndex = $shareCleanup.IndexOf(
    "_owner.QueueCreatedShareCleanup(",
    $disposeIndex,
    [StringComparison]::Ordinal)
Assert-True `
    "AfterWrite releases shares that Outlook persisted" `
    ($afterWriteIndex -ge 0 `
        -and $releaseIndex -gt $afterWriteIndex `
        -and $releaseIndex -lt $unloadIndex)
Assert-True `
    "Inspector close and inline unload share one captured-state finalizer" `
    ($unloadIndex -ge 0 `
        -and $inspectorCloseIndex -gt $unloadIndex `
        -and $completionIndex -gt $inspectorCloseIndex `
        -and $drainIndex -gt $completionIndex `
        -and $disposeIndex -gt $drainIndex `
        -and $cleanupQueueIndex -gt $disposeIndex)
Assert-NotContains `
    "Share cleanup terminal handlers do not inspect the MailItem" `
    $shareCleanup `
    "_mail"

Assert-Contains `
    "Compose subscription hooks AfterWrite" `
    $subscription `
    "_events.AfterWrite += OnAfterWrite;"
Assert-Contains `
    "Compose subscription hooks Unload" `
    $subscription `
    "_events.Unload += OnUnload;"
Assert-Contains `
    "Compose subscription unhooks AfterWrite" `
    $subscription `
    "_events.AfterWrite -= OnAfterWrite;"
Assert-Contains `
    "Compose subscription unhooks Unload" `
    $subscription `
    "_events.Unload -= OnUnload;"
Assert-Contains `
    "Compose subscription hooks the concrete Inspector close event" `
    $subscription `
    "inspectorEvents.Close += OnInspectorClosed;"
Assert-Contains `
    "Compose subscription unhooks the concrete Inspector close event" `
    $subscription `
    "inspectorEvents.Close -= OnInspectorClosed;"
Assert-Contains `
    "A concrete replacement Inspector rebinds the lifecycle sink" `
    $subscription `
    "ComInteropScope.AreSameObject("
Assert-Contains `
    "Fallback binding does not replace an existing Inspector sink" `
    $subscription `
    "(_inspectorEvents != null && inspector == null)"
Assert-Contains `
    "NewInspector passes the concrete Inspector to the compose subscription" `
    ($hooks.Replace("`r`n", "`n")) `
    "null,`n                        inspector);"

$composeCleanupSources = $subscription + "`n" + $send + "`n" + $shareCleanup
foreach ($obsoleteClosePath in @(
    "_events.Close += OnClose;",
    "ScheduleSurfaceCloseVerification",
    "OnCleanupGraceTimerTick",
    "IsMailComposeSurfaceOpen",
    "_cleanupGraceTimer"
)) {
    Assert-NotContains `
        ("Compose cleanup does not use close polling: " + $obsoleteClosePath) `
        $composeCleanupSources `
        $obsoleteClosePath
}
Assert-NotContains `
    "Compose cleanup does not poll Inspector state" `
    $shareCleanup `
    "IsMailComposeSurfaceOpen"

Assert-NotContains `
    "FileLink launch no longer blocks on current-user lookup" `
    $fileLink `
    "currentUserIdTask"
Assert-Contains `
    "Lifecycle origin remains attached to compose share state" `
    $fileLink `
    "ComposeLifecycleOrigin.Create("

$deleteMethodStart = $compose.IndexOf(
    "internal bool TryDeleteComposeShareFolder(",
    [StringComparison]::Ordinal)
$deleteMethodEnd = $compose.IndexOf(
    "internal void CaptureSeparatePasswordSignatureSnapshot(",
    $deleteMethodStart,
    [StringComparison]::Ordinal)
Assert-True `
    "Compose cleanup method is present" `
    ($deleteMethodStart -ge 0 `
        -and $deleteMethodEnd -gt $deleteMethodStart)
$deleteMethod = $compose.Substring(
    $deleteMethodStart,
    $deleteMethodEnd - $deleteMethodStart)
Assert-Contains `
    "Compose cleanup requires its captured origin" `
    $deleteMethod `
    "entry.Origin == null || !entry.Origin.IsComplete()"
Assert-Contains `
    "Compose cleanup uses its captured origin" `
    $deleteMethod `
    "entry.Origin.ToConfiguration()"
Assert-NotContains `
    "Compose cleanup never falls back to current settings" `
    $deleteMethod `
    "_owner.CurrentSettings"

$removed = @(
    "Models\ComposeLifecycleRecord.cs",
    "Controllers\ComposeLifecycleCoordinator.cs",
    "Controllers\ComposeLifecycleCoordinator.FolderLookup.cs",
    "Controllers\ComposeLifecycleCoordinator.Marker.cs",
    "Controllers\ComposeLifecycleCoordinator.Processing.cs",
    "Controllers\ComposeLifecycleCoordinator.Recovery.cs",
    "Controllers\ComposeLifecycleCoordinator.SentItems.cs",
    "Controllers\SeparatePasswordDispatchController.cs",
    "Controllers\SeparatePasswordDispatchController.Preparation.cs",
    "Controllers\SeparatePasswordDispatchController.Recipients.cs",
    "Controllers\SeparatePasswordDispatchController.Sender.cs",
    "Controllers\SeparatePasswordDispatchController.Signature.cs",
    "Services\ComposeLifecycleJournal.cs"
)
foreach ($relative in $removed) {
    Assert-True `
        ("Legacy source removed: " + $relative) `
        (-not (Test-Path (
            Join-Path `
                $ProjectRoot `
                ("src\NcTalkOutlookAddIn\" + $relative))))
    Assert-NotContains `
        ("Legacy project include removed: " + $relative) `
        $project `
        $relative
}
Assert-Contains `
    "Project includes the compact cleanup record" `
    $project `
    'Models\ComposeShareCleanupRecord.cs'
Assert-Contains `
    "Project includes the compose cleanup tracker" `
    $project `
    'Controllers\ComposeShareCleanupTracker.cs'
Assert-Contains `
    "Project includes the compose cleanup subscription partial" `
    $project `
    'NextcloudTalkAddIn.MailComposeSubscription.ShareCleanup.cs'

Write-Host "All Outlook compose lifecycle regression checks passed."
