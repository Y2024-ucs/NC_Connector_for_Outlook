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
    [string]$Signature
) {
    $start = $Source.IndexOf($Signature, [StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw "Could not locate method '$Signature'."
    }
    $lineStart = $Source.LastIndexOf([char]10, $start) + 1
    $indent = $Source.Substring($lineStart, $start - $lineStart)
    $closingLine = [string][char]10 + $indent + "}"
    $end = $Source.IndexOf($closingLine, $start, [StringComparison]::Ordinal)
    if ($end -lt 0) {
        throw "Could not isolate method '$Signature'."
    }
    return $Source.Substring($start, $end + $closingLine.Length - $start)
}

$composePath = "src\NcTalkOutlookAddIn\Controllers\SeparatePasswordDeliveryController.cs"
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
$attachmentPolicy = Read-Source "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentPolicy.cs"
$attachmentMaterialization = Read-Source "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentMaterialization.cs"
$attachmentQueue = Read-Source "src\NcTalkOutlookAddIn\NextcloudTalkAddIn.MailComposeSubscription.AttachmentQueue.cs"
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
    "internal void DispatchSeparatePasswordMailQueue("
$passwordMailPopulation = Get-MethodSlice `
    $compose `
    "private List<string> PopulatePasswordMail("
$manualPasswordFallback = Get-MethodSlice `
    $compose `
    "private bool TryOpenSeparatePasswordFallback("
$fileLinkWizardUi = Get-MethodSlice `
    $fileLink `
    "private bool RunFileLinkWizardOnUiThread("

$attachmentAdd = Get-MethodSlice `
    $attachmentFlow `
    "private void OnAttachmentAdd("
$beforeAttachmentAdd = Get-MethodSlice `
    $attachmentFlow `
    "private void OnBeforeAttachmentAdd("
$snapshotAttachments = Get-MethodSlice `
    $attachmentMaterialization `
    "private List<AttachmentSnapshot> SnapshotAttachments()"
$hiddenAttachment = Get-MethodSlice `
    $attachmentMaterialization `
    "private static bool IsHiddenAttachment("
$collectAttachments = Get-MethodSlice `
    $attachmentMaterialization `
    "private void CollectAttachmentSelectionsForShare("
$startAttachmentShareFlow = Get-MethodSlice `
    $attachmentQueue `
    "private async Task StartComposeAttachmentShareFlowAsync("
$prepareAttachmentSelections = Get-MethodSlice `
    $attachmentQueue `
    "private bool PrepareComposeAttachmentSelections("
$removeAttachmentsByIndices = Get-MethodSlice `
    $attachmentMaterialization `
    "private void RemoveAttachmentsByIndices("
$lastAddedBatch = Get-MethodSlice `
    $attachmentMaterialization `
    "private AttachmentBatchInfo BuildLastAddedBatchInfo("
$readAttachmentSettings = Get-MethodSlice `
    $attachmentPolicy `
    "private AttachmentAutomationSettings ReadAttachmentAutomationSettings()"
$attachmentSendGate = Get-MethodSlice `
    $attachmentPolicy `
    "private bool TryValidateAttachmentPolicyBeforeSend("
$removeSuppressedAttachment = Get-MethodSlice `
    $attachmentMaterialization `
    "private void RemoveSuppressedBeforeAddAttachmentByName("
$removeLastAttachmentBatch = Get-MethodSlice `
    $attachmentMaterialization `
    "private void RemoveLastAddedAttachmentBatch("
$queueSelectionScan = Get-MethodSlice `
    $fileLinkWizardFiles `
    "private async Task AddSelectionsAsync("
$initialFileSelection = Get-MethodSlice `
    $fileLinkWizardDragDrop `
    "private bool TryAddInitialFileSelection("

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
Assert-NotContains `
    "Attachment positions are not captured before server prefetch" `
    $startAttachmentShareFlow `
    "CollectAttachmentSelectionsForShare("
Assert-Contains `
    "Post-add capture is deferred until the UI handoff" `
    $startAttachmentShareFlow `
    "launchOptions.PrepareInitialSelections = () =>"
Assert-Contains `
    "The UI handoff captures the current attachment collection" `
    $prepareAttachmentSelections `
    "CollectAttachmentSelectionsForShare(selections, removeIndices, tempFiles);"
Assert-NotContains `
    "Attachment capture cannot yield before queue adoption" `
    $prepareAttachmentSelections `
    "await "
Assert-Precedes `
    "Current attachments are captured on the STA before wizard construction" `
    $fileLinkWizardUi `
    "!launchOptions.PrepareInitialSelections()" `
    "new FileLinkWizardForm("
Assert-NotContains `
    "Wizard construction and attachment adoption do not yield" `
    $fileLinkWizardUi `
    "await "
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
    $attachmentPolicy `
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

$deleteMethod = Read-Source "src\NcTalkOutlookAddIn\Services\ComposeShareCleanupService.cs"
Assert-Contains `
    "Compose cleanup method is present" `
    $deleteMethod `
    "internal bool TryDeleteComposeShareFolder("
Assert-NotContains `
    "Password delivery does not own remote cleanup" `
    $compose `
    "TryDeleteComposeShareFolder"
Assert-NotContains `
    "Remote cleanup has no Outlook COM dependency" `
    $deleteMethod `
    "Microsoft.Office.Interop"
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

$attachmentHandoffHarness = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class FileLinkSelection
    {
        internal string LocalPath;
        internal FileLinkSelection(string path) { LocalPath = path; }
    }
    __LAUNCH_OPTIONS__
}

namespace Outlook
{
    internal sealed class Attachments
    {
        internal readonly List<string> Names = new List<string>();
        internal int Count { get { return Names.Count; } }
        internal void Remove(int index) { Names.RemoveAt(index - 1); }
    }
}

public static class AttachmentHandoffRegression
{
    private static class OutlookAttachmentAutomationGuardService
    {
        internal sealed class GuardState { }
    }
    private static class DiagnosticsLogger
    {
        internal static void LogException(string category, string message, Exception ex) { }
    }
    private static class LogCategories { internal const string FileLink = "FILELINK"; }
    private static class ComInteropScope
    {
        internal static void TryRelease(object value, string category, string message) { }
    }
    private sealed class AttachmentBatchInfo
    {
        internal string Name;
        internal long SizeBytes;
    }
    private sealed class Mail
    {
        internal readonly Outlook.Attachments Attachments = new Outlook.Attachments();
    }
    private sealed class Owner
    {
        internal readonly TaskCompletionSource<bool> Prefetch = new TaskCompletionSource<bool>();
        internal readonly List<string> Queue = new List<string>();
        internal bool RejectQueue;
        internal bool TryGetAttachmentAutomationGuardState(
            string stage, string key, out OutlookAttachmentAutomationGuardService.GuardState state)
        { state = null; return false; }
        internal async Task<bool> RunFileLinkWizardForMailAsync(Mail mail, FileLinkWizardLaunchOptions options)
        {
            await Prefetch.Task;
            if (!options.PrepareInitialSelections()) { return false; }
            if (RejectQueue) { return false; }
            foreach (FileLinkSelection item in options.InitialSelections) { Queue.Add(item.LocalPath); }
            options.OnInitialQueueAdopted();
            return false; // The user cancels the wizard after adoption.
        }
    }
    private sealed class Subscription
    {
        internal readonly Mail _mail = new Mail();
        internal readonly Owner _owner = new Owner();
        internal bool _disposed;
        private readonly string _composeKey = "test";
        private bool _attachmentSuppressed;
        private readonly List<string> _pendingAddedBatch = new List<string>();
        internal int Captures;
        internal int CleanupCalls;
        internal Task Start() { return StartComposeAttachmentShareFlowAsync("threshold", 12, 1, null); }
        private void LogFileLink(string message) { }
        private void EndAttachmentSuppression(string reason) { _attachmentSuppressed = false; }
        private void CleanupTemporaryFiles(List<string> files) { CleanupCalls++; }
        private void CollectAttachmentSelectionsForShare(
            List<FileLinkSelection> selections, List<int> indices, List<string> files)
        {
            Captures++;
            for (int index = 0; index < _mail.Attachments.Count; index++)
            {
                selections.Add(new FileLinkSelection(_mail.Attachments.Names[index]));
                indices.Add(index + 1);
            }
        }
        __START_FLOW__
        __PREPARE_SELECTIONS__
        __REMOVE_ATTACHMENTS__
    }
    private static void Check(bool condition, string message)
    {
        if (!condition) { throw new InvalidOperationException(message); }
    }
    public static void Run()
    {
        var changed = new Subscription();
        changed._mail.Attachments.Names.AddRange(new[] { "A.pdf", "B.pdf" });
        Task changedFlow = changed.Start();
        Check(changed.Captures == 0, "Attachments were captured before prefetch completed.");
        changed._mail.Attachments.Remove(1);
        changed._mail.Attachments.Names.Add("C.pdf");
        changed._owner.Prefetch.SetResult(true);
        changedFlow.GetAwaiter().GetResult();
        Check(string.Join(",", changed._owner.Queue) == "B.pdf,C.pdf", "Queue did not capture the current attachments.");
        Check(changed._mail.Attachments.Count == 0, "Adopted originals remained or were restored after cancellation.");
        Check(changed.CleanupCalls == 1, "Temporary-file cleanup was skipped.");

        var rejected = new Subscription();
        rejected._mail.Attachments.Names.Add("keep.pdf");
        rejected._owner.RejectQueue = true;
        Task rejectedFlow = rejected.Start();
        rejected._owner.Prefetch.SetResult(true);
        rejectedFlow.GetAwaiter().GetResult();
        Check(rejected._mail.Attachments.Count == 1, "A rejected queue removed an Outlook attachment.");

        var closed = new Subscription();
        closed._mail.Attachments.Names.Add("closed.pdf");
        Task closedFlow = closed.Start();
        closed._disposed = true;
        closed._owner.Prefetch.SetResult(true);
        closedFlow.GetAwaiter().GetResult();
        Check(closed.Captures == 0 && closed._mail.Attachments.Count == 1, "A closed compose item was accessed after prefetch.");

        var empty = new Subscription();
        empty._mail.Attachments.Names.Add("removed.pdf");
        Task emptyFlow = empty.Start();
        empty._mail.Attachments.Remove(1);
        empty._owner.Prefetch.SetResult(true);
        emptyFlow.GetAwaiter().GetResult();
        Check(empty._owner.Queue.Count == 0, "An empty compose item opened a stale queue.");

        var failed = new Subscription();
        failed._mail.Attachments.Names.Add("offline.pdf");
        Task failedFlow = failed.Start();
        failed._owner.Prefetch.SetException(new InvalidOperationException("prefetch failed"));
        try { failedFlow.GetAwaiter().GetResult(); }
        catch (InvalidOperationException) { }
        Check(failed.Captures == 0 && failed._mail.Attachments.Count == 1, "Failed prefetch touched the attachments.");
        Check(failed.CleanupCalls == 1, "Failed prefetch skipped cleanup.");
    }
}
'@
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace(
    "__LAUNCH_OPTIONS__",
    (Get-MethodSlice $fileLinkLaunchOptions "internal sealed class FileLinkWizardLaunchOptions"))
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__START_FLOW__", $startAttachmentShareFlow)
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__PREPARE_SELECTIONS__", $prepareAttachmentSelections)
$attachmentHandoffHarness = $attachmentHandoffHarness.Replace("__REMOVE_ATTACHMENTS__", $removeAttachmentsByIndices)
Add-Type -TypeDefinition $attachmentHandoffHarness -Language CSharp -IgnoreWarnings -WarningAction SilentlyContinue
[AttachmentHandoffRegression]::Run()
Write-Host "[OK] Deferred attachment capture, changed collection, cancellation, rejected queue and failed prefetch"

Write-Host "All Outlook compose lifecycle regression checks passed."
