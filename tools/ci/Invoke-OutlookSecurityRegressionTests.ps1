Param(
    [string]$ProjectRoot = "."
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-security-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    $testSource = Join-Path $TempRoot "OutlookSecurityRegressionTests.cs"
    @'
using System;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class AppDataPaths
    {
        internal static string GetLocalRootDirectory()
        {
            return Path.GetTempPath();
        }
    }

    internal static class AddinVersionInfo
    {
        internal static string GetVersion()
        {
            return "3.3.0";
        }
    }

    internal static class Strings
    {
        internal static string TalkVersionUnknown { get { return "unknown"; } }
        internal static string UpdateChangelogEmpty { get { return "empty"; } }
        internal static string UpdateAvailableMessageFormat { get { return "{0}: {1}"; } }
        internal static string UpdateChangelogAdded { get { return "Added"; } }
        internal static string UpdateChangelogChanged { get { return "Changed"; } }
        internal static string UpdateChangelogFixed { get { return "Fixed"; } }
        internal static string ConnectionFailureCertificateSummary { get { return "certificate"; } }
        internal static string ConnectionFailureCertificateGuidance { get { return "certificate guidance"; } }
        internal static string ConnectionFailureDnsSummary { get { return "dns"; } }
        internal static string ConnectionFailureDnsGuidance { get { return "dns guidance"; } }
        internal static string ConnectionFailureProxySummary { get { return "proxy"; } }
        internal static string ConnectionFailureProxyGuidance { get { return "proxy guidance"; } }
        internal static string ConnectionFailureTimeoutSummary { get { return "timeout"; } }
        internal static string ConnectionFailureTimeoutGuidance { get { return "timeout guidance"; } }
        internal static string ConnectionFailureTlsSummary { get { return "tls"; } }
        internal static string ConnectionFailureTlsGuidance { get { return "tls guidance"; } }
        internal static string ConnectionFailureGenericSummary { get { return "generic"; } }
        internal static string ConnectionFailureGenericGuidance { get { return "generic guidance"; } }
    }
}

namespace NcTalkOutlookAddIn.Settings
{
    internal sealed class AddinSettings
    {
        internal bool UpdateNotifyEnabled { get; set; }
        internal string UpdateInstallId { get; set; }
        internal string UpdateLastCheckedAtUtc { get; set; }
        internal string UpdateLatestVersion { get; set; }
        internal string UpdateReleaseUrl { get; set; }
        internal string UpdateDownloadUrl { get; set; }
        internal string UpdatePublishedAt { get; set; }
        internal string UpdateChangelogTitle { get; set; }
        internal string UpdateChangelogText { get; set; }
        internal string UpdateLastNotifiedVersion { get; set; }
        internal string UpdateLastNotifiedDateUtc { get; set; }
        internal bool TransportTlsUseSystemDefault { get; set; }
        internal bool TransportTlsEnable12 { get; set; }
        internal bool TransportTlsEnable13 { get; set; }
    }
}

internal static class OutlookSecurityRegressionTests
{
    private static int failures;

    private static void Check(string name, bool condition, string detail = "")
    {
        if (condition)
        {
            Console.WriteLine("[OK] " + name);
            return;
        }

        failures++;
        Console.Error.WriteLine("[FAIL] " + name + (string.IsNullOrEmpty(detail) ? "" : ": " + detail));
    }

    public static int Main()
    {
        TestNextcloudUriBoundary();
        TestStructuredSecretRedaction();
        TestUpdateTargetPolicy();
        TestAtomicSettingsTransaction();
        TestTransportSecurityConfigurator();

        if (failures > 0)
        {
            Console.Error.WriteLine(failures + " security regression test(s) failed.");
            return 1;
        }

        Console.WriteLine("All Outlook security regression tests passed.");
        return 0;
    }

    private static void TestNextcloudUriBoundary()
    {
        string normalized;
        Check(
            "Nextcloud base URL defaults to HTTPS",
            NextcloudUriValidator.TryNormalizeBaseUrl("cloud.example.test/nextcloud/", out normalized)
                && normalized == "https://cloud.example.test/nextcloud",
            normalized);
        Check(
            "Explicit HTTP Nextcloud base URL is rejected",
            !NextcloudUriValidator.TryNormalizeBaseUrl("http://cloud.example.test", out normalized));
        Check(
            "Nextcloud URL credentials are rejected",
            !NextcloudUriValidator.TryNormalizeBaseUrl("https://user:secret@cloud.example.test", out normalized));

        string resolved;
        Check(
            "Relative endpoint remains on configured Nextcloud path",
            NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "index.php/login/v2",
                "https://cloud.example.test/nextcloud",
                out resolved)
                && resolved == "https://cloud.example.test/nextcloud/index.php/login/v2",
            resolved);
        Check(
            "Cross-origin endpoint is rejected",
            !NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "https://evil.example.test/token",
                "https://cloud.example.test",
                out resolved));
        Check(
            "Scheme-relative cross-origin endpoint is rejected",
            !NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "//evil.example.test/token",
                "https://cloud.example.test",
                out resolved));
        Check(
            "Different Nextcloud port is a different origin",
            !NextcloudUriValidator.TryResolveSameOriginHttpsUrl(
                "https://cloud.example.test:8443/token",
                "https://cloud.example.test",
                out resolved));
    }

    private static void TestStructuredSecretRedaction()
    {
        DiagnosticsLogger.SetAnonymization(false, string.Empty);
        MethodInfo sanitizer = typeof(DiagnosticsLogger).GetMethod(
            "SanitizeMessage",
            BindingFlags.NonPublic | BindingFlags.Static);
        string input =
            "pollToken=alpha token: 'bravo' appPassword=\"charlie\" "
            + "{\"roomToken\":\"delta\"} https://cloud.example.test/path?shareToken=echo";
        string output = Convert.ToString(sanitizer.Invoke(null, new object[] { input }));

        Check(
            "Structured secrets are redacted with PII anonymization disabled",
            !output.Contains("alpha")
                && !output.Contains("bravo")
                && !output.Contains("charlie")
                && !output.Contains("delta")
                && !output.Contains("echo"),
            output);
        Check("Redaction marker is present", output.Contains("<REDACTED>"), output);
    }

    private static void TestUpdateTargetPolicy()
    {
        var result = new UpdateCheckResult
        {
            DownloadUrl = "https://evil.example.test/update.msi",
            ReleaseUrl = "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/tag/v3.4.0"
        };
        Check(
            "Untrusted download falls back to trusted release page",
            UpdateCheckService.GetPreferredOpenUrl(result)
                == "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/tag/v3.4.0");

        result.DownloadUrl =
            "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/download/v3.4.0/NC-Connector.msi";
        Check(
            "Trusted GitHub release download is accepted",
            UpdateCheckService.GetPreferredOpenUrl(result) == result.DownloadUrl);

        result.DownloadUrl = "http://github.com/nc-connector/NC_Connector_for_Outlook/releases/download/v3.4.0/file.msi";
        result.ReleaseUrl = "https://github.com/another/repository/releases/tag/v3.4.0";
        Check(
            "HTTP and wrong-repository update targets are rejected",
            UpdateCheckService.GetPreferredOpenUrl(result) == string.Empty);

        var settings = new AddinSettings
        {
            UpdateLatestVersion = "3.2.9",
            UpdateDownloadUrl =
                "https://github.com/nc-connector/NC_Connector_for_Outlook/releases/download/v3.2.9/file.msi"
        };
        UpdateCheckResult cached = UpdateCheckService.BuildCachedResult(settings);
        Check("Cached server state cannot downgrade local version comparison", !cached.UpdateAvailable);

        settings.UpdateLatestVersion = "3.4.0";
        cached = UpdateCheckService.BuildCachedResult(settings);
        Check("Newer cached version is detected locally", cached.UpdateAvailable);
    }

    private static void TestAtomicSettingsTransaction()
    {
        string directory = Path.Combine(Path.GetTempPath(), "nc4ol-settings-transaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string primary = Path.Combine(directory, "settings_profile.xml");
            var transaction = new SettingsFileTransaction(primary);
            Func<string, bool> healthy = path =>
                File.Exists(path) && File.ReadAllText(path, Encoding.UTF8).StartsWith("valid:", StringComparison.Ordinal);

            using (transaction.AcquireLock())
            {
                transaction.Commit(stream => Write(stream, "valid:first"), healthy);
                transaction.Commit(stream => Write(stream, "valid:second"), healthy);
            }

            Check("Atomic settings commit writes the new primary", File.ReadAllText(primary) == "valid:second");
            Check("Atomic settings commit keeps the previous valid backup", File.ReadAllText(transaction.BackupPath) == "valid:first");

            File.WriteAllText(primary, "corrupt", Encoding.UTF8);
            using (transaction.AcquireLock())
            {
                transaction.Commit(stream => Write(stream, "valid:third"), healthy);
            }
            Check("Corrupt primary does not replace the last valid backup", File.ReadAllText(transaction.BackupPath) == "valid:first");

            File.WriteAllText(primary, "corrupt-again", Encoding.UTF8);
            using (transaction.AcquireLock())
            {
                Check("Valid backup restores the primary", transaction.TryRestorePrimaryFromBackup(healthy));
            }
            Check("Restored primary equals the valid backup", File.ReadAllText(primary) == "valid:first");
            Check("Atomic settings transaction leaves no temp files", Directory.GetFiles(directory, "*.tmp").Length == 0);
        }
        finally
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
    }

    private static void TestTransportSecurityConfigurator()
    {
        SecurityProtocolType previous = ServicePointManager.SecurityProtocol;
        try
        {
            var settings = new AddinSettings
            {
                TransportTlsUseSystemDefault = false,
                TransportTlsEnable12 = true,
                TransportTlsEnable13 = false
            };
            SecurityProtocolType applied =
                TransportSecurityConfigurator.ApplyFromSettings(settings, "security_regression_test");

            Check(
                "TLS 1.2 setting changes the active runtime protocol",
                applied == SecurityProtocolType.Tls12
                    && ServicePointManager.SecurityProtocol == SecurityProtocolType.Tls12,
                ServicePointManager.SecurityProtocol.ToString());

            bool systemDefaultTlsDisabled;
            bool strongCryptoDisabled;
            Check(
                "Runtime enables system-default TLS support",
                AppContext.TryGetSwitch(
                    "Switch.System.Net.DontEnableSystemDefaultTlsVersions",
                    out systemDefaultTlsDisabled)
                    && !systemDefaultTlsDisabled);
            Check(
                "Runtime enables strong cryptography support",
                AppContext.TryGetSwitch(
                    "Switch.System.Net.DontEnableSchUseStrongCrypto",
                    out strongCryptoDisabled)
                    && !strongCryptoDisabled);
            Check(
                "System-default TLS maps to the runtime default",
                TransportSecurityConfigurator.BuildProtocol(true, false, false)
                    == SecurityProtocolType.SystemDefault);
            Check(
                "TLS 1.3 selection keeps its runtime protocol flag",
                (int)TransportSecurityConfigurator.BuildProtocol(false, false, true) == 12288);
        }
        finally
        {
            ServicePointManager.SecurityProtocol = previous;
        }
    }

    private static void Write(Stream stream, string content)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8

    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc"
    }

    $sources = @(
        $testSource,
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Models\UpdateCheckResult.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\UpdateCheckService.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Settings\SettingsFileTransaction.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\DiagnosticsLogger.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\HttpFailureDiagnostics.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\LogCategories.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NcJson.cs"),
        (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudUriValidator.cs")
    )
    $references = @(
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Security.dll",
        "/reference:System.Web.Extensions.dll"
    )

    $exe = Join-Path $TempRoot "OutlookSecurityRegressionTests.exe"
    & $csc /nologo /target:exe "/out:$exe" @references @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    & $exe
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    $settingsFormPath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\SettingsForm.cs"
    $settingsFormSource = Get-Content -LiteralPath $settingsFormPath -Raw
    $settingsGeneralSource = Get-Content -LiteralPath (Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\UI\SettingsForm.General.cs") -Raw
    if ($settingsFormSource -notmatch [regex]::Escape("_saveButton.DialogResult = DialogResult.None;")) {
        throw "Settings save button can close the form before URL validation completes."
    }
    if ($settingsFormSource -notmatch "TryNormalizeBaseUrl\(requestedServerUrl,\s*out normalizedServerUrl\)") {
        throw "Settings save path does not validate and normalize the configured Nextcloud URL."
    }
    if ($settingsGeneralSource -notmatch "TryNormalizeBaseUrl\(baseUrl,\s*out normalizedUrl\)") {
        throw "Settings connection test does not validate and normalize the configured Nextcloud URL."
    }
    Write-Host "[OK] Settings UI rejects invalid Nextcloud URLs before save or connection test"
}
finally {
    if (Test-Path $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}
