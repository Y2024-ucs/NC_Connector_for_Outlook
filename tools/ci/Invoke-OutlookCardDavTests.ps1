Param([string]$ProjectRoot = ".")
$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-carddav-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null
try {
    $testSource = Join-Path $TempRoot "CardDavTests.cs"
    @'
using System;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Services;

internal static class CardDavTests
{
    private const string Book = "https://cloud.example/remote.php/dav/addressbooks/users/stefan/contacts/";
    private static int checks;
    private static void Check(bool value, string name)
    {
        if (!value) throw new Exception(name);
        checks++;
    }
    private static string Entry(string href, string status)
    {
        return "<d:response><d:href>" + href + "</d:href><d:propstat><d:prop><d:getetag>&quot;v1&quot;</d:getetag></d:prop><d:status>HTTP/1.1 "
            + status + "</d:status></d:propstat></d:response>";
    }
    private static CardDavRemoteSnapshot Parse(string contents)
    {
        return CardDavRemoteSnapshot.Parse(Book, XDocument.Parse("<d:multistatus xmlns:d='DAV:'>" + contents + "</d:multistatus>"));
    }
    private static void Reject(string contents, string name)
    {
        bool rejected = false;
        try { Parse(contents); }
        catch (InvalidOperationException) { rejected = true; }
        Check(rejected, name);
    }
    public static void Main()
    {
        var snapshot = Parse(Entry(Book + "keep.vcf", "200 OK"));
        Check(!snapshot.IsMissing(Book + "keep.vcf"), "Unchanged remote contact survives");
        Check(snapshot.IsMissing(Book + "deleted.vcf"), "Deleted managed contact is reconciled");
        Check(!snapshot.IsMissing(""), "Unmanaged local contact survives");
        Check(!snapshot.IsMissing("not-a-uri"), "Malformed marker survives");
        Check(!snapshot.IsMissing(Book.Replace("/contacts/", "/disabled/") + "deleted.vcf"), "Disabled address book survives");
        Check(!snapshot.IsMissing(Book.Replace("/stefan/", "/other/") + "deleted.vcf"), "Other account survives");
        Check(!snapshot.IsMissing(Book.Replace("cloud.example", "other.example") + "deleted.vcf"), "Other server survives");
        Check(!snapshot.IsMissing(Book.TrimEnd('/') + "-other/deleted.vcf"), "Prefix collision survives");
        Check(!snapshot.IsMissing(Book + "nested/deleted.vcf"), "Nested collection survives");
        Check(!snapshot.IsMissing(Book + "deleted.vcf?query=1"), "Ambiguous query survives");
        Check(Parse("").IsMissing(Book + "last.vcf"), "Deleting last remote contact works");
        Check(!Parse(Entry("keep.vcf", "200 OK")).IsMissing(Book + "keep.vcf"), "Relative href resolves");
        Check(!Parse(Entry("/remote.php/dav/addressbooks/users/stefan/contacts/keep.vcf", "200 OK")).IsMissing(Book + "keep.vcf"), "Root relative href resolves");
        Check(!Parse(Entry(Book + "a%20b.vcf", "200 OK")).IsMissing(Book + "a b.vcf"), "Escaped href normalizes");
        Reject(Entry(Book + "keep.vcf", "403 Forbidden"), "Per-resource failure blocks deletion");
        Reject("<d:error><d:number-of-matches-within-limits/></d:error>", "Truncated listing blocks deletion");
        Reject("<d:response><d:href>" + Book + "a.vcf</d:href><d:status>HTTP/1.1 404 Not Found</d:status></d:response>", "Missing response blocks deletion");
        Reject("<d:response><d:href>" + Book + "a.vcf</d:href></d:response>", "Missing properties block deletion");
        Reject(Entry("", "200 OK"), "Missing href blocks deletion");
        Reject(Entry("https://other.example/a.vcf", "200 OK"), "Foreign href blocks deletion");
        Reject(Entry(Book + "a.vcf", "200 OK") + Entry(Book + "a.vcf", "200 OK"), "Duplicate href blocks deletion");
        Reject("<d:next>page2</d:next>", "Unexpected pagination blocks deletion");
        bool invalidRoot = false;
        try { CardDavRemoteSnapshot.Parse(Book, XDocument.Parse("<html/>")); }
        catch (InvalidOperationException) { invalidRoot = true; }
        Check(invalidRoot, "Non-DAV response blocks deletion");
        VCardChecks();
        FolderNameChecks();
        Console.WriteLine("CardDAV safety checks passed: " + checks);
    }
    private static void FolderNameChecks()
    {
        Check(NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Matches("Sifa (Nur dieser Computer)", "Sifa"), "Local-only suffix names the same folder");
        Check(NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Matches("kunden (This computer only)", "Kunden"), "English local-only suffix is ignored");
        Check(!NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Matches("Sifa (Nextcloud)", "Sifa"), "Collision folder stays a different folder");
        Check(NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Normalize("IBP - Firmenverzeichnis (Nur dieser Computer)").EndsWith(" - Firmenverzeichnis"), "Destination suffix survives normalization");
    }

    private static void VCardChecks()
    {
        string folded = CardDavVCardSupport.Normalize("BEGIN:VCARD\r\nNOTE:Erster Teil\r\n  Termin am 01.02. um 10:30\r\nEND:VCARD");
        Check(folded.Contains("\r\n  Termin am 01.02. um 10:30\r\n"), "Folded continuation line stays intact");
        string grouped = CardDavVCardSupport.Normalize("item1.TEL:+49 1\r\nitem1.X-ABLabel:_$!<Mobile>!$_\r\nitem2.EMAIL;TYPE=WORK:mailto:a@b.example");
        Check(grouped.Contains("item1.TEL:+49 1") && grouped.Contains("item1.X-ABLabel:"), "Group prefixes survive normalization");
        Check(grouped.Contains("item2.EMAIL;TYPE=WORK:a@b.example"), "Grouped mailto prefix is removed");

        string value;
        string uri;
        string jpeg = Convert.ToBase64String(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46 });
        Check(jpeg.StartsWith("/9j/"), "JPEG base64 fixture starts with a slash");
        byte[] photo = CardDavVCardSupport.TryExtractEmbeddedPhoto("PHOTO;ENCODING=b;TYPE=JPEG:" + jpeg, out value, out uri);
        Check(CardDavVCardSupport.IsSupportedImage(photo) && uri.Length == 0, "Embedded JPEG is decoded, not fetched");
        photo = CardDavVCardSupport.TryExtractEmbeddedPhoto("PHOTO;ENCODING=b:" + jpeg.Substring(0, 4) + "\r\n " + jpeg.Substring(4), out value, out uri);
        Check(CardDavVCardSupport.IsSupportedImage(photo), "Folded embedded JPEG is decoded");
        photo = CardDavVCardSupport.TryExtractEmbeddedPhoto("PHOTO:data:image/jpeg;base64," + jpeg, out value, out uri);
        Check(CardDavVCardSupport.IsSupportedImage(photo) && uri.Length == 0, "Data URI JPEG is decoded");
        photo = CardDavVCardSupport.TryExtractEmbeddedPhoto("PHOTO;VALUE=uri:https://cloud.example/p.jpg", out value, out uri);
        Check(photo == null && uri == "https://cloud.example/p.jpg", "Absolute photo URI is kept as URI");
        photo = CardDavVCardSupport.TryExtractEmbeddedPhoto("PHOTO;VALUE=uri:/remote.php/dav/p.jpg", out value, out uri);
        Check(photo == null && uri == "/remote.php/dav/p.jpg", "Relative photo URI is kept as URI");

        var configuration = new TalkServiceConfiguration { BaseUrl = "https://cloud.example/nextcloud" };
        Check(CardDavVCardSupport.ResolvePhotoUrl("https://cloud.example/nextcloud/p.jpg", configuration) == "https://cloud.example/nextcloud/p.jpg", "Same-origin photo URL is accepted");
        Check(CardDavVCardSupport.ResolvePhotoUrl("/remote.php/dav/p.jpg", configuration) == "https://cloud.example/nextcloud/remote.php/dav/p.jpg", "Relative photo URL resolves below base path");
        Check(CardDavVCardSupport.ResolvePhotoUrl("https://attacker.example/p.jpg", configuration) == "", "Foreign photo host is rejected");
        Check(CardDavVCardSupport.ResolvePhotoUrl("http://cloud.example/nextcloud/p.jpg", configuration) == "", "Plain HTTP photo URL is rejected");
        Check(CardDavVCardSupport.ResolvePhotoUrl("https://cloud.example:8443/p.jpg", configuration) == "", "Other port is rejected");
        Check(CardDavVCardSupport.ResolvePhotoUrl("https://user@cloud.example/p.jpg", configuration) == "", "User info in photo URL is rejected");
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    $exe = Join-Path $TempRoot "CardDavTests.exe"
    $snapshotSourcePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\CardDavRemoteSnapshot.cs"
    $snapshotSource = Get-Content -LiteralPath $snapshotSourcePath -Raw
    $snapshotClass = [regex]::Match(
        $snapshotSource,
        '(?s)    internal sealed class CardDavRemoteSnapshot.*?(?=\r?\n    internal static class CardDavContactFolderSync)')
    if (-not $snapshotClass.Success) {
        throw "CardDavRemoteSnapshot class could not be isolated for the safety test."
    }
    $snapshotTestSource = Join-Path $TempRoot "CardDavRemoteSnapshot.TestSource.cs"
    @(
        "using System;"
        "using System.Collections.Generic;"
        "using System.Linq;"
        "using System.Xml.Linq;"
        ""
        "namespace NcTalkOutlookAddIn.Services"
        "{"
        $snapshotClass.Value
        "}"
    ) | Set-Content -Path $snapshotTestSource -Encoding UTF8
    $vcardSourcePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\CardDavVCardSupport.cs"
    $vcardSource = Get-Content -LiteralPath $vcardSourcePath -Raw
    $vcardMethods = foreach ($methodName in @("Normalize", "TryExtractEmbeddedPhoto", "TryDecodeBase64", "ResolvePhotoUrl", "LooksLikePhotoUri", "IsSupportedImage")) {
        $method = [regex]::Match($vcardSource, "(?ms)^        (?:private|internal) static [^\r\n]*? $methodName\(.*?^        }")
        if (-not $method.Success) {
            throw "CardDavVCardSupport.$methodName could not be isolated for the vCard test."
        }
        $method.Value -replace '^        private static', '        internal static'
    }
    $vcardTestSource = Join-Path $TempRoot "CardDavVCardSupport.TestSource.cs"
    @(
        "using System;"
        "using System.Text;"
        "using NcTalkOutlookAddIn.Utilities;"
        ""
        "namespace NcTalkOutlookAddIn.Services"
        "{"
        "    internal sealed class TalkServiceConfiguration { internal string BaseUrl { get; set; } }"
        "    internal static class CardDavVCardSupport"
        "    {"
        $vcardMethods
        "    }"
        "}"
    ) | Set-Content -Path $vcardTestSource -Encoding UTF8
    $folderNamesSource = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\OutlookFolderNames.cs"
    $uriValidatorSource = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudUriValidator.cs"
    & $csc /nologo /target:exe "/out:$exe" /r:System.Core.dll /r:System.Xml.Linq.dll $testSource $snapshotTestSource $vcardTestSource $folderNamesSource $uriValidatorSource
    if ($LASTEXITCODE -ne 0) { throw "CardDAV test compilation failed." }
    & $exe
    if ($LASTEXITCODE -ne 0) { throw "CardDAV safety tests failed." }
}
finally {
    Remove-Item -Recurse -Force $TempRoot
}
