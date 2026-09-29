Param([string]$ProjectRoot = ".")
$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("nc4ol-carddav-tests-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null
try {
    $testSource = Join-Path $TempRoot "CardDavTests.cs"
    @'
using System;
using System.Collections.Generic;
using System.Linq;
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
        CustomerPlannerChecks();
        CustomerVCardChecks();
        FolderNameChecks();
        Console.WriteLine("CardDAV safety checks passed: " + checks);
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

    private const string CustomerBook = "https://cloud.example/remote.php/dav/addressbooks/users/stefan/ibp-kunden/";
    private static CardDavCustomerLocalEntry Local(string href, string etag, bool changed)
    {
        return new CardDavCustomerLocalEntry
        {
            EntryId = Guid.NewGuid().ToString("N"),
            Uid = "uid-" + Guid.NewGuid().ToString("N"),
            Href = href,
            StoredEtag = etag,
            StoredHash = "h1",
            CurrentHash = changed ? "h2" : "h1"
        };
    }
    private static CardDavCustomerAction Single(CardDavCustomerPlan plan, CardDavCustomerLocalEntry local)
    {
        return plan.Actions.SingleOrDefault(a => a.Local == local);
    }
    private static CardDavCustomerPlan Plan(IEnumerable<CardDavCustomerLocalEntry> locals, Dictionary<string, string> remote, params string[] known)
    {
        return CardDavCustomerPlanner.Plan(locals, remote, known, CustomerBook);
    }
    private static void CustomerPlannerChecks()
    {
        string a = CustomerBook + "a.vcf";
        string b = CustomerBook + "b.vcf";
        var remote = new Dictionary<string, string> { { a, "\"e1\"" } };

        var fresh = Local("", "", true);
        CardDavCustomerAction action = Single(Plan(new[] { fresh }, remote), fresh);
        Check(action.Kind == CardDavCustomerActionKind.UploadNew && CardDavCustomerPlanner.BelongsTo(CustomerBook, action.Href), "New Outlook customer is uploaded into the customer book");

        var unchanged = Local(a, "\"e1\"", false);
        CardDavCustomerPlan plan = Plan(new[] { unchanged }, remote, a);
        Check(plan.Actions.Count == 0, "Unchanged customer needs no action");
        Check(Plan(new[] { Local(a, "W/\"e1\"", false) }, remote, a).Actions.Count == 0, "Weak ETag marker compares equal");

        var remoteEdit = Local(a, "\"e0\"", false);
        Check(Single(Plan(new[] { remoteEdit }, remote, a), remoteEdit).Kind == CardDavCustomerActionKind.DownloadUpdate, "Cloud edit is downloaded");

        var bothEdit = Local(a, "\"e0\"", true);
        action = Single(Plan(new[] { bothEdit }, remote, a), bothEdit);
        Check(action.Kind == CardDavCustomerActionKind.UploadUpdate && action.RemoteEtag == "\"e1\"", "Outlook wins a two-sided edit with If-Match");

        var cloudDeleted = Local(b, "\"e1\"", false);
        Check(Single(Plan(new[] { cloudDeleted }, remote, a, b), cloudDeleted).Kind == CardDavCustomerActionKind.DeleteLocal, "Cloud deletion of a synced customer deletes it in Outlook");
        var cloudDeletedButEdited = Local(b, "\"e1\"", true);
        action = Single(Plan(new[] { cloudDeletedButEdited }, remote, a, b), cloudDeletedButEdited);
        Check(action.Kind == CardDavCustomerActionKind.UploadNew && action.Href == b, "Outlook edit wins over a cloud deletion");
        var neverConfirmed = Local(b, "\"e1\"", false);
        Check(Single(Plan(new[] { neverConfirmed }, remote), neverConfirmed).Kind == CardDavCustomerActionKind.UploadNew, "Missing but unknown href is uploaded, not deleted");

        plan = Plan(new CardDavCustomerLocalEntry[0], remote, a);
        Check(plan.Actions.Single().Kind == CardDavCustomerActionKind.DeleteRemote, "Outlook deletion of a synced customer deletes it in the cloud");
        plan = Plan(new CardDavCustomerLocalEntry[0], remote);
        Check(plan.Actions.Single().Kind == CardDavCustomerActionKind.DownloadNew, "New cloud customer is downloaded");

        var original = Local(a, "\"e1\"", false);
        var copy = Local(a, "\"e1\"", false);
        plan = Plan(new[] { original, copy }, remote, a);
        Check(Single(plan, original) == null && Single(plan, copy).Kind == CardDavCustomerActionKind.UploadNew && Single(plan, copy).Href != a, "Copied Outlook contact becomes a new customer");
        var foreign = Local("https://cloud.example/remote.php/dav/addressbooks/users/stefan/contacts/x.vcf", "\"e1\"", false);
        action = Single(Plan(new[] { foreign }, remote, a), foreign);
        Check(action.Kind == CardDavCustomerActionKind.UploadNew && CardDavCustomerPlanner.BelongsTo(CustomerBook, action.Href), "Foreign CardDAV marker gets a new customer href");
        Check(!CardDavCustomerPlanner.BelongsTo(CustomerBook, CustomerBook.TrimEnd('/') + "-old/a.vcf"), "Prefix collision is not the customer book");

        var many = new Dictionary<string, string>();
        var knownHrefs = new List<string>();
        for (int i = 0; i < 10; i++)
        {
            string href = CustomerBook + "m" + i + ".vcf";
            many[href] = "\"e\"";
            knownHrefs.Add(href);
        }
        plan = Plan(new CardDavCustomerLocalEntry[0], many, knownHrefs.ToArray());
        Check(plan.DeletionsBlocked && plan.Actions.Count == 0 && plan.BlockedDeletionHrefs.Count == 10, "Emptied Outlook folder does not wipe the cloud book");
        var locals = knownHrefs.Skip(6).Select(h => Local(h, "\"e\"", false)).ToList();
        plan = Plan(locals, many, knownHrefs.ToArray());
        Check(plan.DeletionsBlocked, "Deleting more than half of many customers is held back");
        locals = knownHrefs.Skip(4).Select(h => Local(h, "\"e\"", false)).ToList();
        plan = Plan(locals, many, knownHrefs.ToArray());
        Check(!plan.DeletionsBlocked && plan.Actions.Count(x => x.Kind == CardDavCustomerActionKind.DeleteRemote) == 4, "Few deletions pass the mass deletion guard");
    }

    private static void FolderNameChecks()
    {
        Check(NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Matches("Sifa (Nur dieser Computer)", "Sifa"), "Local-only suffix names the same folder");
        Check(NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Matches("kunden (This computer only)", "Kunden"), "English local-only suffix is ignored");
        Check(!NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Matches("Sifa (Nextcloud)", "Sifa"), "Collision folder stays a different folder");
        Check(NcTalkOutlookAddIn.Utilities.OutlookFolderNames.Normalize("IBP - Firmenverzeichnis (Nur dieser Computer)").EndsWith(" - Firmenverzeichnis"), "Destination suffix survives normalization");
    }

    private static void CustomerVCardChecks()
    {
        var record = new CardDavContactRecord
        {
            FullName = "Erika Mustermann",
            FirstName = "Erika",
            LastName = "Mustermann",
            Company = "Muster GmbH; Co, KG",
            JobTitle = "Einkauf",
            Email1 = "erika@muster.example",
            Email2 = "privat@muster.example",
            BusinessPhone = "+49 30 1",
            MobilePhone = "+49 170 2",
            HomePhone = "+49 30 3",
            BusinessFax = "+49 30 4",
            OtherPhone = "+49 30 5",
            BusinessAddressStreet = "Hauptstraße 1",
            BusinessAddressCity = "Berlin",
            BusinessAddressPostalCode = "10115",
            BusinessAddressCountry = "Deutschland",
            Notes = "Zeile 1\r\nZeile 2 mit einem sehr langen Text, der über die Grenze von fünfundsiebzig Zeichen hinausgeht."
        };
        DateTime now = new DateTime(2026, 9, 29, 8, 0, 0, DateTimeKind.Utc);
        string built = CardDavCustomerVCard.Build(record, "uid-1", null, now);
        Check(built.StartsWith("BEGIN:VCARD\r\nVERSION:3.0\r\n") && built.EndsWith("END:VCARD\r\n"), "Customer vCard is a CRLF vCard 3.0");
        Check(built.Contains("UID:uid-1\r\n") && built.Contains("ORG:Muster GmbH\\; Co\\, KG\r\n"), "Customer vCard escapes text values");
        Check(built.Split(new[] { "\r\n" }, StringSplitOptions.None).All(line => line.Length <= 75), "Customer vCard lines are folded");

        CardDavContactRecord parsed = CardDavReadOnlySync.ParseVCard(CardDavVCardSupport.Normalize(built));
        Check(CardDavCustomerVCard.ComputeHash(parsed) == CardDavCustomerVCard.ComputeHash(record), "Customer vCard round-trips all mapped fields");
        Check(parsed.OtherPhone == "+49 30 5" && string.IsNullOrEmpty(parsed.BusinessPhone2), "Other phone stays other phone");

        string existing = "BEGIN:VCARD\r\nVERSION:3.0\r\nUID:uid-1\r\nFN:Alt\r\nN:Alt;;;;\r\nTEL;TYPE=WORK:+49 1\r\n"
            + "item1.TEL:+49 9\r\nitem1.X-ABLabel:Werk\r\nitem2.URL:https://muster.example\r\nitem2.X-ABLabel:Web\r\n"
            + "EMAIL:alt@muster.example\r\nADR;TYPE=WORK:;;Altweg 1;Alt;;;\r\nADR;TYPE=HOME:;;Heimweg 2;Potsdam;;14467;\r\n"
            + "BDAY:1980-01-01\r\nCATEGORIES:VIP\r\nPHOTO;ENCODING=b;TYPE=JPEG:/9j/4AAQ\r\n SkZJRg==\r\nREV:20200101T000000Z\r\nEND:VCARD\r\n";
        string merged = CardDavCustomerVCard.Build(record, "uid-1", existing, now);
        string unfolded = merged.Replace("\r\n ", "");
        string[] mergedLines = unfolded.Split(new[] { "\r\n" }, StringSplitOptions.None);
        Check(unfolded.Contains("BDAY:1980-01-01") && unfolded.Contains("CATEGORIES:VIP") && unfolded.Contains("PHOTO;ENCODING=b;TYPE=JPEG:/9j/4AAQSkZJRg=="), "Upload keeps unmapped cloud properties");
        Check(unfolded.Contains("ADR;TYPE=HOME:;;Heimweg 2") && !unfolded.Contains("Altweg"), "Upload replaces only the business address");
        Check(unfolded.Contains("item2.URL:") && unfolded.Contains("item2.X-ABLabel:Web") && !unfolded.Contains("item1."), "Grouped labels follow their property");
        Check(!mergedLines.Contains("TEL;TYPE=WORK:+49 1") && !unfolded.Contains("alt@muster") && !mergedLines.Contains("FN:Alt"), "Upload replaces mapped cloud properties");
        Check(mergedLines.Count(l => l == "BEGIN:VCARD") == 1 && mergedLines.Count(l => l.StartsWith("REV:")) == 1, "Merged vCard has one envelope and one REV");
        parsed = CardDavReadOnlySync.ParseVCard(CardDavVCardSupport.Normalize(merged));
        Check(parsed.BusinessAddressStreet == "Hauptstraße 1", "Business address wins over a preserved home address");
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
    $syncSourcePath = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\CardDavReadOnlySync.cs"
    $syncSource = Get-Content -LiteralPath $syncSourcePath -Raw
    $recordClass = [regex]::Match($syncSource, '(?ms)^    internal sealed class CardDavContactRecord.*?^    }')
    if (-not $recordClass.Success) {
        throw "CardDavContactRecord could not be isolated for the customer test."
    }
    $parserMethods = foreach ($methodName in @("ParseVCard", "ReadVCardGroupLabels", "NormalizeVCardPropertyName", "GetVCardPropertyGroup", "GetVCardPropertyToken", "NormalizeAppleVCardLabel", "AssignTelephone", "TryAssignPhoneValue", "UnfoldVCard", "IsQuotedPrintableProperty", "DecodeVCardValue", "DecodeQuotedPrintableValue", "DecodeQuotedPrintableUtf8", "HexValue", "SplitVCardComponents", "DecodeVCardText")) {
        $method = [regex]::Match($syncSource, "(?ms)^        (?:private|internal) static [^\r\n]*? $methodName\(.*?^        }")
        if (-not $method.Success) {
            throw "CardDavReadOnlySync.$methodName could not be isolated for the customer test."
        }
        $method.Value -replace '^        private static', '        internal static'
    }
    $parserTestSource = Join-Path $TempRoot "CardDavParser.TestSource.cs"
    @(
        "using System;"
        "using System.Collections.Generic;"
        "using System.Text;"
        ""
        "namespace NcTalkOutlookAddIn.Services"
        "{"
        $recordClass.Value
        "    internal static class CardDavReadOnlySync"
        "    {"
        $parserMethods
        "    }"
        "}"
    ) | Set-Content -Path $parserTestSource -Encoding UTF8
    $plannerSource = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\CardDavCustomerPlanner.cs"
    $folderNamesSource = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\OutlookFolderNames.cs"
    $uriValidatorSource = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Utilities\NextcloudUriValidator.cs"
    & $csc /nologo /target:exe "/out:$exe" /r:System.Core.dll /r:System.Xml.Linq.dll $testSource $snapshotTestSource $vcardTestSource $parserTestSource $plannerSource $folderNamesSource $uriValidatorSource
    if ($LASTEXITCODE -ne 0) { throw "CardDAV test compilation failed." }
    & $exe
    if ($LASTEXITCODE -ne 0) { throw "CardDAV safety tests failed." }
}
finally {
    Remove-Item -Recurse -Force $TempRoot
}
