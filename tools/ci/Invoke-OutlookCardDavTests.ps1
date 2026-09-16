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
        Console.WriteLine("CardDAV safety checks passed: " + checks);
    }
}
'@ | Set-Content -Path $testSource -Encoding UTF8
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    $exe = Join-Path $TempRoot "CardDavTests.exe"
    $snapshotSource = Join-Path $ProjectRoot "src\NcTalkOutlookAddIn\Services\CardDavRemoteSnapshot.cs"
    & $csc /nologo /target:exe "/out:$exe" /r:System.Core.dll /r:System.Xml.Linq.dll $testSource $snapshotSource
    if ($LASTEXITCODE -ne 0) { throw "CardDAV test compilation failed." }
    & $exe
    if ($LASTEXITCODE -ne 0) { throw "CardDAV safety tests failed." }
}
finally {
    Remove-Item -Recurse -Force $TempRoot
}
