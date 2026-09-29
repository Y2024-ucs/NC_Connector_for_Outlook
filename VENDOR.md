# Third-Party Dependencies

## HtmlSanitizer

- Package: `HtmlSanitizer`
- Version: `9.0.892`
- Source: https://www.nuget.org/packages/HtmlSanitizer/9.0.892
- Upstream repository: https://github.com/mganss/HtmlSanitizer
- Included file: `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/HtmlSanitizer.dll`
- License: MIT
- Usage in this add-in:
  - Sanitization of backend-provided Share/Talk HTML templates
  - `template` elements are additionally excluded by the add-in policy as defense in depth
  - Runtime consumers:
    - `src/NcTalkOutlookAddIn/Utilities/HtmlTemplateSanitizer.cs`
    - `src/NcTalkOutlookAddIn/Utilities/FileLinkHtmlBuilder.cs`
    - `src/NcTalkOutlookAddIn/Controllers/TalkDescriptionTemplateController.cs`

## AngleSharp

- Package: `AngleSharp`
- Version: `0.17.1`
- Source: https://www.nuget.org/packages/AngleSharp/0.17.1
- Upstream repository: https://github.com/AngleSharp/AngleSharp
- Included file: `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/AngleSharp.dll`
- License: MIT
- Usage in this add-in:
  - Runtime dependency of `HtmlSanitizer`

## AngleSharp.Css

- Package: `AngleSharp.Css`
- Version: `0.17.0`
- Source: https://www.nuget.org/packages/AngleSharp.Css/0.17.0
- Upstream repository: https://github.com/AngleSharp/AngleSharp.Css
- Included file: `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/AngleSharp.Css.dll`
- License: MIT
- Usage in this add-in:
  - Runtime dependency of `HtmlSanitizer`

## .NET Runtime Dependencies (vendored with sanitizer stack)

- Packages:
  - `System.Buffers` (`4.6.1` package, `4.0.5.0` assembly)
  - `System.Collections.Immutable` (`10.0.0` package and assembly)
  - `System.Memory` (`4.6.3` package, `4.0.5.0` assembly)
  - `System.Numerics.Vectors` (`4.6.1` package, `4.1.6.0` assembly)
  - `System.Runtime.CompilerServices.Unsafe` (`6.1.2` package, `6.0.3.0` assembly)
  - `System.Text.Encoding.CodePages` (`6.0.0` package)
- Sources:
  - https://www.nuget.org/packages/System.Buffers
  - https://www.nuget.org/packages/System.Collections.Immutable
  - https://www.nuget.org/packages/System.Memory
  - https://www.nuget.org/packages/System.Numerics.Vectors
  - https://www.nuget.org/packages/System.Runtime.CompilerServices.Unsafe
  - https://www.nuget.org/packages/System.Text.Encoding.CodePages
- Included files:
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Buffers.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Collections.Immutable.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Memory.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Numerics.Vectors.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Runtime.CompilerServices.Unsafe.dll`
  - `src/NcTalkOutlookAddIn/vendor/htmlsanitizer/System.Text.Encoding.CodePages.dll`
- License: MIT
- Usage in this add-in:
  - Runtime dependencies required by `HtmlSanitizer`/`AngleSharp`

## Outlook CalDav Synchronizer

- Package: `CalDavSynchronizer.Setup.msi`
- Version: `4.7.1`
- Source: https://github.com/aluxnimm/outlookcaldavsynchronizer/releases/tag/v4.7.1
- Upstream repository: https://github.com/aluxnimm/outlookcaldavsynchronizer
- Included file: `installer/vendor/caldavsynchronizer/CalDavSynchronizer.Setup.msi` (unchanged, SHA-256 pinned in `build.ps1`)
- License: AGPL-3.0
- Usage:
  - Separate Outlook add-in, installed next to NC Connector by the combined setup or the GPO package
  - NC Connector writes its calendar profile: `src/NcTalkOutlookAddIn/Services/CalDavSynchronizerProvisioning.cs`
