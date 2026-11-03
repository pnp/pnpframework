## Getting Started with PnP Page Transformation

The PnP Page Transformation tooling is a set of tools and libraries that enable the transformation of classic SharePoint pages to modern client-side pages. 

The transformation process involves analyzing the structure and content of classic pages, and then generating modern client-side pages that replicate the structure and content of the classic pages. In general, site and list configuration must be prepared separately. Enterprise Wiki publishing-page transformation also prepares the destination prerequisites described below.

Master pages and page layouts are not converted but they are mapped to their closest equivalent in the modern page model.


## What versions of SharePoint are supported?

The PnP Page Transformation tooling supports SharePoint Online and SharePoint 2019, 2016 and 2013. It is not supported for SharePoint 2010 or earlier versions. However you can an older version of the tooling to transform pages in SharePoint 2010.

You can convert from SharePoint Online classic to SharePoint Online modern pages. 

## Example of transformation

The following example shows a classic SharePoint page and the modern page that was generated from it. 

```csharp

// Converts a publishing page to modern page example

using (var targetClientContext = GetClientContext("https://<tenant>.sharepoint.com/sites/modernsite"))
{
    using (var sourceClientContext = GetClientContext("https://<tenant>.sharepoint.com/sites/classic-site"))
    {
        
        var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext , @"path\to\mapping\custom-page-layout-mapping.xml");
        // pageTransformator.RegisterObserver(new MarkdownObserver(folder: "d:\\temp", includeVerbose:true));
        
        var pages = sourceClientContext.Web.GetPagesFromList("Pages", "");
        
        foreach (var page in pages)
        {
            // Options for transformation
            PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
            {
                // If target page exists, then overwrite it
                Overwrite = true,
                KeepPageCreationModificationInformation = true,
                PostAsNews = true,
                TermMappingFile = @"path\to\mapping\term_mapping.csv",
                UrlMappingFile = @"path\to\mapping\url_mapping.csv",
                UserMappingFile = @"path\to\mapping\user_mapping.csv",
                DisablePageComments = true,                
                PublishCreatedPage = true,
            };

            pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
            pti.MappingProperties["UseCommunityScriptEditor"] = "true";

            // Transform the page
            var result = pageTransformator.Transform(pti);
        }

        // Writes output to logs
        pageTransformator.FlushObservers();
    }
}

```


## Same-site wiki and web part pages

When a supplied target context belongs to the source site collection,
`PageTransformator` uses source-only (in-place) mode. The destination is the
source web's `SitePages` library, even when the supplied target refers to a
different subweb in that collection. Cross-site flags and target-version
metadata are reset, but `TargetPageFolder`, its override option, and an explicit
target page name are retained. Blog and publishing pages have separate rules.

For a custom source library, set `TargetPageFolderOverridesDefaultFolder = true`
to place the page directly in the requested folder under `SitePages`; otherwise
the source library/folder path is included. Do not prefix `TargetPageFolder`
with `SitePages`.

To verify a local framework change through PnP.PowerShell, build that module
against the local framework using its `PnPFrameworkPath` build property, import
the isolated module in a new PowerShell process, and confirm the loaded
`PnP.Framework` assembly path and file hash match the build. An installed
module can contain an older framework even when this repository builds.

Use an authorized test web, an existing classic page in a custom library, and
unique output folders. Run `ConvertTo-PnPPage` with no target parameters,
`-TargetConnection` only, `-TargetWebUrl` only, and both parameters pointing to
the same web. Retain `-TargetPageFolder` and
`-TargetPageFolderOverridesDefaultFolder` in each run. Verify the resulting
`SitePages` item has modern page metadata and canvas content, the source page
version is unchanged, and no critical transformation error was reported.
Remove only the output folders created by the test. A successful command
invocation alone is insufficient: older versions can report a transformation
error without returning a terminating PowerShell error.

## Table styles

Plain tables and the classic `ms-rteTable-0` ("no grid") style are transformed
using the modern editor's native borderless table style. Classic style classes
are recognized even when a table has additional CSS classes. Existing grid,
header, and banded style mappings are retained.

Visible inline borders on a table, row, or cell, or a positive HTML `border`
attribute, prevent a borderless mapping. Borders belonging to a nested table do
not affect the containing table. An explicit borderless declaration on a table
with an unrecognized class is honored only when there are no visible inline
borders; otherwise unrecognized classes retain the existing bordered fallback.
External stylesheet rules and exact custom border colors, widths, or patterns
are not reproduced.

These mappings apply when transforming a page; previously transformed pages
are not automatically updated.

## Enterprise Wiki pages

Set `EnableEnterpriseWikiCapabilities = true` to prepare SitePages using existing category/rating site columns and copy their values. When false or omitted, all destinations retain legacy transformation behavior and custom mappings, without ESWiki preparation or sidebar insertion. Inconsistent rating history stops migration only with the new grouped rating handling enabled.

When opted in, the sidebar requires destination support for the native modernized-control host. Set `IncludeEnterpriseWikiSidebar = false` where unavailable.

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework/pnp-modernization" aria-hidden="true" />