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


## Page headers and authors

Wiki and web-part page transformation now creates a no-image header by default and attempts to display the source page's **Created By** user and available job title. `SetAuthorInPageHeader` remains available for compatibility but is ignored: omitted, `true`, and `false` all use this default. This also applies to the switch passed through by `ConvertTo-PnPPage` when it uses this Framework version. Explicit default/custom `PageHeader` configurations supplied through the .NET API retain their existing behavior.

Converted source home pages use **Article** layout because Home layout does not persist a page header. Setting `ReplaceHomePageWithDefaultHomePage = true` (the `-ReplaceHomePageWithDefault` cmdlet option) still keeps the stock modern home-page layout/content and skips normal page conversion and author-header creation.

For publishing pages, layout mappings with `PageHeader="None"` or `PageHeader="Default"` now produce a no-image header using Created By. `PageHeader="Custom"` retains its image/layout settings and Authors mapping (normally `ToAuthors({PublishingContact})`); an empty custom mapping does not fall back to Created By.

Author identities are resolved in the destination site, and the header's display data and person byline are set together for a resolvable single author. Job titles come from the existing User Information List lookup; no profile refresh is performed. Missing authors produce warnings without removing the page title/header. Custom author JSON is preserved, including multi-author payloads; unsupported or unresolved byline identities are reported rather than guessed.

This header attribution is independent of `KeepPageCreationModificationInformation`, which controls preservation of audit fields. Whether a populated job title is visibly rendered depends on the SharePoint client.

**SDK limitation:** with PnP.Core `1.18.58-nightly`, loading a saved no-image header can classify it as `None`. A subsequent SDK `Save` then removes the header and its byline. Initial transformation saves the header correctly, but load/save preservation requires a correction in the PnP Core header parser; it must not be assumed from successful migration alone.

## Enterprise Wiki pages

Set `EnableEnterpriseWikiCapabilities = true` to prepare SitePages using existing category/rating site columns and copy their values. When false or omitted, all destinations retain legacy transformation behavior and custom mappings, without ESWiki preparation or sidebar insertion. Inconsistent rating history stops migration only with the new grouped rating handling enabled.

When opted in, the sidebar requires destination support for the native modernized-control host. Set `IncludeEnterpriseWikiSidebar = false` where unavailable.

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework/pnp-modernization" aria-hidden="true" />