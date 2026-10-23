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


## Web Part page templates

Standard Web Part page layouts include document workspaces (`defaultdws.aspx`), project sites, Document Centers, and basic pages (`bpstd.aspx`), in addition to the existing team-site and `spstd1.aspx` through `spstd8.aspx` layouts. Document workspace and project-site descriptions are copied as text at transformation time.

Template mapping preserves supported zone placement, not all classic page behavior. Dynamic document-workspace breadcrumbs and arbitrary content outside Web Part zones are not converted.

## Enterprise Wiki pages

Set `EnableEnterpriseWikiCapabilities = true` to prepare SitePages using existing category/rating site columns and copy their values. When false or omitted, all destinations retain legacy transformation behavior and custom mappings, without ESWiki preparation or sidebar insertion. Inconsistent rating history stops migration only with the new grouped rating handling enabled.

When opted in, the sidebar requires destination support for the native modernized-control host. Set `IncludeEnterpriseWikiSidebar = false` where unavailable.

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework/pnp-modernization" aria-hidden="true" />