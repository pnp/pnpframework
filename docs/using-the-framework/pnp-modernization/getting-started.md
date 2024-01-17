## Getting Started with PnP Page Transformation

The PnP Page Transformation tooling is a set of tools and libraries that enable the transformation of classic SharePoint pages to modern client-side pages. 

The transformation process involves analyzing the structure and content of classic pages, and then generating modern client-side pages that replicate the structure and content of the classic pages. This does not include the configuration of the SharePoint Sites or lists, only the page contents.

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


> Working on it... more to come.

<img src="https://m365-visitor-stats.azurewebsites.net/pnpframework/docs/using-the-framework/pnp-modernization" aria-hidden="true" />