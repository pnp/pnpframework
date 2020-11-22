using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Publishing;
using PnP.Framework.Modernization.Telemetry.Observers;

namespace PnP.Framework.Modernization.Tests.Transform.Publishing
{
    [TestClass]
    public class PublishingPageTests
    {
        #region Test initialization
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            //using (var cc = TestCommon.CreateClientContext())
            //{
            //    // Clean all migrated pages before next run
            //    var pages = cc.Web.GetPages("Migrated_");

            //    foreach (var page in pages.ToList())
            //    {
            //        page.DeleteObject();
            //    }

            //    cc.ExecuteQueryRetry();
            //}
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {

        }
        #endregion

        [TestMethod]
        public void BasicPublishingPageTest()
        {
            using (var targetClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                //https://bertonline.sharepoint.com/sites/modernizationtestportal
                using (var sourceClientContext = TestCommon.CreateClientContext("https://bertonline.sharepoint.com/sites/devportal/en-us"))
                {
                    //"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Publishing\custompagelayoutmapping.xml"
                    //"C:\temp\mappingtest.xml"
                    //@"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Publishing\webpartmapping.xml"
                    //var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext, @"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Publishing\webpartmapping.xml", @"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Publishing\custompagelayoutmapping.xml");
                    var pageTransformator = new PublishingPageTransformator(sourceClientContext, targetClientContext , @"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Publishing\custompagelayoutmapping.xml");
                    pageTransformator.RegisterObserver(new MarkdownObserver(folder: "c:\\temp", includeVerbose:true));
                    //pageTransformator.RegisterObserver(new MarkdownToSharePointObserver(targetClientContext, includeVerbose: true));

                    var pages = sourceClientContext.Web.GetPagesFromList("Pages", "kevin");
                    //var pages = sourceClientContext.Web.GetPagesFromList("Pages", folder:"News");

                    foreach (var page in pages)
                    {
                        PublishingPageTransformationInformation pti = new PublishingPageTransformationInformation(page)
                        {
                            // If target page exists, then overwrite it
                            Overwrite = true,

                            // Don't log test runs
                            SkipTelemetry = true,  

                            KeepPageCreationModificationInformation = true,
                            
                            PostAsNews = true,

                            //TermMappingFile = @"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Mapping\termmapping_sample2.csv",

                            //UrlMappingFile = @"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Mapping\urlmapping_sample.csv",

                            //UserMappingFile = @"C:\github\sp-dev-modernization\Tools\SharePoint.Modernization\PnP.Framework.Modernization.Tests\Transform\Mapping\usermapping_sample2.csv",

                            //DisablePageComments = true,

                            //PublishCreatedPage = true,

                            //RemoveEmptySectionsAndColumns = false,

                            // Configure the page header, empty value means ClientSidePageHeaderType.None
                            //PageHeader = new ClientSidePageHeader(cc, ClientSidePageHeaderType.None, null),

                            // Replace embedded images and iframes with a placeholder and add respective images and video web parts at the bottom of the page
                            // HandleWikiImagesAndVideos = false,

                            // Callout to your custom code to allow for title overriding
                            //PageTitleOverride = titleOverride,

                            // Callout to your custom layout handler
                            //LayoutTransformatorOverride = layoutOverride,

                            // Callout to your custom content transformator...in case you fully want replace the model
                            //ContentTransformatorOverride = contentOverride,
                        };

                        pti.MappingProperties["SummaryLinksToQuickLinks"] = "true";
                        pti.MappingProperties["UseCommunityScriptEditor"] = "true";

                        var result = pageTransformator.Transform(pti);
                    }

                    pageTransformator.FlushObservers();
                }
            }
        }

        [TestMethod]
        public void PageLayout_AnalyzeByPages_Test()
        {
            //https://bertonline.sharepoint.com/sites/modernizationtestportal
            using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
            {
                var pages = sourceClientContext.Web.GetPagesFromList("Pages");
                var analyzer = new PageLayoutAnalyser(sourceClientContext);

                foreach (var page in pages)
                {
                    analyzer.AnalysePageLayoutFromPublishingPage(page);    
                }

                analyzer.GenerateMappingFile("c:\\temp", "mappingtest.xml");
            }
        }

        [TestMethod]
        public void PageLayout_AnalyseAll_Test()
        {
            //https://bertonline.sharepoint.com/sites/modernizationtestportal
            using (var sourceClientContext = TestCommon.CreateClientContext(TestCommon.AppSetting("SPODevSiteUrl")))
            {
                
                var analyzer = new PageLayoutAnalyser(sourceClientContext);
                analyzer.AnalyseAll();                

                analyzer.GenerateMappingFile("c:\\temp", "mappingalltest.xml");
            }
        }


    }
}
