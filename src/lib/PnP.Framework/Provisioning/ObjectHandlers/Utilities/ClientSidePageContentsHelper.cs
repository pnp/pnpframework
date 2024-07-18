using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using PnPCore = PnP.Core.Model.SharePoint;


namespace PnP.Framework.Provisioning.ObjectHandlers.Utilities
{
    /// <summary>
    /// Helper class holding public methods that used by the client side page object handler. The purpose is to be able to reuse these public methods in a extensibility provider
    /// </summary>
    public class ClientSidePageContentsHelper
    {

        internal const string PromotedStateField = "PromotedState";
        internal const string SpaceContentField = "SpaceContent";
        internal const string TopicEntityId = "_EntityId";
        internal const string TopicEntityRelations = "_EntityRelations";
        internal const string TopicEntityType = "_EntityType";
        
        private const string ContentTypeIdField = "ContentTypeId";

        public void ExtractClientSidePage(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string pageUrl, string pageName, bool isHomePage, bool isTemplate = false)
        {
            PageToExport page = new PageToExport()
            {
                PageName = pageName,
                PageUrl = pageUrl,
                IsHomePage = isHomePage,
                IsTemplate = isTemplate,
                IsTranslation = false
            };

            ExtractClientSidePage(web, template, creationInfo, scope, page);
        }

        /// <summary>
        /// Extracts a client side page
        /// </summary>
        /// <param name="web">Web to extract the page from</param>
        /// <param name="template">Current provisioning template that will hold the extracted page</param>
        /// <param name="creationInfo">ProvisioningTemplateCreationInformation passed into the provisioning engine</param>
        /// <param name="scope">Scope used for logging</param>
        /// <param name="page">page to be exported</param>
        public void ExtractClientSidePage(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, PageToExport page)
        {
            bool excludeAuthorInformation = false;
            if (creationInfo.ExtractConfiguration != null && creationInfo.ExtractConfiguration.Pages != null)
            {
                excludeAuthorInformation = creationInfo.ExtractConfiguration.Pages.ExcludeAuthorInformation;
            }
            try
            {
                List<string> errorneousOrNonImageFileGuids = new List<string>();
                //var pageToExtract = web.LoadClientSidePage(page.PageName);
                var pageToExtract = web.LoadClientSidePage(page.PageName) as PnPCore.IPage;

                if (pageToExtract.Sections.Count == 0 && pageToExtract.Controls.Count == 0 && page.IsHomePage)
                {
                    // This is default home page which was not customized...and as such there's no page definition stored in the list item. We don't need to extact this page.
                    scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_DefaultHomePage);
                }
                else
                {
                    // Get the page content type
                    string pageContentTypeId = pageToExtract.PageListItem[ContentTypeIdField].ToString();

                    if (!string.IsNullOrEmpty(pageContentTypeId))
                    {
                        pageContentTypeId = GetParentIdValue(pageContentTypeId);
                    }

                    int promotedState = 0;
                    //if (pageToExtract.PageListItem[PnP.Framework.Pages.ClientSidePage.PromotedStateField] != null)
                    if (pageToExtract.PageListItem[PromotedStateField] != null)
                    {
                        //int.TryParse(pageToExtract.PageListItem[PnP.Framework.Pages.ClientSidePage.PromotedStateField].ToString(), out promotedState);
                        int.TryParse(pageToExtract.PageListItem[PromotedStateField].ToString(), out promotedState);
                    }

                    //var isNews = pageToExtract.LayoutType != Pages.ClientSidePageLayoutType.Home && promotedState == (int)Pages.PromotedState.Promoted;
                    var isNews = pageToExtract.LayoutType != PnPCore.PageLayoutType.Home && promotedState == (int)PnPCore.PromotedState.Promoted;

                    // Create the page;
                    BaseClientSidePage extractedPageInstance;
                    if (page.IsTranslation)
                    {
                        extractedPageInstance = new TranslatedClientSidePage();
                        (extractedPageInstance as TranslatedClientSidePage).PageName = page.PageName;
                    }
                    else
                    {
                        extractedPageInstance = new ClientSidePage();
                        (extractedPageInstance as ClientSidePage).PageName = page.PageName;
                    }

                    extractedPageInstance.PromoteAsNewsArticle = isNews;
                    extractedPageInstance.PromoteAsTemplate = page.IsTemplate;
                    extractedPageInstance.Overwrite = true;
                    extractedPageInstance.Publish = true;
                    extractedPageInstance.Layout = pageToExtract.LayoutType.ToString();
                    //extractedPageInstance.EnableComments = !pageToExtract.CommentsDisabled;
                    extractedPageInstance.EnableComments = !pageToExtract.AreCommentsDisabled();
                    extractedPageInstance.Title = pageToExtract.PageTitle;
                    extractedPageInstance.ContentTypeID = !pageContentTypeId.Equals(BuiltInContentTypeId.ModernArticlePage, StringComparison.InvariantCultureIgnoreCase) ? pageContentTypeId : null;
                    extractedPageInstance.ThumbnailUrl = pageToExtract.ThumbnailUrl != null ? TokenizeJsonControlData(web, pageToExtract.ThumbnailUrl) : "";

                    //if (pageToExtract.PageHeader != null && pageToExtract.LayoutType != Pages.ClientSidePageLayoutType.Topic)
                    if (pageToExtract.PageHeader != null && pageToExtract.LayoutType != PnPCore.PageLayoutType.Topic)
                    {

                        var extractedHeader = new ClientSidePageHeader()
                        {
                            //Type = (ClientSidePageHeaderType)Enum.Parse(typeof(Pages.ClientSidePageHeaderType), pageToExtract.PageHeader.Type.ToString()),
                            Type = (ClientSidePageHeaderType)Enum.Parse(typeof(ClientSidePageHeaderType), pageToExtract.PageHeader.Type.ToString()),
                            ServerRelativeImageUrl = TokenizeJsonControlData(web, pageToExtract.PageHeader.ImageServerRelativeUrl),
                            TranslateX = pageToExtract.PageHeader.TranslateX,
                            TranslateY = pageToExtract.PageHeader.TranslateY,
                            LayoutType = (ClientSidePageHeaderLayoutType)Enum.Parse(typeof(ClientSidePageHeaderLayoutType), pageToExtract.PageHeader.LayoutType.ToString()),
                            TextAlignment = (ClientSidePageHeaderTextAlignment)Enum.Parse(typeof(ClientSidePageHeaderTextAlignment), pageToExtract.PageHeader.TextAlignment.ToString()),
                            ShowTopicHeader = pageToExtract.PageHeader.ShowTopicHeader,
                            ShowPublishDate = pageToExtract.PageHeader.ShowPublishDate,
                            TopicHeader = pageToExtract.PageHeader.TopicHeader,
                            AlternativeText = pageToExtract.PageHeader.AlternativeText,
                            Authors = !excludeAuthorInformation ? pageToExtract.PageHeader.Authors : "",
                            AuthorByLine = !excludeAuthorInformation ? pageToExtract.PageHeader.AuthorByLine : "",
                            AuthorByLineId = !excludeAuthorInformation ? pageToExtract.PageHeader.AuthorByLineId : -1
                        };

                        extractedPageInstance.Header = extractedHeader;

                        // Add the page header image to template if that was requested
                        if (creationInfo.PersistBrandingFiles && !string.IsNullOrEmpty(pageToExtract.PageHeader.ImageServerRelativeUrl))
                        {
                            IncludePageHeaderImageInExport(web, pageToExtract.PageHeader.ImageServerRelativeUrl, template, creationInfo, scope);
                        }
                    }

                    // define reusable RegEx pre-compiled objects
                    string guidPattern = "\"{?[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}}?\"";
                    Regex regexGuidPattern = new Regex(guidPattern, RegexOptions.Compiled);

                    string guidPatternEncoded = "=[a-fA-F0-9]{8}(?:%2D|-)([a-fA-F0-9]{4}(?:%2D|-)){3}[a-fA-F0-9]{12}";
                    Regex regexGuidPatternEncoded = new Regex(guidPatternEncoded, RegexOptions.Compiled);

                    string guidPatternNoDashes = "[a-fA-F0-9]{32}";
                    Regex regexGuidPatternNoDashes = new Regex(guidPatternNoDashes, RegexOptions.Compiled);

                    string guidPatternOptionalBrackets = "(?<Bracket>\\{)?[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}(?(Bracket)\\}|)";
                    Regex regexGuidPatternOptionalBrackets = new Regex(guidPatternOptionalBrackets, RegexOptions.Compiled);

                    string siteAssetUrlsPattern = "(?:\")(?<AssetUrl>[\\w|\\.|\\/|:|-]*\\/SiteAssets\\/SitePages\\/[\\w|\\.|\\/|:|-]*)(?:\")";
                    // OLD RegEx with Catastrophic Backtracking: @".*""(.*?/SiteAssets/SitePages/.+?)"".*";
                    Regex regexSiteAssetUrls = new Regex(siteAssetUrlsPattern, RegexOptions.Compiled);

                    if (creationInfo.PersistBrandingFiles && !string.IsNullOrEmpty(extractedPageInstance.ThumbnailUrl))
                    {
                        var thumbnailFileIds = new List<Guid>();
                        CollectImageFilesFromGenericGuids(regexGuidPatternNoDashes, null, regexGuidPatternOptionalBrackets, extractedPageInstance.ThumbnailUrl, thumbnailFileIds);
                        if (thumbnailFileIds.Count == 1)
                        {
                            try{
                                
                                var file = web.GetFileById(thumbnailFileIds[0]);
                                web.Context.Load(file, f => f.Level, f => f.ServerRelativePath, f => f.UniqueId);
                                web.Context.ExecuteQueryRetry();

                                // Item1 = was file added to the template
                                // Item2 = file name (if file found)
                                var imageAddedTuple = LoadAndAddPageImage(web, file, template, creationInfo, scope);
                                if (imageAddedTuple.Item1)
                                {
                                    extractedPageInstance.ThumbnailUrl = Regex.Replace(extractedPageInstance.ThumbnailUrl, file.UniqueId.ToString("N"), $"{{fileuniqueid:{file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray())}}}");
                                }
                                
                            }
                            catch(ServerException ex)
                            {
                                //Catch File Not found exception if Guid does not match with a file
                                //There can be a thumbnail image url containing a Guid without having to be an image in the SharePoint site
                                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                                {
                                    scope.LogDebug($"File with id {thumbnailFileIds[0]} not loaded.");
                                }
                                else
                                {
                                    throw ex;
                                }
                            }

                        }
                    }

                    // Add the sections
                    foreach (var section in pageToExtract.Sections)
                    {
                        // Set order
                        var sectionInstance = new CanvasSection()
                        {
                            Order = section.Order,
                            BackgroundEmphasis = (Emphasis)section.ZoneEmphasis,
                        };
                        if (section.VerticalSectionColumn != null)
                        {
                            sectionInstance.VerticalSectionEmphasis = (Emphasis)section.VerticalSectionColumn.VerticalSectionEmphasis;
                        }
                        // Set section type
                        switch (section.Type)
                        {
                            case PnPCore.CanvasSectionTemplate.OneColumn:
                                sectionInstance.Type = CanvasSectionType.OneColumn;
                                break;
                            case PnPCore.CanvasSectionTemplate.TwoColumn:
                                sectionInstance.Type = CanvasSectionType.TwoColumn;
                                break;
                            case PnPCore.CanvasSectionTemplate.TwoColumnLeft:
                                sectionInstance.Type = CanvasSectionType.TwoColumnLeft;
                                break;
                            case PnPCore.CanvasSectionTemplate.TwoColumnRight:
                                sectionInstance.Type = CanvasSectionType.TwoColumnRight;
                                break;
                            case PnPCore.CanvasSectionTemplate.ThreeColumn:
                                sectionInstance.Type = CanvasSectionType.ThreeColumn;
                                break;
                            case PnPCore.CanvasSectionTemplate.OneColumnFullWidth:
                                sectionInstance.Type = CanvasSectionType.OneColumnFullWidth;
                                break;
                            case PnPCore.CanvasSectionTemplate.OneColumnVerticalSection:
                                sectionInstance.Type = CanvasSectionType.OneColumnVerticalSection;
                                break;
                            case PnPCore.CanvasSectionTemplate.TwoColumnVerticalSection:
                                sectionInstance.Type = CanvasSectionType.TwoColumnVerticalSection;
                                break;
                            case PnPCore.CanvasSectionTemplate.TwoColumnLeftVerticalSection:
                                sectionInstance.Type = CanvasSectionType.TwoColumnLeftVerticalSection;
                                break;
                            case PnPCore.CanvasSectionTemplate.TwoColumnRightVerticalSection:
                                sectionInstance.Type = CanvasSectionType.TwoColumnRightVerticalSection;
                                break;
                            case PnPCore.CanvasSectionTemplate.ThreeColumnVerticalSection:
                                sectionInstance.Type = CanvasSectionType.ThreeColumnVerticalSection;
                                break;
                            default:
                                sectionInstance.Type = CanvasSectionType.OneColumn;
                                break;
                        }

                        // Add controls to section
                        foreach (var column in section.Columns)
                        {
                            foreach (var control in column.Controls)
                            {
                                // Create control
                                CanvasControl controlInstance = new CanvasControl()
                                {
                                    Column = column.IsVerticalSectionColumn ? section.Columns.IndexOf(column) + 1 : column.Order,
                                    ControlId = control.InstanceId,
                                    Order = control.Order,
                                };

                                // Set control type
                                //if (control.Type == typeof(Pages.ClientSideText))
                                if (control is PnPCore.IPageText)
                                {
                                    controlInstance.Type = WebPartType.Text;

                                    // Set text content
                                    controlInstance.ControlProperties = new System.Collections.Generic.Dictionary<string, string>(1)
                                        {
                                            //{ "Text", TokenizeJsonTextData(web, (control as Pages.ClientSideText).Text) }
                                            { "Text", TokenizeJsonTextData(web, (control as PnPCore.IPageText).Text) }
                                        };
                                }
                                else
                                {
                                    // set ControlId to webpart id
                                    //controlInstance.ControlId = Guid.Parse((control as Pages.ClientSideWebPart).WebPartId);
                                    controlInstance.ControlId = Guid.Parse((control as PnPCore.IPageWebPart).WebPartId);
                                    //var webPartType = Pages.ClientSidePage.NameToClientSideWebPartEnum((control as Pages.ClientSideWebPart).WebPartId);
                                    var webPartType = pageToExtract.WebPartIdToDefaultWebPart((control as PnPCore.IPageWebPart).WebPartId);
                                    switch (webPartType)
                                    {
                                        case PnPCore.DefaultWebPart.ContentRollup:
                                            controlInstance.Type = WebPartType.ContentRollup;
                                            break;
                                        case PnPCore.DefaultWebPart.BingMap:
                                            controlInstance.Type = WebPartType.BingMap;
                                            break;
                                        case PnPCore.DefaultWebPart.Button:
                                            controlInstance.Type = WebPartType.Button;
                                            break;
                                        case PnPCore.DefaultWebPart.CallToAction:
                                            controlInstance.Type = WebPartType.CallToAction;
                                            break;
                                        case PnPCore.DefaultWebPart.News:
                                            controlInstance.Type = WebPartType.News;
                                            break;
                                        case PnPCore.DefaultWebPart.PowerBIReportEmbed:
                                            controlInstance.Type = WebPartType.PowerBIReportEmbed;
                                            break;
                                        case PnPCore.DefaultWebPart.Sites:
                                            controlInstance.Type = WebPartType.Sites;
                                            break;
                                        case PnPCore.DefaultWebPart.GroupCalendar:
                                            controlInstance.Type = WebPartType.GroupCalendar;
                                            break;
                                        case PnPCore.DefaultWebPart.MicrosoftForms:
                                            controlInstance.Type = WebPartType.MicrosoftForms;
                                            break;
                                        case PnPCore.DefaultWebPart.ClientWebPart:
                                            controlInstance.Type = WebPartType.ClientWebPart;
                                            break;
                                        case PnPCore.DefaultWebPart.ContentEmbed:
                                            controlInstance.Type = WebPartType.ContentEmbed;
                                            break;
                                        case PnPCore.DefaultWebPart.DocumentEmbed:
                                            controlInstance.Type = WebPartType.DocumentEmbed;
                                            break;
                                        case PnPCore.DefaultWebPart.Image:
                                            controlInstance.Type = WebPartType.Image;
                                            break;
                                        case PnPCore.DefaultWebPart.ImageGallery:
                                            controlInstance.Type = WebPartType.ImageGallery;
                                            break;
                                        case PnPCore.DefaultWebPart.LinkPreview:
                                            controlInstance.Type = WebPartType.LinkPreview;
                                            break;
                                        case PnPCore.DefaultWebPart.NewsFeed:
                                            controlInstance.Type = WebPartType.NewsFeed;
                                            break;
                                        case PnPCore.DefaultWebPart.NewsReel:
                                            controlInstance.Type = WebPartType.NewsReel;
                                            break;
                                        case PnPCore.DefaultWebPart.QuickChart:
                                            controlInstance.Type = WebPartType.QuickChart;
                                            break;
                                        case PnPCore.DefaultWebPart.SiteActivity:
                                            controlInstance.Type = WebPartType.SiteActivity;
                                            break;
                                        case PnPCore.DefaultWebPart.VideoEmbed:
                                            controlInstance.Type = WebPartType.VideoEmbed;
                                            break;
                                        case PnPCore.DefaultWebPart.YammerEmbed:
                                            controlInstance.Type = WebPartType.YammerEmbed;
                                            break;
                                        case PnPCore.DefaultWebPart.Events:
                                            controlInstance.Type = WebPartType.Events;
                                            break;
                                        case PnPCore.DefaultWebPart.Hero:
                                            controlInstance.Type = WebPartType.Hero;
                                            break;
                                        case PnPCore.DefaultWebPart.List:
                                            controlInstance.Type = WebPartType.List;
                                            break;
                                        case PnPCore.DefaultWebPart.PageTitle:
                                            controlInstance.Type = WebPartType.PageTitle;
                                            break;
                                        case PnPCore.DefaultWebPart.People:
                                            controlInstance.Type = WebPartType.People;
                                            break;
                                        case PnPCore.DefaultWebPart.QuickLinks:
                                            controlInstance.Type = WebPartType.QuickLinks;
                                            break;
                                        case PnPCore.DefaultWebPart.CustomMessageRegion:
                                            controlInstance.Type = WebPartType.CustomMessageRegion;
                                            break;
                                        case PnPCore.DefaultWebPart.Divider:
                                            controlInstance.Type = WebPartType.Divider;
                                            break;
                                        case PnPCore.DefaultWebPart.Spacer:
                                            controlInstance.Type = WebPartType.Spacer;
                                            break;
                                        case PnPCore.DefaultWebPart.ThirdParty:
                                            controlInstance.Type = WebPartType.Custom;
                                            break;
                                        default:
                                            controlInstance.Type = WebPartType.Custom;
                                            break;
                                    }
                                    if (excludeAuthorInformation)
                                    {
                                        // CHECK:
                                        //if (webPartType == PnPCore.DefaultWebPart.News)
                                        //{
                                        //    var properties = (control as PnPCore.IPageWebPart).Properties;
                                        //    var authorTokens = properties.SelectTokens("$..author").ToList();
                                        //    foreach (var authorToken in authorTokens)
                                        //    {
                                        //        authorToken.Parent.Remove();
                                        //    }
                                        //    var authorAccountNameTokens = properties.SelectTokens("$..authorAccountName").ToList();
                                        //    foreach (var authorAccountNameToken in authorAccountNameTokens)
                                        //    {
                                        //        authorAccountNameToken.Parent.Remove();
                                        //    }

                                        //    (control as PnPCore.IPageWebPart).PropertiesJson = properties.ToString();
                                        //}
                                    }
                                    string jsonControlData = "\"id\": \"" + (control as PnPCore.IPageWebPart).WebPartId + "\", \"instanceId\": \"" + (control as PnPCore.IPageWebPart).InstanceId + "\", \"title\": " + JsonConvert.ToString((control as PnPCore.IPageWebPart).Title) + ", \"description\": " + JsonConvert.ToString((control as PnPCore.IPageWebPart).Description) + ", \"dataVersion\": \"" + (control as PnPCore.IPageWebPart).DataVersion + "\", \"properties\": " + (control as PnPCore.IPageWebPart).PropertiesJson + "";

                                    // set the control properties
                                    if (!(control as PnPCore.IPageWebPart).ServerProcessedContent.Equals(default))
                                    {
                                        // If we have serverProcessedContent then also export that one, it's important as some controls depend on this information to be present
                                        string serverProcessedContent = (control as PnPCore.IPageWebPart).ServerProcessedContent.ToString();
                                        jsonControlData = jsonControlData + ", \"serverProcessedContent\": " + serverProcessedContent + "";
                                    }

                                    if (!(control as PnPCore.IPageWebPart).DynamicDataPaths.Equals(default))
                                    {
                                        // If we have serverProcessedContent then also export that one, it's important as some controls depend on this information to be present
                                        string dynamicDataPaths = (control as PnPCore.IPageWebPart).DynamicDataPaths.ToString();
                                        jsonControlData = jsonControlData + ", \"dynamicDataPaths\": " + dynamicDataPaths + "";
                                    }

                                    if (!(control as PnPCore.IPageWebPart).DynamicDataValues.Equals(default))
                                    {
                                        // If we have serverProcessedContent then also export that one, it's important as some controls depend on this information to be present
                                        string dynamicDataValues = (control as PnPCore.IPageWebPart).DynamicDataValues.ToString();
                                        jsonControlData = jsonControlData + ", \"dynamicDataValues\": " + dynamicDataValues + "";
                                    }

                                    controlInstance.JsonControlData = "{" + jsonControlData + "}";

                                    var untokenizedJsonControlData = controlInstance.JsonControlData;
                                    // Tokenize the JsonControlData
                                    controlInstance.JsonControlData = TokenizeJsonControlData(web, controlInstance.JsonControlData);
                                    TokenizeBeforeExport(web, template, creationInfo, scope, errorneousOrNonImageFileGuids, regexGuidPattern, regexGuidPatternEncoded, regexGuidPatternOptionalBrackets, regexSiteAssetUrls, controlInstance, untokenizedJsonControlData);
                                }
                                // add control to section
                                sectionInstance.Controls.Add(controlInstance);
                            }
                        }

                        extractedPageInstance.Sections.Add(sectionInstance);
                    }

                    // Renumber the sections...when editing modern homepages you can end up with section with order 0.5 or 0.75...let's ensure we render section as of 1
                    int sectionOrder = 1;
                    foreach (var sectionInstance in extractedPageInstance.Sections)
                    {
                        sectionInstance.Order = sectionOrder;
                        sectionOrder++;
                    }

                    // Spaces support
                    if (pageToExtract.LayoutType == PnPCore.PageLayoutType.Spaces && !string.IsNullOrEmpty(pageToExtract.SpaceContent))
                    {
                        extractedPageInstance.FieldValues.Add(SpaceContentField, pageToExtract.SpaceContent);
                    }


                    if (pageToExtract.LayoutType == PnPCore.PageLayoutType.Topic)
                    {
                        // Extract the topic page header controls (the controls which cannot be moved around on the page). 
                        // These controls will be stored in a one-column section with a order of 999999. 
                        // TODO: this requires a schema change to store these controls in a more elegant manner
                        // Create section

                        var sectionInstance = new CanvasSection()
                        {
                            Order = 999999,
                            Type = CanvasSectionType.OneColumn,
                        };

                        foreach (var headerControl in pageToExtract.HeaderControls)
                        {
                            // Create control
                            CanvasControl controlInstance = new CanvasControl()
                            {
                                Column = 1,
                                ControlId = headerControl.InstanceId,
                                Order = headerControl.Order,
                            };

                            controlInstance.ControlId = Guid.Parse((headerControl as PnPCore.IPageWebPart).WebPartId);
                            controlInstance.Type = WebPartType.Custom;

                            string jsonControlData = "\"id\": \"" + (headerControl as PnPCore.IPageWebPart).WebPartId + "\", \"instanceId\": \"" + (headerControl as PnPCore.IPageWebPart).InstanceId + "\", \"title\": " + JsonConvert.ToString((headerControl as PnPCore.IPageWebPart).Title) + ", \"description\": " + JsonConvert.ToString((headerControl as PnPCore.IPageWebPart).Description) + ", \"dataVersion\": \"" + (headerControl as PnPCore.IPageWebPart).DataVersion + "\", \"properties\": " + (headerControl as PnPCore.IPageWebPart).PropertiesJson + "";

                            // set the control properties
                            if (!(headerControl as PnPCore.IPageWebPart).ServerProcessedContent.Equals(default))
                            {
                                // If we have serverProcessedContent then also export that one, it's important as some controls depend on this information to be present
                                string serverProcessedContent = (headerControl as PnPCore.IPageWebPart).ServerProcessedContent.ToString();
                                jsonControlData = jsonControlData + ", \"serverProcessedContent\": " + serverProcessedContent + "";
                            }

                            controlInstance.JsonControlData = "{" + jsonControlData + "}";

                            var untokenizedJsonControlData = controlInstance.JsonControlData;
                            // Tokenize the JsonControlData
                            controlInstance.JsonControlData = TokenizeJsonControlData(web, controlInstance.JsonControlData);
                            TokenizeBeforeExport(web, template, creationInfo, scope, errorneousOrNonImageFileGuids, regexGuidPattern, regexGuidPatternEncoded, regexGuidPatternOptionalBrackets, regexSiteAssetUrls, controlInstance, untokenizedJsonControlData);
                            // add control to section
                            sectionInstance.Controls.Add(controlInstance);
                        }

                        extractedPageInstance.Sections.Add(sectionInstance);

                        // Extract the topic pages fields                        
                        extractedPageInstance.FieldValues.Add(TopicEntityId, pageToExtract.EntityId == null ? "" : pageToExtract.EntityId);
                        extractedPageInstance.FieldValues.Add(TopicEntityType, pageToExtract.EntityType == null ? "" : pageToExtract.EntityType);
                        extractedPageInstance.FieldValues.Add(TopicEntityRelations, pageToExtract.EntityRelations == null ? "" : pageToExtract.EntityRelations);
                    }

                    // Add the page to the template
                    if (page.IsTranslation)
                    {
                        var parentPage = template.ClientSidePages.Where(p => p.PageName == page.SourcePageName).FirstOrDefault();
                        if (parentPage != null)
                        {
                            var translatedPageInstance = (TranslatedClientSidePage)extractedPageInstance;
                            translatedPageInstance.LCID = new CultureInfo(page.Language).LCID;
                            parentPage.Translations.Add(translatedPageInstance);
                        }
                    }
                    else
                    {
                        var clientSidePageInstance = (ClientSidePage)extractedPageInstance;
                        if (page.TranslatedLanguages != null && page.TranslatedLanguages.Count > 0)
                        {
                            clientSidePageInstance.CreateTranslations = true;
                            clientSidePageInstance.LCID = Convert.ToInt32(web.EnsureProperty(p => p.Language));
                        }
                        template.ClientSidePages.Add(clientSidePageInstance);
                    }

                    // Set the homepage
                    if (page.IsHomePage)
                    {
                        if (template.WebSettings == null)
                        {
                            template.WebSettings = new WebSettings();
                        }

                        if (page.PageUrl.StartsWith(web.ServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase))
                        {
                            template.WebSettings.WelcomePage = page.PageUrl.Replace(web.ServerRelativeUrl + "/", "");
                        }
                        else
                        {
                            template.WebSettings.WelcomePage = page.PageUrl;
                        }
                    }
                }
            }
            catch (ArgumentException ex)
            {
                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_NoValidPage, ex.Message);
            }
        }

        #region Helper methods
        private void TokenizeBeforeExport(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, List<string> errorneousOrNonImageFileGuids, Regex regexGuidPattern, Regex regexGuidPatternEncoded, Regex regexGuidPatternOptionalBrackets, Regex regexSiteAssetUrls, CanvasControl controlInstance, string untokenizedJsonControlData)
        {
            // Export relevant files if this flag is set
            if (creationInfo.PersistBrandingFiles)
            {
                List<Guid> fileGuids = new List<Guid>();
                Dictionary<string, string> exportedFiles = new Dictionary<string, string>();
                Dictionary<string, string> exportedPages = new Dictionary<string, string>();

                CollectSiteAssetImageFiles(regexSiteAssetUrls, web, untokenizedJsonControlData, fileGuids);
                CollectImageFilesFromGenericGuids(regexGuidPattern, regexGuidPatternEncoded, regexGuidPatternOptionalBrackets, untokenizedJsonControlData, fileGuids);

                // Iterate over the found guids to see if they're exportable files
                foreach (var uniqueId in fileGuids)
                {
                    try
                    {
                        if (!exportedFiles.ContainsKey(uniqueId.ToString()) && !errorneousOrNonImageFileGuids.Contains(uniqueId.ToString()))
                        {
                            // Try to see if this is a file
                            var file = web.GetFileById(uniqueId);
                            web.Context.Load(file, f => f.Level, f => f.ServerRelativePath, f => f.ServerRelativeUrl);
                            web.Context.ExecuteQueryRetry();

                            // Item1 = was file added to the template
                            // Item2 = file name (if file found)
                            var imageAddedTuple = LoadAndAddPageImage(web, file, template, creationInfo, scope);

                            if (!string.IsNullOrEmpty(imageAddedTuple.Item2))
                            {
                                if (!imageAddedTuple.Item2.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (imageAddedTuple.Item1)
                                    {
                                        // Keep track of the exported file path and it's UniqueId
                                        exportedFiles.Add(uniqueId.ToString(), file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()));
                                    }
                                }
                                else
                                {
                                    if (!exportedPages.ContainsKey(uniqueId.ToString()))
                                    {
                                        exportedPages.Add(uniqueId.ToString(), file.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()));
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_ErrorDuringFileExport, ex.Message);
                        errorneousOrNonImageFileGuids.Add(uniqueId.ToString());
                    }
                }

                // Tokenize based on the found files, use a different token for encoded guids do we can later on replace by a new encoded guid
                foreach (var exportedFile in exportedFiles)
                {
                    controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedFile.Key.Replace("-", "%2D"), $"{{fileuniqueidencoded:{exportedFile.Value}}}", RegexOptions.IgnoreCase);
                    controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedFile.Key, $"{{fileuniqueid:{exportedFile.Value}}}", RegexOptions.IgnoreCase);
                }
                foreach (var exportedPage in exportedPages)
                {
                    controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedPage.Key.Replace("-", "%2D"), $"{{pageuniqueidencoded:{exportedPage.Value}}}", RegexOptions.IgnoreCase);
                    controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedPage.Key, $"{{pageuniqueid:{exportedPage.Value}}}", RegexOptions.IgnoreCase);
                    controlInstance.JsonControlData = Regex.Replace(controlInstance.JsonControlData, exportedPage.Key.Replace("-", ""), $"{{pageuniqueid:{exportedPage.Value}}}", RegexOptions.IgnoreCase);
                }
            }
        }

        private static void CollectImageFilesFromGenericGuids(Regex regexGuidPattern, Regex regexGuidPatternEncoded, Regex regexGuidPatternOptionalBrackets, string jsonControlData, List<Guid> fileGuids)
        {
            // grab all the guids in the already tokenized json and check try to get them as a file
            if (regexGuidPattern != null)
            {
                if (regexGuidPattern.IsMatch(jsonControlData))
                {
                    foreach (Match guidMatch in regexGuidPattern.Matches(jsonControlData))
                    {
                        Guid uniqueId;
                        if (Guid.TryParse(guidMatch.Value.Trim("\"".ToCharArray()), out uniqueId))
                        {
                            fileGuids.Add(uniqueId);
                        }
                    }
                }
            }
            // grab potentially encoded guids in the already tokenized json and check try to get them as a file
            if (regexGuidPatternEncoded != null)
            {
                if (regexGuidPatternEncoded.IsMatch(jsonControlData))
                {
                    foreach (Match guidMatch in regexGuidPatternEncoded.Matches(jsonControlData))
                    {
                        Guid uniqueId;
                        if (Guid.TryParse(guidMatch.Value.TrimStart("=".ToCharArray()), out uniqueId))
                        {
                            fileGuids.Add(uniqueId);
                        }
                    }
                }
            }
            if (regexGuidPatternOptionalBrackets != null)
            {
                if (regexGuidPatternOptionalBrackets.IsMatch(jsonControlData))
                {
                    foreach (Match guidMatch in regexGuidPatternOptionalBrackets.Matches(jsonControlData))
                    {
                        Guid uniqueId;
                        if (Guid.TryParse(guidMatch.Value, out uniqueId))
                        {
                            if (!fileGuids.Contains(uniqueId))
                            {
                                fileGuids.Add(uniqueId);
                            }
                        }
                    }
                }
            }
        }

        private void IncludePageHeaderImageInExport(Web web, string imageServerRelativeUrl, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope)
        {
            try
            {
                if (!imageServerRelativeUrl.StartsWith("/_LAYOUTS", StringComparison.OrdinalIgnoreCase))
                {
                    var pageHeaderImage = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(imageServerRelativeUrl));
                    web.Context.Load(pageHeaderImage, p => p.Level, p => p.ServerRelativePath);
                    web.Context.ExecuteQueryRetry();

                    LoadAndAddPageImage(web, pageHeaderImage, template, creationInfo, scope);
                }
            }
            catch (Exception)
            {
                // Eat possible exceptions as header images may point to locations outside of the current site (other site collections, _layouts, CDN's, internet)
            }
        }

        private Tuple<bool, string> LoadAndAddPageImage(Web web, Microsoft.SharePoint.Client.File pageImage, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope)
        {
            var baseUri = new Uri(web.Url);
            var fullUri = new Uri(baseUri, pageImage.ServerRelativePath.DecodedUrl);
            var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
            var fileName = Uri.UnescapeDataString(fullUri.Segments[fullUri.Segments.Length - 1]);

            if (!fileName.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase))
            {
                var templateFolderPath = folderPath.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());

                // Avoid duplicate file entries
                var fileAlreadyExported = template.Files.Where(p => p.Folder.Equals(templateFolderPath, StringComparison.CurrentCultureIgnoreCase) &&
                                                                    p.Src.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefault();
                if (fileAlreadyExported == null)
                {
                    // Add a File to the template
                    template.Files.Add(new Model.File()
                    {
                        Folder = templateFolderPath,
                        Src = $"{templateFolderPath}/{fileName}",
                        Overwrite = true,
                        Level = (Model.FileLevel)Enum.Parse(typeof(Model.FileLevel), pageImage.Level.ToString())
                    });

                    // Export the file
                    PersistFile(web, creationInfo, scope, folderPath, fileName);

                    return new Tuple<bool, string>(true, fileName);
                }
            }

            return new Tuple<bool, string>(false, fileName);
        }

        private static void CollectSiteAssetImageFiles(Regex regexSiteAssetUrls, Web web, string untokenizedJsonControlData, List<Guid> fileGuids)
        {
            // match urls to SiteAssets library
            if (regexSiteAssetUrls.IsMatch(untokenizedJsonControlData))
            {
                foreach (Match siteAssetUrlMatch in regexSiteAssetUrls.Matches(untokenizedJsonControlData))
                {
                    var s = siteAssetUrlMatch.Groups[1]?.Value;
                    if (s != null)
                    {
                        // Check if the URL is relative
                        if (s.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase))
                        {
                            // and if not make it relative to the current root site, if it is from the current host
                            var webUrl = web.EnsureProperty(w => w.Url);
                            var slashIndex = webUrl.IndexOf("/", 9);
                            var hostUrl = string.Empty;
                            if (slashIndex == -1)
                            {
                                // Assume we're in a root site
                                hostUrl = webUrl;
                            }
                            else
                            {
                                hostUrl = webUrl.Substring(0, slashIndex);
                            }
                            if (s.StartsWith(hostUrl))
                            {
                                s = s.Substring(hostUrl.Length);
                            }
                        }

                        try
                        {
                            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(s));
                            web.Context.Load(file, f => f.UniqueId);
                            web.Context.ExecuteQueryRetry();
                            fileGuids.Add(file.UniqueId);
                        }
                        catch (Microsoft.SharePoint.Client.ServerException ex)
                        {
                            if (ex.ServerErrorTypeName != "System.IO.FileNotFoundException")
                            {
                                throw ex;
                            }
                        }

                    }
                }
            }
        }

        private string GetParentIdValue(string contentTypeId)
        {
            int length = 0;
            //Exclude the 0x part
            string contentTypeIdValue = contentTypeId.Substring(2);
            for (int i = 0; i < contentTypeIdValue.Length; i += 2)
            {
                length = i;
                if (contentTypeIdValue.Substring(i, 2).Equals("00", StringComparison.OrdinalIgnoreCase))
                {
                    i += 32;
                }
            }
            string parentIdValue = string.Empty;
            if (length > 0)
            {
                parentIdValue = "0x" + contentTypeIdValue.Substring(0, length);
            }
            return parentIdValue;
        }

        private void PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string folderPath, string fileName)
        {
            if (creationInfo.FileConnector != null)
            {
                var fileConnector = creationInfo.FileConnector;
                SharePointConnector connector = new SharePointConnector(web.Context, web.Url, "dummy");
                Uri u = new Uri(web.Url);

                if (u.PathAndQuery != "/")
                {
                    if (folderPath.IndexOf(u.PathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1)
                    {
                        folderPath = folderPath.Replace(u.PathAndQuery, "");
                    }
                }

                folderPath = Uri.UnescapeDataString(folderPath);
                String container = Uri.UnescapeDataString(folderPath).Trim('/').Replace("/", "\\");
                String persistenceFileName = Uri.UnescapeDataString(fileName);

                if (fileConnector.Parameters.ContainsKey(FileConnectorBase.CONTAINER))
                {
                    container = Path.Combine(fileConnector.GetContainer(), container);
                }

                using (Stream s = connector.GetFileStream(persistenceFileName, folderPath))
                {
                    if (s != null)
                    {
                        creationInfo.FileConnector.SaveFileStream(
                            persistenceFileName, container, s);
                    }
                }
            }
            else
            {
                scope.LogError($"No connector present to persist {fileName}.");
            }
        }

        private string TokenizeJsonControlData(Web web, string json)
        {
            if (string.IsNullOrEmpty(json))
            {
                return json;
            }

            var lists = web.Lists;
            var site = (web.Context as ClientContext).Site;
            web.Context.Load(site, s => s.Id, s => s.GroupId);
            web.Context.Load(web, w => w.ServerRelativeUrl, w => w.Id, w => w.Url);
            web.Context.Load(lists, ls => ls.Include(l => l.Id, l => l.Title, l => l.Views.Include(v => v.Id, v => v.Title)));
            web.Context.ExecuteQueryRetry();

            // Tokenize list and list view id's as they can be used by client side web parts (like the list web part)
            foreach (var list in lists)
            {
                json = Regex.Replace(json, list.Id.ToString(), $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}", RegexOptions.IgnoreCase);
                foreach (var view in list.Views)
                {
                    json = Regex.Replace(json, view.Id.ToString(), $"{{viewid:{System.Security.SecurityElement.Escape(list.Title)},{System.Security.SecurityElement.Escape(view.Title)}}}", RegexOptions.IgnoreCase);
                }
            }

            // Some webparts might already contains the site URL using ~sitecollection token (i.e: CQWP) - shouldn't be needed for client side web parts, but just in case
            //json = Regex.Replace(json, "\"~sitecollection/(.)*\"", "\"{site}\"", RegexOptions.IgnoreCase);
            //json = Regex.Replace(json, "'~sitecollection/(.)*'", "'{site}'", RegexOptions.IgnoreCase);
            //json = Regex.Replace(json, ">~sitecollection/(.)*<", ">{site}<", RegexOptions.IgnoreCase);

            // HostUrl token replacement
            var uri = new Uri(web.Url);

            if (web.ServerRelativeUrl != "/")
            {
                json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}:{uri.Port}", $"{uri.Scheme}://{{fqdn}}", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}", $"{uri.Scheme}://{{fqdn}}", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, $"{uri.DnsSafeHost}", "{fqdn}");
            }
            else
            {
                json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}:{uri.Port}", $"{uri.Scheme}://{{fqdn}}{{site}}", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}", $"{uri.Scheme}://{{fqdn}}{{site}}", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, $"{uri.DnsSafeHost}", $"{{fqdn}}", RegexOptions.IgnoreCase);

            }
            // Site token replacement, also replace "encoded" guids
            json = Regex.Replace(json, site.Id.ToString(), "{sitecollectionid}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, site.Id.ToString().Replace("-", "%2D"), "{sitecollectionidencoded}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, site.Id.ToString("N"), "{sitecollectionid}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, web.Id.ToString(), "{siteid}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, web.Id.ToString().Replace("-", "%2D"), "{siteidencoded}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, web.Id.ToString("N"), "{siteid}", RegexOptions.IgnoreCase);
            if (web.ServerRelativeUrl != "/")
            {
                // Normal site collection
                json = Regex.Replace(json, "(\"" + web.ServerRelativeUrl + ")(?!&)", "\"{site}", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, "'" + web.ServerRelativeUrl, "'{site}", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, ">" + web.ServerRelativeUrl, ">{site}", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, web.ServerRelativeUrl, "{site}", RegexOptions.IgnoreCase);
            }
            else
            {
                // Root site collection
                json = Regex.Replace(json, "(\"" + web.ServerRelativeUrl + ")(?!&)", "\"{site}/", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, "'" + web.ServerRelativeUrl, "'{site}/", RegexOptions.IgnoreCase);
                json = Regex.Replace(json, ">" + web.ServerRelativeUrl, ">{site}/", RegexOptions.IgnoreCase);

            }

            // Connected Office 365 group tokenization
            if (!site.GroupId.Equals(Guid.Empty))
            {
                json = Regex.Replace(json, site.GroupId.ToString(), "{sitecollectionconnectedoffice365groupid}", RegexOptions.IgnoreCase);
            }

            return json;
        }
        private string TokenizeJsonTextData(Web web, string json)
        {
            web.Context.Load(web, w => w.ServerRelativeUrl, w => w.Id, w => w.Url);
            web.Context.ExecuteQueryRetry();

            // Only replace links to content, other content is considered to be part of the "Text"
            json = Regex.Replace(json, "href=\"" + web.ServerRelativeUrl, "href=\"{site}", RegexOptions.IgnoreCase);

            return json;
        }
        #endregion
    }
}
