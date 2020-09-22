using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using PnP.Framework.Pages;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Extensions;
using PnP.Framework.Modernization.Functions;
using PnP.Framework.Modernization.Telemetry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace PnP.Framework.Modernization.Transform
{
    /// <summary>
    /// Transforms content from "classic" page to modern client side page
    /// </summary>
    public class ContentTransformator: BaseTransform, IContentTransformator
    {
        private ClientSidePage page;
        private PageTransformation pageTransformation;
        private FunctionProcessor functionProcessor;
        private List<CombinedMapping> combinedMappinglist;
        private ClientContext sourceClientContext;
        private Dictionary<string, string> globalTokens;
        private bool isCrossSiteTransfer;
        private BaseTransformationInformation transformationInformation;

        class CombinedMapping
        {
            public int Order { get; set; }
            public ClientSideText ClientSideText { get; set; }
            public ClientSideWebPart ClientSideWebPart { get; set; }
        }

        #region Construction
        /// <summary>
        /// Instantiates the content transformator
        /// </summary>
        /// <param name="page">Client side page that will be updates</param>
        /// <param name="pageTransformation">Transformation information</param>
        public ContentTransformator(ClientContext sourceClientContext, ClientSidePage page, PageTransformation pageTransformation, BaseTransformationInformation transformationInformation, IList<ILogObserver> logObservers = null) : base()
        {
            
            //Register any existing observers
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            this.page = page ?? throw new ArgumentException("Page cannot be null");
            this.pageTransformation = pageTransformation ?? throw new ArgumentException("pageTransformation cannot be null");
            this.globalTokens = CreateGlobalTokenList(page.Context, transformationInformation.MappingProperties);
            this.functionProcessor = new FunctionProcessor(sourceClientContext, this.page, this.pageTransformation, transformationInformation, base.RegisteredLogObservers);
            this.transformationInformation = transformationInformation;

            this.sourceClientContext = sourceClientContext;
            this.isCrossSiteTransfer = IsCrossSiteTransfer();

            
        }
        #endregion

        /// <summary>
        /// Transforms the passed web parts into the loaded client side page
        /// </summary>
        /// <param name="webParts">List of web parts that need to be transformed</param>
        public void Transform(List<WebPartEntity> webParts)
        {
            LogInfo(LogStrings.ContentTransformingWebParts, LogStrings.Heading_ContentTransform);

            if (webParts == null || webParts.Count == 0)
            {
                // nothing to transform
                LogWarning(LogStrings.NothingToTransform, LogStrings.Heading_ContentTransform);
                return;
            }

            // find the default mapping, will be used for webparts for which the model does not contain a mapping
            var defaultMapping = pageTransformation.BaseWebPart.Mappings.Mapping.Where(p => p.Default == true).FirstOrDefault();
            if (defaultMapping == null)
            {
                LogError(LogStrings.Error_NoDefaultMappingFound, LogStrings.Heading_ContentTransform);
                throw new Exception(LogStrings.Error_NoDefaultMappingFound);
            }

            // Load existing available controls
            var componentsToAdd = CacheManager.Instance.GetClientSideComponents(page);

            if (this.transformationInformation.SourcePage != null && !this.transformationInformation.SourcePage.PageType().Equals("WebPartPage", StringComparison.InvariantCultureIgnoreCase))
            {
                // Normalize row numbers as there can be gaps if the analyzed page contained wiki tables
                int newRowOrder = 0;
                int lastExistingRowOrder = -1;
                foreach (var webPart in webParts.OrderBy(p => p.Row))
                {
                    if (lastExistingRowOrder < webPart.Row)
                    {
                        newRowOrder++;
                        lastExistingRowOrder = webPart.Row;
                    }

                    webPart.Row = newRowOrder;
                }
            }

            // Iterate over the web parts, important to order them by row, column and zoneindex
            foreach (var webPart in webParts.OrderBy(p => p.Row).ThenBy(p => p.Column).ThenBy(p => p.Order))
            {
                LogInfo(string.Format(LogStrings.ContentWebPartBeingTransformed, webPart.Title, webPart.TypeShort()), LogStrings.Heading_MappingWebParts);

                // Title bar will never be migrated
                if (webPart.Type == WebParts.TitleBar)
                {
                    LogInfo(LogStrings.NotTransformingTitleBar, LogStrings.Heading_MappingWebParts);
                    continue;
                }
                
                // Assign the default mapping, if we're a more specific mapping than that will overwrite this mapping
                Mapping mapping = defaultMapping;

                // Does the web part have a mapping defined?
                // Older version of SharePoint 
                var webPartData = pageTransformation.WebParts.Where(p => p.Type.GetTypeShort() == webPart.Type.GetTypeShort()).FirstOrDefault();

                // Check for cross site transfer support
                if (webPartData != null && this.isCrossSiteTransfer)
                {
                    if (!webPartData.CrossSiteTransformationSupported)
                    {
                        LogWarning(LogStrings.CrossSiteNotSupported, LogStrings.Heading_MappingWebParts);
                        continue;
                    }
                }

                if (webPartData != null && webPartData.Mappings != null)
                {
                    // Add site level (e.g. site) tokens to the web part properties and model so they can be used in the same manner as a web part property
                    UpdateWebPartDataProperties(webPart, webPartData, this.globalTokens);

                    string selectorResult = null;
                    try
                    {
                        // The mapping can have a selector function defined, is so it will be executed. If a selector was executed the selectorResult will contain the name of the mapping to use
                        LogDebug(LogStrings.ProcessingSelectorFunctions, LogStrings.Heading_MappingWebParts);
                        selectorResult = functionProcessor.Process(ref webPartData, webPart);
                    }
                    catch(Exception ex)
                    {
                        // NotAvailableAtTargetException is used to "skip" a web part since it's not valid for the target site collection (only applies to cross site collection transfers)
                        if (ex.InnerException is NotAvailableAtTargetException)
                        {
                            LogError(LogStrings.Error_NotValidForTargetSiteCollection, LogStrings.Heading_MappingWebParts, ex, true);
                            continue;
                        }

                        if (ex.InnerException is MediaWebpartConfigurationException)
                        {
                            LogError(LogStrings.Error_MediaWebpartConfiguration, LogStrings.Heading_MappingWebParts, ex, true);
                            continue;
                        }

                        LogError($"{LogStrings.Error_AnErrorOccurredFunctions} - {ex.Message}", LogStrings.Heading_MappingWebParts, ex);
                        throw;                          
                    }

                    Mapping webPartMapping = null;
                    // Get the needed mapping:
                    // - use the mapping returned by the selector
                    // - if no selector then take the default mapping
                    // - if no mapping found we'll fall back to the default web part mapping
                    if (!string.IsNullOrEmpty(selectorResult))
                    {
                        webPartMapping = webPartData.Mappings.Mapping.Where(p => p.Name.Equals(selectorResult, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    }
                    else
                    {
                        // If there's only one mapping let's take that one, even if not specified as default
                        if (webPartData.Mappings.Mapping.Length == 1)
                        {
                            webPartMapping = webPartData.Mappings.Mapping[0];
                        }
                        else
                        {
                            webPartMapping = webPartData.Mappings.Mapping.Where(p => p.Default == true).FirstOrDefault();
                        }
                    }

                    if (webPartMapping != null)
                    {
                        mapping = webPartMapping;
                    }
                    else
                    {
                        LogWarning(LogStrings.ContentWebPartMappingNotFound, LogStrings.Heading_MappingWebParts);
                    }

                    // Process mapping specific functions (if any)
                    if (!String.IsNullOrEmpty(mapping.Functions))
                    {
                        try
                        {
                            LogInfo(LogStrings.ProcessingMappingFunctions, LogStrings.Heading_MappingWebParts);
                            functionProcessor.ProcessMappingFunctions(ref webPartData, webPart, mapping);
                        }
                        catch (Exception ex)
                        {
                            // NotAvailableAtTargetException is used to "skip" a web part since it's not valid for the target site collection (only applies to cross site collection transfers)
                            if (ex.InnerException is NotAvailableAtTargetException)
                            {
                                LogError(LogStrings.Error_NotValidForTargetSiteCollection, LogStrings.Heading_MappingWebParts, ex, true);
                                continue;
                            }

                            LogError($"{LogStrings.Error_AnErrorOccurredFunctions} - {ex.Message}", LogStrings.Heading_MappingWebParts, ex);
                            throw;
                        }
                    }
                }

                // Use the mapping data => make one list of Text and WebParts to allow for correct ordering
                LogDebug("Combining mapping data", LogStrings.Heading_MappingWebParts);
                combinedMappinglist = new List<CombinedMapping>();
                if (mapping.ClientSideText != null)
                {
                    foreach (var map in mapping.ClientSideText.OrderBy(p => p.Order))
                    {
                        if (!Int32.TryParse(map.Order, out Int32 mapOrder))
                        {
                            mapOrder = 0;
                        }

                        combinedMappinglist.Add(new CombinedMapping { ClientSideText = map, ClientSideWebPart = null, Order = mapOrder });
                    }
                }
                if (mapping.ClientSideWebPart != null)
                {
                    foreach (var map in mapping.ClientSideWebPart.OrderBy(p => p.Order))
                    {
                        if (!Int32.TryParse(map.Order, out Int32 mapOrder))
                        {
                            mapOrder = 0;
                        }

                        combinedMappinglist.Add(new CombinedMapping { ClientSideText = null, ClientSideWebPart = map, Order = mapOrder });
                    }
                }

                // Get the order of the last inserted control in this column
                int order = LastColumnOrder(webPart.Row - 1, webPart.Column - 1);
                // Interate the controls for this mapping using their order
                foreach (var map in combinedMappinglist.OrderBy(p => p.Order))
                {
                    order++;

                    if (map.ClientSideText != null)
                    {
                        // Insert a Text control
                        PnP.Framework.Pages.ClientSideText text = new PnP.Framework.Pages.ClientSideText()
                        {
                            Text = TokenParser.ReplaceTokens(map.ClientSideText.Text, webPart)
                        };

                        page.AddControl(text, page.Sections[webPart.Row - 1].Columns[webPart.Column - 1], order);
                        LogInfo(LogStrings.AddedClientSideTextWebPart, LogStrings.Heading_AddingWebPartsToPage);
                        
                    }
                    else if (map.ClientSideWebPart != null)
                    {
                        // Insert a web part
                        ClientSideComponent baseControl = null;

                        if (map.ClientSideWebPart.Type == ClientSideWebPartType.Custom)
                        {
                            // Parse the control ID to support generic web part placement scenarios
                            map.ClientSideWebPart.ControlId = TokenParser.ReplaceTokens(map.ClientSideWebPart.ControlId, webPart);
                            // Check if this web part belongs to the list of "usable" web parts for this site
                            baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals($"{{{map.ClientSideWebPart.ControlId}}}", StringComparison.InvariantCultureIgnoreCase));
                            LogInfo(LogStrings.UsingCustomModernWebPart, LogStrings.Heading_AddingWebPartsToPage);
                        }
                        else
                        {
                            string webPartName = "";
                            switch (map.ClientSideWebPart.Type)
                            {
                                case ClientSideWebPartType.List:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.List);
                                        break;
                                    }
                                case ClientSideWebPartType.Image:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Image);
                                        break;
                                    }
                                case ClientSideWebPartType.ContentRollup:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.ContentRollup);
                                        break;
                                    }
                                case ClientSideWebPartType.BingMap:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.BingMap);
                                        break;
                                    }
                                case ClientSideWebPartType.ContentEmbed:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.ContentEmbed);
                                        break;
                                    }
                                case ClientSideWebPartType.DocumentEmbed:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.DocumentEmbed);
                                        break;
                                    }
                                case ClientSideWebPartType.ImageGallery:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.ImageGallery);
                                        break;
                                    }
                                case ClientSideWebPartType.LinkPreview:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.LinkPreview);
                                        break;
                                    }
                                case ClientSideWebPartType.NewsFeed:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.NewsFeed);
                                        break;
                                    }
                                case ClientSideWebPartType.NewsReel:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.NewsReel);
                                        break;
                                    }
                                case ClientSideWebPartType.News:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.News);
                                        break;
                                    }
                                case ClientSideWebPartType.PowerBIReportEmbed:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.PowerBIReportEmbed);
                                        break;
                                    }
                                case ClientSideWebPartType.QuickChart:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.QuickChart);
                                        break;
                                    }
                                case ClientSideWebPartType.SiteActivity:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.SiteActivity);
                                        break;
                                    }
                                case ClientSideWebPartType.VideoEmbed:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.VideoEmbed);
                                        break;
                                    }
                                case ClientSideWebPartType.YammerEmbed:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.YammerEmbed);
                                        break;
                                    }
                                case ClientSideWebPartType.Events:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Events);
                                        break;
                                    }
                                case ClientSideWebPartType.GroupCalendar:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.GroupCalendar);
                                        break;
                                    }
                                case ClientSideWebPartType.Hero:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Hero);
                                        break;
                                    }
                                case ClientSideWebPartType.PageTitle:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.PageTitle);
                                        break;
                                    }
                                case ClientSideWebPartType.People:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.People);
                                        break;
                                    }
                                case ClientSideWebPartType.QuickLinks:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.QuickLinks);
                                        break;
                                    }
                                case ClientSideWebPartType.Divider:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Divider);
                                        break;
                                    }
                                case ClientSideWebPartType.MicrosoftForms:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.MicrosoftForms);
                                        break;
                                    }
                                case ClientSideWebPartType.Spacer:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Spacer);
                                        break;
                                    }
                                case ClientSideWebPartType.ClientWebPart:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.ClientWebPart);
                                        break;
                                    }
                                case ClientSideWebPartType.PowerApps:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.PowerApps);
                                        break;
                                    }
                                case ClientSideWebPartType.CodeSnippet:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.CodeSnippet);
                                        break;
                                    }
                                case ClientSideWebPartType.PageFields:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.PageFields);
                                        break;
                                    }
                                case ClientSideWebPartType.Weather:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Weather);
                                        break;
                                    }
                                case ClientSideWebPartType.YouTube:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.YouTube);
                                        break;
                                    }
                                case ClientSideWebPartType.MyDocuments:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.MyDocuments);
                                        break;
                                    }
                                case ClientSideWebPartType.YammerFullFeed:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.YammerFullFeed);
                                        break;
                                    }
                                case ClientSideWebPartType.CountDown:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.CountDown);
                                        break;
                                    }
                                case ClientSideWebPartType.ListProperties:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.ListProperties);
                                        break;
                                    }
                                case ClientSideWebPartType.MarkDown:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.MarkDown);
                                        break;
                                    }
                                case ClientSideWebPartType.Planner:
                                    {
                                        webPartName = ClientSidePage.ClientSideWebPartEnumToName(DefaultClientSideWebParts.Planner);
                                        break;
                                    }
                                default:
                                    {
                                        break;
                                    }
                            }

                            // SharePoint add-ins can be added on client side pages...all add-ins are added via the client web part, so we need additional logic to find the one we need
                            if (map.ClientSideWebPart.Type == ClientSideWebPartType.ClientWebPart)
                            {
                                var addinComponents = componentsToAdd.Where(p => p.Name.Equals(webPartName, StringComparison.InvariantCultureIgnoreCase));
                                foreach(var addin in addinComponents)
                                {
                                    // Find the right add-in web part via title matching...maybe not bullet proof but did find anything better for now
                                    JObject wpJObject = JObject.Parse(addin.Manifest);

                                    // As there can be multiple classic web parts (via provider hosted add ins or SharePoint hosted add ins) we're looping to find the first one with a matching title
                                    foreach(var addinEntry in wpJObject["preconfiguredEntries"])
                                    {
                                        if (addinEntry["title"]["default"].Value<string>() == webPart.Title)
                                        {
                                            baseControl = addin;

                                            var jsonProperties = addinEntry;

                                            // Fill custom web part properties in this json. Custom properties are listed as child elements under clientWebPartProperties, 
                                            // replace their "default" value with the value we got from the web part's properties
                                            jsonProperties = PopulateAddInProperties(jsonProperties, webPart);

                                            // Override the JSON data we read from the model as this is fully dynamic due to the nature of the add-in client part
                                            map.ClientSideWebPart.JsonControlData = jsonProperties.ToString(Newtonsoft.Json.Formatting.None);

                                            LogInfo($"{LogStrings.ContentUsingAddinWebPart} '{baseControl.Name}' ", LogStrings.Heading_AddingWebPartsToPage);
                                            break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                baseControl = componentsToAdd.FirstOrDefault(p => p.Name.Equals(webPartName, StringComparison.InvariantCultureIgnoreCase));
                                LogInfo($"{LogStrings.ContentUsing} '{ map.ClientSideWebPart.Type.ToString() }' {LogStrings.ContentModernWebPart}", LogStrings.Heading_AddingWebPartsToPage);
                            }
                        }

                        // If we found the web part as a possible candidate to use then add it
                        if (baseControl != null)
                        {
                            var jsonDecoded = WebUtility.HtmlDecode(TokenParser.ReplaceTokens(map.ClientSideWebPart.JsonControlData, webPart));
                            PnP.Framework.Pages.ClientSideWebPart myWebPart = new PnP.Framework.Pages.ClientSideWebPart(baseControl)
                            {
                                Order = map.Order,
                                PropertiesJson = jsonDecoded
                            };

                            page.AddControl(myWebPart, page.Sections[webPart.Row - 1].Columns[webPart.Column - 1], order);
                            LogInfo($"{LogStrings.ContentAdded} '{ myWebPart.Title }' {LogStrings.ContentClientToTargetPage}", LogStrings.Heading_AddingWebPartsToPage);
                        }
                        else
                        {
                            LogWarning(LogStrings.ContentWarnModernNotFound, LogStrings.Heading_AddingWebPartsToPage);
                        }

                    }
                }
            }

            LogInfo(LogStrings.ContentTransformationComplete, LogStrings.Heading_ContentTransform);
        }

        #region Helper methods
        private bool IsCrossSiteTransfer()
        {

            if (this.sourceClientContext == null)
            {
                return false;
            }
            
            if (this.sourceClientContext.Web.GetUrl().Equals(this.page.Context.Web.GetUrl(), StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }

            return true;
        }

        private void UpdateWebPartDataProperties(WebPartEntity webPart, WebPart webPartData, Dictionary<string,string> globalProperties)
        {
            List<Property> tempList = new List<Property>();
            if (webPartData.Properties != null)
            {
                tempList.AddRange(webPartData.Properties);
            }

            // Add properties listed on the Base web part
            var baseProperties = this.pageTransformation.BaseWebPart.Properties;
            foreach (var baseProperty in baseProperties)
            {
                // Only add the global property once as the webPartData.Properties collection is reused across web parts and pages
                var propAlreadyAdded = tempList.Where(p => p.Name.Equals(baseProperty.Name, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (propAlreadyAdded == null)
                {
                    // Add parameter to model
                    tempList.Add(new Property()
                    {
                        Functions = baseProperty.Functions,
                        Name = baseProperty.Name,
                        Type = PropertyType.@string
                    });
                }
            }

            // Add global properties
            foreach (var token in globalProperties)
            {
                // Add property to web part
                if (!webPart.Properties.ContainsKey(token.Key))
                {
                    webPart.Properties.Add(token.Key, token.Value);
                }

                // Only add the global property once as the webPartData.Properties collection is reused across web parts and pages
                var propAlreadyAdded = tempList.Where(p => p.Name.Equals(token.Key, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (propAlreadyAdded == null)
                {
                    // Add parameter to model
                    tempList.Add(new Property()
                    {
                        Functions = "",
                        Name = token.Key,
                        Type = PropertyType.@string
                    });
                }
            }
            webPartData.Properties = tempList.ToArray();
        }

        private JToken PopulateAddInProperties(JToken jsonProperties, WebPartEntity webpart)
        {
            foreach(JToken property in jsonProperties["properties"]["clientWebPartProperties"])
            {
                var wpProp = property["name"].Value<string>();
                if (!string.IsNullOrEmpty(wpProp))
                {
                    if (webpart.Properties.ContainsKey(wpProp))
                    {
                        if (jsonProperties["properties"]["userDefinedProperties"][wpProp] != null)
                        {
                            jsonProperties["properties"]["userDefinedProperties"][wpProp] = webpart.Properties[wpProp].ToString();
                        }
                        else
                        {
                            JToken newProp = JObject.Parse($"{{\"{wpProp}\": \"{webpart.Properties[wpProp].ToString()}\"}}");
                            (jsonProperties["properties"]["userDefinedProperties"] as JObject).Merge(newProp);
                        }
                    }
                }
            }

            return jsonProperties;
        }

        private Dictionary<string, string> CreateGlobalTokenList(ClientContext cc, Dictionary<string, string> mappingProperties)
        {
            Dictionary<string, string> globalTokens = new Dictionary<string, string>(5);

            var url = cc.Web.GetUrl();
            Uri hostUri = new Uri(url);
            
            // Add the fixed properties
            globalTokens.Add("Host", $"{hostUri.Scheme}://{hostUri.DnsSafeHost}");
            globalTokens.Add("Web", cc.Web.ServerRelativeUrl.TrimEnd('/'));
            globalTokens.Add("SiteCollection", cc.Site.RootWeb.ServerRelativeUrl.TrimEnd('/'));
            globalTokens.Add("WebId", cc.Web.Id.ToString());
            globalTokens.Add("SiteId", cc.Site.Id.ToString());

            // Add the properties provided via configuration
            foreach(var property in mappingProperties)
            {
                globalTokens.Add(property.Key, property.Value);
            }

            return globalTokens;
        }

        private Int32 LastColumnOrder(int row, int col)
        {
            var lastControl = page.Sections[row].Columns[col].Controls.OrderBy(p => p.Order).LastOrDefault();
            if (lastControl != null)
            {
                return lastControl.Order;
            }
            else
            {
                return -1;
            }
        }
        #endregion

    }
}
