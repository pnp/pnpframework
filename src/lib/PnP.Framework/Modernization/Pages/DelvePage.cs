using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using PnP.Framework.Modernization.Delve;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Extensions;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Transform;
using PnP.Framework.Pages;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Modernization.Pages
{
    public class DelvePage : BasePage
    {
        #region construction
        /// <summary>
        /// Instantiates a web part page object
        /// </summary>
        /// <param name="page">ListItem holding the page to analyze</param>
        /// <param name="pageTransformation">Page transformation information</param>
        public DelvePage(ListItem page, PageTransformation pageTransformation, IList<ILogObserver> logObservers = null) : base(page, null, pageTransformation, logObservers)
        {
        }
        #endregion

        public virtual Tuple<PageLayout, List<WebPartEntity>> AnalyzeAndTransform(DelvePageTransformationInformation pageTransformationInformation, ClientSidePage targetPage)
        {
            List<WebPartEntity> webparts = new List<WebPartEntity>();

            // Load the page
            string delvePageUrl = page[Constants.FileRefField].ToString();
            string fileContents = cc.Web.GetFileByServerRelativeUrlAsString(delvePageUrl);

            if (!string.IsNullOrEmpty(fileContents))
            {
                // Delve page is a json blob
                var pageJson = JToken.Parse(fileContents);

                string pageHeaderImage = null;

                foreach (var controlDataElement in pageJson["ControlData"])
                {
                    //Grab the first
                    var contentElement = controlDataElement.First;

                    string controlType = GetValueAsString(contentElement, "ControlName");

                    if (string.IsNullOrEmpty(controlType))
                    {
                        continue;
                    }

                    if (controlType.Equals("RichTextControl"))
                    {
                       int subType = GetSubType(contentElement["DataContext"]);
                        string textContent = GetValueAsString(contentElement["DataContext"], "Value");

                        // H1 header
                        if (subType == 0)
                        {
                            webparts.Add(CreateWikiTextPart($"<H2>{textContent}</H2>", 1, 1, GetOrder(controlDataElement.First, pageJson["ControlMap"]["Rows"])));
                        }
                        // Text
                        else if (subType == 1)
                        {
                            // Clean html
                            textContent = CleanHtml(textContent);
                            webparts.Add(CreateWikiTextPart(textContent, 1, 1, GetOrder(controlDataElement.First, pageJson["ControlMap"]["Rows"])));
                        }
                        // Pull quote
                        else if (subType == 2)
                        {
                            webparts.Add(CreateWikiTextPart($"<blockquote>{textContent}</blockquote>", 1, 1, GetOrder(controlDataElement.First, pageJson["ControlMap"]["Rows"])));
                        }
                    }
                    else if (controlType.Equals("ImageControl"))
                    {
                        Dictionary<string, string>  props = new Dictionary<string, string>();

                        string imageSource = GetValueAsString(contentElement["DataContext"], "ImageSource");
                        string imageCaption = GetValueAsString(contentElement["DataContext"], "CaptionText");

                        if (!string.IsNullOrEmpty(imageSource))
                        {
                            props.Add("Title", "Image in wiki text");
                            props.Add("Description", "");
                            props.Add("ImageUrl", imageSource.Replace("about://", ""));
                            props.Add("Width", "");
                            props.Add("Height", "");
                            props.Add("Anchor", "");
                            props.Add("Caption", imageCaption ?? "");
                        }

                        webparts.Add(CreateImagePart(props, 1, 1, GetOrder(controlDataElement.First, pageJson["ControlMap"]["Rows"])));
                    }
                    else if (controlType.Equals("VideoControl"))
                    {
                        Dictionary<string, string> props = new Dictionary<string, string>();

                        string videoSource = GetValueAsString(contentElement["DataContext"], "VideoSource");
                        string videoCaption = GetValueAsString(contentElement["DataContext"], "CaptionText");

                        if (!string.IsNullOrEmpty(videoSource))
                        {
                            Uri videoSourceUri = new Uri(videoSource);
                            string queryString = videoSourceUri.Query;
                            var queryDictionary = System.Web.HttpUtility.ParseQueryString(queryString);

                            if (queryDictionary["chid"] != null && queryDictionary["vid"] != null)
                            {
                                string newVideoSource = $"{videoSourceUri.Scheme}://{videoSourceUri.DnsSafeHost}{videoSourceUri.LocalPath}".Replace("pointpublishing.aspx", "VideoEmbedHost.aspx", StringComparison.InvariantCultureIgnoreCase);
                                props.Add("Title", "Video in wiki text");
                                props.Add("Description", "");
                                props.Add("IFrameEmbed", $"<iframe width=640 height=360 src='{newVideoSource}?chId={queryDictionary["chid"]}&vId={queryDictionary["vid"]}&width=640&height=360&autoPlay=false&showInfo=false' allowfullscreen></iframe>");
                                props.Add("Source", videoSource);

                                webparts.Add(CreateVideoPart(props, 1, 1, GetOrder(controlDataElement.First, pageJson["ControlMap"]["Rows"])));
                            }
                        }
                    }
                    else if (controlType.Equals("DocumentHookControl"))
                    {
                        Dictionary<string, string> props = new Dictionary<string, string>();

                        string docSource = GetValueAsString(contentElement["DataContext"]["OfficeDocumentDataContext"], "FullDocumentUrl");

                        if (!string.IsNullOrEmpty(docSource))
                        {
                            props.Add("Title", "Document embed in wiki text");
                            props.Add("Description", "");
                            props.Add("DocumentUrl", docSource);
                            props.Add("Width", "");
                            props.Add("Height", "");
                        }

                        webparts.Add(CreateEmbedPart(props, 1, 1, GetOrder(controlDataElement.First, pageJson["ControlMap"]["Rows"])));
                    }
                    else if (controlType.Equals("ImageHeaderControl"))
                    {
                        pageHeaderImage = GetValueAsString(contentElement["DataContext"], "ImageSource");
                    }
                }

                // Configure page header
                targetPage.SetDefaultPageHeader();

                // Do we have an image page header?
                if (!string.IsNullOrEmpty(pageHeaderImage))
                {
                    string newHeaderImageServerRelativeUrl = "";
                    try
                    {
                        // make server relative image url
                        pageHeaderImage = new Uri(pageHeaderImage).PathAndQuery;
                        
                        // Integrate asset transformator
                        AssetTransfer assetTransfer = new AssetTransfer(page.Context as ClientContext, targetPage.Context, base.RegisteredLogObservers);
                        newHeaderImageServerRelativeUrl = assetTransfer.TransferAsset(pageHeaderImage, System.IO.Path.GetFileNameWithoutExtension(page[Constants.TitleField].ToString()));
                    }
                    catch (Exception ex)
                    {
                        LogError(LogStrings.Error_HeaderImageAssetTransferFailed, LogStrings.Heading_PublishingPageHeader, ex);
                    }

                    if (!string.IsNullOrEmpty(newHeaderImageServerRelativeUrl))
                    {
                        LogInfo(string.Format(LogStrings.SettingHeaderImage, newHeaderImageServerRelativeUrl), LogStrings.Heading_PublishingPageHeader);
                        targetPage.SetCustomPageHeader(newHeaderImageServerRelativeUrl);                       
                    }
                }

                targetPage.PageTitle = pageJson["Title"].Value<string>();

                if (pageTransformationInformation.KeepSubTitle)
                {
                    var topicHeader = GetValueAsString(pageJson, "SubTitle");
                    if (!string.IsNullOrEmpty(topicHeader))
                    {
                        if (topicHeader.Length > 40)
                        {
                            topicHeader = topicHeader.Substring(0, 40);
                        }

                        targetPage.PageHeader.ShowTopicHeader = true;
                        targetPage.PageHeader.TopicHeader = topicHeader;
                    }
                }

                if (pageTransformationInformation.PageTitleOverride != null)
                {
                    var title = pageTransformationInformation.PageTitleOverride(targetPage.PageTitle);
                    targetPage.PageTitle = title;

                    LogInfo($"{LogStrings.TransformPageTitleOverride} - page title: {title}", LogStrings.Heading_ArticlePageHandling);
                }

            }
            else
            {
                // Should never happen
            }

            return new Tuple<PageLayout, List<WebPartEntity>>(PageLayout.Wiki_OneColumn, webparts);
        }

        #region Helper methods
        private int GetOrder(JToken controlId, JToken controlMapRows)
        {
            string controlIdValue = controlId.Path.Replace("ControlData.", "");

            int i = 0;
            foreach(var mapRow in controlMapRows)
            {
                i++;
                var mapRowControlId = GetValueAsString(mapRow["Columns"][0], "ControlId");

                if (mapRowControlId.Equals(controlIdValue, StringComparison.InvariantCultureIgnoreCase))
                {
                    return i;
                }
            }

            return 1;
        }

        private static WebPartEntity CreateImagePart(Dictionary<string, string> properties, int row, int col, int order)
        {
            return new WebPartEntity()
            {
                Title = "WikiImage",
                Type = WebParts.WikiImage,
                Id = Guid.Empty,
                Row = row,
                Column = col,
                Order = order,
                Properties = properties,
            };
        }

        private static WebPartEntity CreateVideoPart(Dictionary<string, string> properties, int row, int col, int order)
        {
            return new WebPartEntity()
            {
                Title = "WikiVideo",
                Type = WebParts.WikiVideo,
                Id = Guid.Empty,
                Row = row,
                Column = col,
                Order = order,
                Properties = properties,
            };
        }

        private static WebPartEntity CreateEmbedPart(Dictionary<string, string> properties, int row, int col, int order)
        {
            return new WebPartEntity()
            {
                Title = "WikiEmbed",
                Type = WebParts.WikiEmbed,
                Id = Guid.Empty,
                Row = row,
                Column = col,
                Order = order,
                Properties = properties,
            };
        }
        private static string CleanHtml(string html)
        {
            if (html.IndexOf("<span style=", StringComparison.InvariantCultureIgnoreCase) >= 0)
            {
                html = html.Replace("<span style=\"\">", "<u>", StringComparison.InvariantCultureIgnoreCase);
                html = html.Replace("<span style=\"text-decoration&#58;underline;\">", "<u>", StringComparison.InvariantCultureIgnoreCase);
                html = html.Replace("</span>", "</u>", StringComparison.InvariantCultureIgnoreCase);
            }

            return html;
        }

        private static string GetValueAsString(JToken token, string name)
        {
            if (token[name] != null)
            {
                return token[name].Value<string>();
            }
            else
            {
                return null;
            }
        }

        private static int GetSubType(JToken token)
        {
            if (token["Subtype"] != null)
            {
                return token["Subtype"].Value<int>();
            }
            else
            {
                return -1;
            }
        }

        #endregion

    }

}
