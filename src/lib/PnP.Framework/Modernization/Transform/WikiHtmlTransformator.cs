using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Css.Dom;
using AngleSharp.Html.Parser;
using Microsoft.SharePoint.Client;
using PnP.Framework.Modernization.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using PnP.Framework.Pages;
using PnP.Framework.Modernization.Functions;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Extensions;

namespace PnP.Framework.Modernization.Transform
{
    /// <summary>
    /// Translates wiki text parts in to a list of real and fake web parts
    /// </summary>
    public class WikiHtmlTransformator: BaseTransform
    {
        private HtmlParser parser;
        private ClientContext sourceContext;
        private ClientSidePage page;
        private Dictionary<string, string> mappingProperties;
        private BuiltIn builtInFunctions;

        #region Construction
        /// <summary>
        /// Default constructor
        /// </summary>
        public WikiHtmlTransformator(ClientContext sourceContext, ClientSidePage page, BaseTransformationInformation basePageTransformationInformation, IList<ILogObserver> logObservers = null)
        {
            //Register any existing observers
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            this.sourceContext = sourceContext;
            this.page = page;
            this.mappingProperties = basePageTransformationInformation.MappingProperties;

            // Instantiate BuiltIn functions class
            this.builtInFunctions = new BuiltIn(basePageTransformationInformation, this.page.Context, this.sourceContext, this.page, base.RegisteredLogObservers);

            // Instantiate the AngleSharp Html parser
            parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true });
        }
        #endregion

        /// <summary>
        /// Replaces embedded images and iframes with respective "fake" image and video web parts. Depending on the 
        /// image/iframe position in the html the wiki text is broken up in multiple wiki text parts intermixed 
        /// with image and/or video parts. Later on these web parts will be transformed to client side web parts
        /// </summary>
        /// <param name="wikiPageWebParts">List of web parts on the page</param>
        /// <param name="handleWikiImagesAndVideos">If true images and videos embedded in wiki text will be transformed to actual image/video web parts,
        /// else they'll get a placeholder and will be added as separate web parts at the end of the page</param>
        /// <param name="addTableListImageAsImageWebPart">When an image lives inside a table (or list) then also add it as a separate image web part</param>
        /// <returns>Updated list of web parts</returns>
        public List<WebPartEntity> TransformPlusSplit(List<WebPartEntity> wikiPageWebParts, bool handleWikiImagesAndVideos, bool addTableListImageAsImageWebPart)
        {
            List<WebPartEntity> updatedWebParts = new List<WebPartEntity>(wikiPageWebParts.Count + 10);
            List<WebPartEntity> replacedWebParts = new List<WebPartEntity>(10);

            // Counters used for breaking up wiki text and placing images in between the broken wiki text fragments
            int lastRow = 1;
            int lastColum = 1;
            int extraWebPartCounter = 1;

            // first ensure there's a big gap in the ordering to allow insertion
            foreach (var wp in wikiPageWebParts)
            {
                wp.Order = wp.Order * 1000;
            }

            // Counters used for placing web parts at the end of the page (e.g. image was defined inside a table)
            int lastRow2 = 1;
            int lastColum2 = 1;
            int extraWebPartCounter2 = 1;
            int splitCounter = 1;

            var lastWebPart2 = wikiPageWebParts.OrderByDescending(p => p.Row).ThenByDescending(p => p.Column).ThenByDescending(p => p.Order).FirstOrDefault();
            if (lastWebPart2 != null)
            {
                lastRow2 = lastWebPart2.Row;
                lastColum2 = lastWebPart2.Column;
                extraWebPartCounter2 = lastWebPart2.Order + 100;
            }

            // iterate over all parts found on the wiki page
            foreach (var wp in wikiPageWebParts)
            {
                string htmlToParse = "";
                if (wp.Type == WebParts.WikiText)
                {
                    htmlToParse = wp.Properties["Text"];
                }
                else if (wp.Type.GetTypeShort() == WebParts.ContentEditor.GetTypeShort())
                {
                    string fileContents = "";
                    if (wp.Properties.ContainsKey("ContentLink") && !string.IsNullOrEmpty(wp.Properties["ContentLink"]) && !wp.Properties["ContentLink"].ToLower().EndsWith(".aspx"))
                    {
                        // Load file contents
                        fileContents = this.builtInFunctions.LoadContentFromFile(wp.Properties["ContentLink"]);
                    }

                    // Run the same selector as we're running from the default mapping file
                    var selectorResult = this.builtInFunctions.ContentEmbedSelectorContentLink((wp.Properties.ContainsKey("ContentLink") ? wp.Properties["ContentLink"] : null), (wp.Properties.ContainsKey("Content") ? wp.Properties["Content"] : null), fileContents, this.mappingProperties[Constants.UseCommunityScriptEditorMappingProperty]);

                    if (selectorResult.Equals("NonASPXLinkNoScript", StringComparison.InvariantCultureIgnoreCase) ||
                        selectorResult.Equals("ContentNoScript", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!string.IsNullOrEmpty(fileContents))
                        {
                            htmlToParse = fileContents;
                        }
                        else
                        {
                            htmlToParse = wp.Properties["Content"];
                        }

                        // Is there a need to add the web part title as content?
                        htmlToParse = IncludeWebPartTitle(htmlToParse, wp.Properties);
                    }
                }

                if (!string.IsNullOrEmpty(htmlToParse))
                {
                    // Reset the replaced web parts list
                    replacedWebParts = new List<WebPartEntity>(10);

                    // Parse the html
                    using (var document = this.parser.ParseDocument(htmlToParse))
                    {
                        // Check if this text requires special handling due to embedded images or iframes...
                        var images = document.QuerySelectorAll("img");
                        var iframes = document.QuerySelectorAll("iframe");
                        var elementsToHandle = images.Union(iframes);

                        // No special handling needed, so skip this wiki text part
                        if (!elementsToHandle.Any())
                        {
                            if (wp.Type == WebParts.ContentEditor)
                            {
                                // Since we've already read (and possible extended) the content let's 'rewrite' the web part properties
                                if (wp.Properties.ContainsKey("ContentLink"))
                                {
                                    wp.Properties["ContentLink"] = "";                                    
                                }

                                if (!wp.Properties.ContainsKey("Content"))
                                {
                                    wp.Properties.Add("Content", htmlToParse);
                                }
                                else
                                {
                                    wp.Properties["Content"] = htmlToParse;
                                }
                            }

                            updatedWebParts.Add(wp);
                            continue;
                        }

                        // Right, we've found a wiki text part with images or iframes...
                        lastRow = wp.Row;
                        lastColum = wp.Column;
                        extraWebPartCounter = wp.Order;

                        // Iterate over each each element, need to check each element to ensure we create the 
                        // replacement web parts in the right order
                        foreach (var element in document.All)
                        {
                            Dictionary<string, string> props = new Dictionary<string, string>();

                            // Img elements might require splitting of wiki text
                            if (element is IHtmlImageElement)
                            {
                                bool split = true;
                                bool inUnSplitableElement = InUnSplitableElement(element);

                                // Only split if the image was not living in a table or list
                                if (handleWikiImagesAndVideos && !inUnSplitableElement)
                                {
                                    // Get the current html tree from this element up and add as text part
                                    props.Add("Title", "Wiki text");
                                    props.Add("Text", $"SplitPart{splitCounter}");
                                    splitCounter++;

                                    replacedWebParts.Add(new WebPartEntity()
                                    {
                                        Type = WebParts.WikiText,
                                        Title = "Wiki text",
                                        Row = lastRow,
                                        Column = lastColum,
                                        Order = extraWebPartCounter,
                                        Properties = props
                                    });
                                    extraWebPartCounter++;
                                }
                                else
                                {
                                    split = false;
                                }

                                // Check if this image tag is wrapped inside an Anchor
                                string anchorTag = null;
                                string captionHtml = "";
                                if (element.ParentElement != null && element.ParentElement.TagName.Equals("A", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (element.ParentElement.HasAttribute("href"))
                                    {
                                        anchorTag = element.ParentElement.GetAttribute("href");
                                    }

                                    if (!string.IsNullOrEmpty(element.ParentElement.TextContent))
                                    {
                                        captionHtml = element.ParentElement.TextContent;
                                    }
                                }

                                // Fill properties of the image web part
                                props = new Dictionary<string, string>();
                                if ((element as IHtmlImageElement).Source != null)
                                {
                                    props.Add("Title", "Image in wiki text");
                                    props.Add("Description", "");
                                    props.Add("ImageUrl", (element as IHtmlImageElement).Source.Replace("about://", ""));
                                    props.Add("Width", (element as IHtmlImageElement).DisplayWidth.ToString());
                                    props.Add("Height", (element as IHtmlImageElement).DisplayHeight.ToString());
                                    props.Add("Anchor", anchorTag ?? "");
                                    props.Add("Caption", captionHtml ?? "");
                                }

                                var alt = (element as IElement).Attributes.Where(p => p.Name.Equals("alt", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                                if (alt != null)
                                {
                                    props.Add("AlternativeText", alt.Value.CleanForJSON());
                                }

                                if (!inUnSplitableElement || addTableListImageAsImageWebPart)
                                {
                                    // Add image part
                                    replacedWebParts.Add(new WebPartEntity()
                                    {
                                        Type = WebParts.WikiImage,
                                        Title = "Image in wiki text",
                                        Row = split ? lastRow : lastRow2,
                                        Column = split ? lastColum : lastColum2,
                                        Order = split ? extraWebPartCounter : extraWebPartCounter2,
                                        Properties = props
                                    });
                                }

                                if (split)
                                {
                                    // replace img or img nested in A with "splitter"
                                    var splitter = document.CreateElement("span");
                                    splitter.ClassName = "split";

                                    if (element.ParentElement != null)
                                    {
                                        if (element.ParentElement.TagName.Equals("A", StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            element.ParentElement.ParentElement.ReplaceChild(splitter, element.ParentElement);
                                        }
                                        else
                                        {
                                            element.ParentElement.ReplaceChild(splitter, element);
                                        }
                                    }
                                    extraWebPartCounter++;
                                }
                                else
                                {
                                    extraWebPartCounter2++;
                                }
                            }
                            // IFrame elements might require splitting of wiki text
                            else if (element is IHtmlInlineFrameElement)
                            {
                                bool split = true;
                                bool inUnSplitableElement = InUnSplitableElement(element);

                                // Only split if the iframe was not living in a table or list
                                if (handleWikiImagesAndVideos && !inUnSplitableElement)
                                {
                                    // Get the current html tree from this element up and add as text part
                                    props.Add("Title", "Wiki text");
                                    props.Add("Text", $"SplitPart{splitCounter}");
                                    splitCounter++;

                                    replacedWebParts.Add(new WebPartEntity()
                                    {
                                        Type = WebParts.WikiText,
                                        Title = "Wiki text",
                                        Row = lastRow,
                                        Column = lastColum,
                                        Order = extraWebPartCounter,
                                        Properties = props
                                    });
                                    extraWebPartCounter++;
                                }
                                else
                                {
                                    split = false;
                                }

                                // Fill properties of the video web part
                                props = new Dictionary<string, string>();
                                if ((element as IHtmlInlineFrameElement).Source != null)
                                {
                                    props.Add("Title", "Video in wiki text");
                                    props.Add("Description", "");
                                    props.Add("IFrameEmbed", (element as IElement).OuterHtml);
                                    props.Add("Source", (element as IHtmlInlineFrameElement).Source);
                                }

                                var allowFullScreen = (element as IElement).Attributes.Where(p => p.Name.Equals("allowfullscreen", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                                if (allowFullScreen != null)
                                {
                                    bool.TryParse(allowFullScreen.Value, out bool val);
                                    props.Add("AllowFullScreen", val.ToString());
                                }
                                var size = (element as IElement).Attributes.Where(p => p.Name.Equals("width", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                                if (size != null)
                                {
                                    props.Add("Width", size.Value.ToString());
                                }
                                size = (element as IElement).Attributes.Where(p => p.Name.Equals("height", StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                                if (size != null)
                                {
                                    props.Add("Height", size.Value.ToString());
                                }

                                // Add video part, in contrast with the image we're always adding the video as a separate web part as the modern text editor strips out embedded videos
                                replacedWebParts.Add(new WebPartEntity()
                                {
                                    Type = WebParts.WikiVideo,
                                    Title = "Video in wiki text",
                                    Row = split ? lastRow : lastRow2,
                                    Column = split ? lastColum : lastColum2,
                                    Order = split ? extraWebPartCounter : extraWebPartCounter2,
                                    Properties = props
                                });

                                if (split)
                                {
                                    // replace img or img nested in A with "splitter"
                                    var splitter = document.CreateElement("span");
                                    splitter.ClassName = "split";

                                    if (element.ParentElement != null)
                                    {
                                        element.ParentElement.ReplaceChild(splitter, element);
                                    }
                                    extraWebPartCounter++;
                                }
                                else
                                {
                                    extraWebPartCounter2++;
                                }
                            }
                        }

                        Dictionary<string, string> props2 = new Dictionary<string, string>();
                        props2.Add("Title", "Wiki text");
                        props2.Add("Text", $"SplitPart{splitCounter}");
                        splitCounter++;

                        replacedWebParts.Add(new WebPartEntity()
                        {
                            Type = WebParts.WikiText,
                            Title = "Wiki text",
                            Row = lastRow,
                            Column = lastColum,
                            Order = extraWebPartCounter,
                            Properties = props2
                        });
                        extraWebPartCounter++;

                        // Fix up WikiText parts
                        // Step 1: get the html now that we've replaced img/iframe tags with a span
                        string preppedWikiText = "";
                        if (document.DocumentElement.Children.Length > 1)
                        {
                            preppedWikiText = document.DocumentElement.Children[1].InnerHtml;
                        }
                        // Step 2: split the html text in parts based upon the span we added
                        string[] splitText = preppedWikiText.Split(new string[] { "<span class=\"split\"></span>" }, StringSplitOptions.RemoveEmptyEntries);

                        // Step 3: use AngleSharp to construct valid html from each part and link it back to the WikiText placeholder web part
                        foreach(var replacedWebPart in replacedWebParts.ToList())
                        {
                            if (replacedWebPart.Type == WebParts.WikiText)
                            {
                                if (Int32.TryParse(replacedWebPart.Properties["Text"].Replace("SplitPart", ""), out int index))
                                {
                                    index = index - 1;

                                    if (splitText.Length >= index + 1)
                                    {
                                        using (var docTemp = parser.ParseDocument(splitText[index]))
                                        {
                                            if (docTemp.DocumentElement.Children.Length > 1)
                                            {
                                                // Remove empty DIV's as that's a net result of the splitting
                                                StripEmptyDivAndPfromHtmlTree(docTemp.DocumentElement.Children[1]);

                                                string updatedText = docTemp.DocumentElement.Children[1].InnerHtml;
                                                replacedWebPart.Properties["Text"] = updatedText;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // no text part for this web part, so delete it. This happens when there was no content anymore below the last img/iframe tag
                                        replacedWebParts.Remove(replacedWebPart);
                                    }
                                }
                            }
                        }

                        // reset counter for next wiki zone
                        splitCounter = 1;

                        // Insert into updated web parts list
                        updatedWebParts.AddRange(replacedWebParts);
                    }
                }
                else
                {
                    // Not a text editor web part, so we simply retain it
                    updatedWebParts.Add(wp);
                }
            }

            // Return the collection of "split" web parts
            return updatedWebParts;
        }

        #region Helper methods
        private string IncludeWebPartTitle(string htmlContent, Dictionary<string,string> webPartProperties)
        {
            if (webPartProperties.ContainsKey("ChromeType") && webPartProperties["ChromeType"] != null)
            {
                string chromeType = webPartProperties["ChromeType"].ToString();

                // Add header for all but the chrometype = none and chrometype = border only
                if (chromeType != "2" && chromeType != "4")
                {
                    // Only add a header if there a title set
                    if (webPartProperties.ContainsKey("Title") && webPartProperties["Title"] != null)
                    {
                        htmlContent = $"<H2>{webPartProperties["Title"].ToString()}</H2><div>{htmlContent}</div>";
                    }
                }
            }

            return htmlContent;
        }

        private void StripEmptyDivAndPfromHtmlTree(IElement newWikiTextHtmlFragment)
        {
            var divs = newWikiTextHtmlFragment.QuerySelectorAll("div");
            var ps = newWikiTextHtmlFragment.QuerySelectorAll("p");
            var list = divs.Union(ps);


            if (list.Any())
            {
                foreach (var el in list)
                {
                    if (!el.HasChildNodes)
                    {
                        el.Remove();
                    }
                }
            }
        }

        private bool InUnSplitableElement(INode node)
        {
            IElement start = null;

            if (!(node is IElement))
            {
                start = node.ParentElement;
            }
            else
            {
                start = node as IElement;
            }

            bool unSplitableElementFound = false;

            while (!unSplitableElementFound)
            {
                if (start.TagName == "TD" || start.TagName == "TR" || start.TagName == "TBODY" || // table
                    start.TagName == "LI" || start.TagName == "UL" || start.TagName == "OL") // lists
                {
                    return true;
                }
                else
                {
                    start = start.ParentElement;
                }

                if (start == null)
                {
                    return false;
                }
            }

            return false;
        }
        #endregion
    }
}
