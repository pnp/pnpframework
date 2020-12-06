using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Utilities;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPageContents : ObjectContentHandlerBase
    {
        public override string Name
        {
            get { return "Page Contents"; }
        }

        public override string InternalName => "PageContents";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // This handler only extracts contents and adds them to the Files and Pages collection.
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Extract the Home Page
                web.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);

                var homepageUrl = web.RootFolder.WelcomePage;
                if (string.IsNullOrEmpty(homepageUrl))
                {
                    homepageUrl = "Default.aspx";
                }
                var welcomePageUrl = UrlUtility.Combine(web.ServerRelativeUrl, homepageUrl);

                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(welcomePageUrl));
                try
                {
                    var listItem = file.EnsureProperty(f => f.ListItemAllFields);
                    if (listItem != null)
                    {
                        if (listItem.FieldValues.ContainsKey("WikiField") && listItem.FieldValues["WikiField"] != null)
                        {
                            // Wiki page
                            var fullUri = new Uri(UrlUtility.Combine(web.Url, web.RootFolder.WelcomePage));

                            //var folderPath = fullUri.Segments.Take(fullUri.Segments.Count() - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
                            //var fileName = fullUri.Segments[fullUri.Segments.Count() - 1];

                            var homeFile = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(welcomePageUrl));

                            var limitedWPManager = homeFile.GetLimitedWebPartManager(PersonalizationScope.Shared);

                            web.Context.Load(limitedWPManager);

                            //var webParts = web.GetWebParts(welcomePageUrl);

                            var page = new Page()
                            {
                                Layout = WikiPageLayout.Custom,
                                Overwrite = true,
                                Url = Tokenize(fullUri.PathAndQuery, web.Url),
                            };
                            var pageContents = listItem.FieldValues["WikiField"].ToString();

                            Regex regexClientIds = new Regex(@"id=\""div_(?<ControlId>(\w|\-)+)");
                            if (regexClientIds.IsMatch(pageContents))
                            {
                                foreach (Match webPartMatch in regexClientIds.Matches(pageContents))
                                {
                                    String serverSideControlId = webPartMatch.Groups["ControlId"].Value;

                                    try
                                    {
                                        var serverSideControlIdToSearchFor =
                                            $"g_{serverSideControlId.Replace("-", "_")}";

                                        var webPart = limitedWPManager.WebParts.GetByControlId(serverSideControlIdToSearchFor);
                                        web.Context.Load(webPart,
                                            wp => wp.Id,
                                            wp => wp.WebPart.Title,
                                            wp => wp.WebPart.ZoneIndex
                                            );
                                        web.Context.ExecuteQueryRetry();

                                        var webPartxml = TokenizeWebPartXml(web, web.GetWebPartXml(webPart.Id, welcomePageUrl));

                                        page.WebParts.Add(new Model.WebPart()
                                        {
                                            Title = webPart.WebPart.Title,
                                            Contents = webPartxml,
                                            Order = (uint)webPart.WebPart.ZoneIndex,
                                            Row = 1, // By default we will create a onecolumn layout, add the webpart to it, and later replace the wikifield on the page to position the webparts correctly.
                                            Column = 1 // By default we will create a onecolumn layout, add the webpart to it, and later replace the wikifield on the page to position the webparts correctly.
                                        });

                                        pageContents = Regex.Replace(pageContents, serverSideControlId, $"{{webpartid:{webPart.WebPart.Title}}}", RegexOptions.IgnoreCase);
                                    }
                                    catch (ServerException)
                                    {
                                        scope.LogWarning("Found a WebPart ID which is not available on the server-side. ID: {0}", serverSideControlId);
                                    }
                                }
                            }

                            page.Fields.Add("WikiField", pageContents);
                            template.Pages.Add(page);

                            // Set the homepage
                            if (template.WebSettings == null)
                            {
                                template.WebSettings = new WebSettings();
                            }
                            template.WebSettings.WelcomePage = homepageUrl;


                        }
                        else if (listItem.FieldValues.ContainsKey("ClientSideApplicationId") && listItem.FieldValues["ClientSideApplicationId"] != null && listItem.FieldValues["ClientSideApplicationId"].ToString().ToLower() == "b6917cb1-93a0-4b97-a84d-7cf49975d4ec")
                        {
                            // this is a client side page, so let's skip it since it's handled by the Client Side Page contents handler
                        }
                        else
                        {                            
                            // Not a wikipage
                            template = GetFileContents(web, template, welcomePageUrl, creationInfo, scope);
                            if (template.WebSettings == null)
                            {
                                template.WebSettings = new WebSettings();
                            }
                            template.WebSettings.WelcomePage = homepageUrl;                            
                        }
                    }
                }
                catch (ServerException ex)
                {

                    //ignore this error. The default page is not a page but a list view.
                    if (ex.ServerErrorCode != -2146232832 && ex.HResult != -2146233088)
                    {
                        throw;
                    }
                    else
                    {
                        if (ex.HResult != -2146233088)
                        {                            
                            // Page does not belong to a list, extract the file as is
                            template = GetFileContents(web, template, welcomePageUrl, creationInfo, scope);
                            if (template.WebSettings == null)
                            {
                                template.WebSettings = new WebSettings();
                            }
                            template.WebSettings.WelcomePage = homepageUrl;                            
                        }
                    }
                }

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate GetFileContents(Web web, ProvisioningTemplate template, string welcomePageUrl, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope)
        {
            var homepageUrl = web.RootFolder.WelcomePage;
            if (string.IsNullOrEmpty(homepageUrl))
            {
                homepageUrl = "Default.aspx";
            }

            var fullUri = new Uri(UrlUtility.Combine(web.Url, homepageUrl));

            var folderPath = fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
            var fileName = fullUri.Segments[fullUri.Segments.Length - 1];

            var webParts = web.GetWebParts(welcomePageUrl);

            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(welcomePageUrl));

            file.EnsureProperty(f => f.Level);

            var containerPath = folderPath.StartsWith(web.ServerRelativeUrl) && web.ServerRelativeUrl != "/"
                ? folderPath.Substring(web.ServerRelativeUrl.Length)
                : folderPath;
            var container = containerPath.Trim('/').Replace("%20", " ").Replace("/", "\\");

            var homeFile = new Model.File()
            {
                Folder = Tokenize(folderPath, web.Url),
                Src = !string.IsNullOrEmpty(container) ? $"{container}\\{fileName}" : fileName,
                Overwrite = true,
                Level = (Model.FileLevel)Enum.Parse(typeof(Model.FileLevel), file.Level.ToString())
            };

            // Add field values to file

            RetrieveFieldValues(web, file, homeFile);

            // Add WebParts to file
            foreach (var webPart in webParts)
            {
                var webPartxml = TokenizeWebPartXml(web, web.GetWebPartXml(webPart.Id, welcomePageUrl));

                Model.WebPart newWp = new Model.WebPart()
                {
                    Title = webPart.WebPart.Title,
                    Row = (uint)webPart.WebPart.ZoneIndex,
                    Order = (uint)webPart.WebPart.ZoneIndex,
                    Contents = webPartxml,
                    Zone = webPart.ZoneId
                };

                homeFile.WebParts.Add(newWp);
            }
            template.Files.Add(homeFile);

            // Persist file using connector
            if (creationInfo.PersistBrandingFiles)
            {
                PersistFile(web, creationInfo, scope, folderPath, fileName);
            }
            return template;
        }

        private static string TokenizeWebPartXml(Web web, string xml)
        {
            var lists = web.Lists;
            web.Context.Load(web, w => w.ServerRelativeUrl, w => w.Id);
            web.Context.Load(lists, ls => ls.Include(l => l.Id, l => l.Title));
            web.Context.ExecuteQueryRetry();

            foreach (var list in lists)
            {
                xml = Regex.Replace(xml, list.Id.ToString(), $"{{listid:{System.Security.SecurityElement.Escape(list.Title)}}}", RegexOptions.IgnoreCase);
            }

            string webpartType = null;
            try
            {
                var xmlDoc = new System.Xml.XmlDocument();
                xmlDoc.LoadXml(xml);
                XmlNamespaceManager ns = new XmlNamespaceManager(xmlDoc.NameTable);
                ns.AddNamespace("ns", "http://schemas.microsoft.com/WebPart/v3");
                webpartType = xmlDoc.SelectSingleNode("//webParts/ns:webPart/ns:metaData/ns:type[@name]", ns)?.Attributes["name"]?.Value;
            }
            catch (Exception) { }

            //some webparts already contains the site URL using ~sitecollection token (i.e: CQWP)
            //xml = Regex.Replace(xml, "\"~sitecollection/(.)*\"", "\"{site}\"", RegexOptions.IgnoreCase);
            //xml = Regex.Replace(xml, "'~sitecollection/(.)*'", "'{site}'", RegexOptions.IgnoreCase);

            // Support for ContentBySearchWebParts
            //if (!string.IsNullOrEmpty(webpartType) && webpartType.ToLower().Contains("Microsoft.Office.Server.Search.WebControls.ContentBySearchWebPart".ToLower()))
            //    xml = Regex.Replace(xml, ">~sitecollection/(.)*<", (Match m) => m.ToString().Replace("~sitecollection", "{sitecollection}"), RegexOptions.IgnoreCase);
            //else
            //    xml = Regex.Replace(xml, ">~sitecollection/(.)*<", ">{site}<", RegexOptions.IgnoreCase);

            xml = Regex.Replace(xml, web.Id.ToString(), "{siteid}", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, "(\"" + web.ServerRelativeUrl + ")(?!&)", "\"{site}", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, "'" + web.ServerRelativeUrl, "'{site}", RegexOptions.IgnoreCase);
            xml = Regex.Replace(xml, ">" + web.ServerRelativeUrl, ">{site}", RegexOptions.IgnoreCase);

            return xml;
        }

        private static ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }

    }
}
