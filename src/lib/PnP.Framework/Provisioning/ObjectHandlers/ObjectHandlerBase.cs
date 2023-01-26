﻿using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json.Linq;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal abstract class ObjectHandlerBase
    {
        internal bool? _willExtract;
        internal bool? _willProvision;

        private bool _reportProgress = true;
        public abstract string Name { get; }
        public abstract string InternalName { get; }

        public bool ReportProgress
        {
            get { return _reportProgress; }
            set { _reportProgress = value; }
        }

        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        public abstract bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation);

        public abstract bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo);

        public abstract TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation);

        public abstract ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo);

        internal void WriteMessage(string message, ProvisioningMessageType messageType)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate(message, messageType);
            }
        }

        internal void WriteSubProgress(string title, string message, int step, int total)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate($"{title}|{message}|{step}|{total}", ProvisioningMessageType.Progress);
            }
        }

        /// <summary>
        /// Tokenizes calculated fieldXml to use tokens for field references
        /// </summary>
        /// <param name="fieldXml">the xml to tokenize</param>
        /// <returns></returns>
        [Obsolete("Use ObjectField.TokenizeFieldFormula instead. This method produces incorrect tokenization results.")]
        protected string TokenizeFieldFormula(string fieldXml)
        {
            var schemaElement = XElement.Parse(fieldXml);
            var formula = schemaElement.Descendants("Formula").FirstOrDefault();
            var processedFields = new List<string>();
            if (formula != null)
            {
                var formulaString = formula.Value;
                if (formulaString != null)
                {
                    var fieldRefs = schemaElement.Descendants("FieldRef");
                    foreach (var fieldRef in fieldRefs)
                    {
                        var fieldInternalName = fieldRef.Attribute("Name").Value;
                        if (!processedFields.Contains(fieldInternalName))
                        {
                            formulaString = formulaString.Replace(fieldInternalName, $"[{{fieldtitle:{fieldInternalName}}}]");
                            processedFields.Add(fieldInternalName);
                        }
                    }
                    var fieldRefParent = schemaElement.Descendants("FieldRefs");
                    fieldRefParent.Remove();

                }
                formula.Value = formulaString;
            }

            return schemaElement.ToString();
        }

        /// <summary>
        /// Tokenizes the taxonomy field.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        protected string TokenizeTaxonomyField(Web web, XElement element)
        {
            // Replace Taxonomy field references to SspId, TermSetId with tokens
            TaxonomySession session = TaxonomySession.GetTaxonomySession(web.Context);
            TermStore store = session.GetDefaultSiteCollectionTermStore();

            var sspIdElement = element.XPathSelectElement("./Customization/ArrayOfProperty/Property[Name = 'SspId']/Value");
            if (sspIdElement != null)
            {
                sspIdElement.Value = "{sitecollectiontermstoreid}";
            }
            var termSetIdElement = element.XPathSelectElement("./Customization/ArrayOfProperty/Property[Name = 'TermSetId']/Value");
            if (termSetIdElement != null)
            {
                Guid termSetId;
                if (Guid.TryParse(termSetIdElement.Value, out termSetId) && termSetId != Guid.Empty)
                {
                    Microsoft.SharePoint.Client.Taxonomy.TermSet termSet = store.GetTermSet(termSetId);
                    store.Context.ExecuteQueryRetry();

                    if (!termSet.ServerObjectIsNull())
                    {
                        termSet.EnsureProperties(ts => ts.Name, ts => ts.Group.Name, ts => ts.Group.IsSiteCollectionGroup);

                        termSetIdElement.Value = String.Format("{{termsetid:{0}:{1}}}", termSet.Group.IsSiteCollectionGroup ? "{sitecollectiontermgroupname}" : termSet.Group.Name, termSet.Name);
                    }
                }
            }

            return element.ToString();
        }

        /// <summary>
        /// Check if all tokens where replaced. If the field is a taxonomy field then we will check for the values of the referenced termstore and termset. 
        /// </summary>
        /// <param name="fieldXml">The xml to parse</param>
        /// <param name="parser"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected static bool IsFieldXmlValid(string fieldXml, TokenParser parser, ClientRuntimeContext context)
        {
            var isValid = true;
            var leftOverTokens = parser.GetLeftOverTokens(fieldXml);
            if (!leftOverTokens.Any())
            {
                var fieldElement = XElement.Parse(fieldXml);
                if (fieldElement.Attribute("Type").Value == "TaxonomyFieldType")
                {
                    var termStoreIdElement = fieldElement.XPathSelectElement("//ArrayOfProperty/Property[Name='SspId']/Value");
                    if (termStoreIdElement != null)
                    {
                        var termStoreId = Guid.Parse(termStoreIdElement.Value);
                        TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(context);
                        try
                        {
                            taxSession.EnsureProperty(t => t.TermStores);
                            var store = taxSession.TermStores.GetById(termStoreId);
                            context.Load(store);
                            context.ExecuteQueryRetry();
                            if (store.ServerObjectIsNull.HasValue && !store.ServerObjectIsNull.Value)
                            {
                                var termSetIdElement = fieldElement.XPathSelectElement("//ArrayOfProperty/Property[Name='TermSetId']/Value");
                                if (termSetIdElement != null)
                                {
                                    var termSetId = Guid.Parse(termSetIdElement.Value);
                                    try
                                    {
                                        var termSet = store.GetTermSet(termSetId);
                                        context.Load(termSet);
                                        context.ExecuteQueryRetry();
                                        isValid = termSet.ServerObjectIsNull.HasValue && !termSet.ServerObjectIsNull.Value;
                                    }
                                    catch (Exception)
                                    {
                                        isValid = false;
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            isValid = false;
                        }
                    }
                    else
                    {
                        isValid = false;
                    }
                }
            }
            else
            {
                //Some tokens where not replaced
                isValid = false;
            }
            return isValid;
        }

        /// <summary>
        /// Tokenize a XML snippet based list view attribute with {listcontenttypeid:listname,contenttypeid}, {themecatalog} or {masterpagecatalog}
        /// </summary>
        /// <param name="xml">the XML snippet to tokenize as String</param>
        /// <param name="list">List being used</param>
        /// <param name="web">Web being used</param>
        /// <returns>tokenized xml as String</returns>
        protected string TokenizeListView(string xml, List list, Web web = null)
        {
            if (string.IsNullOrEmpty(xml))
            {
                // nothing to tokenize...
                return string.Empty;
            }
            else
            {
                //Fix since RegEx.Escape in ListContentTypeIdToken.CreateToken does escape whiteSpace with \ it breaks the JSON in NewDocumentTemplates
                string NewDocumentTemplatesJson = null;
                if (xml.Contains("NewDocumentTemplates"))
                {
                    var viewSchema = System.Xml.Linq.XDocument.Parse(xml);
                    var templateElement = viewSchema.Root.Elements().FirstOrDefault(element => element.Name.LocalName == "NewDocumentTemplates");
                    if (templateElement != null)
                    {
                        NewDocumentTemplatesJson = templateElement.Value;
                    }
                }

                foreach (Microsoft.SharePoint.Client.ContentType contentType in list.ContentTypes)
                {
                    string contentTypeReplacement = ListContentTypeIdToken.CreateToken(list.Title, contentType.Id);
                    if (!string.IsNullOrWhiteSpace(NewDocumentTemplatesJson))
                    {
                        string contentTypeReplacementJson = contentTypeReplacement.Replace("\\ ", " ");
                        contentTypeReplacementJson = contentTypeReplacementJson.Replace("\\", "\\\\");
                        NewDocumentTemplatesJson = Regex.Replace(NewDocumentTemplatesJson, contentType.Id.StringValue, contentTypeReplacementJson, RegexOptions.IgnoreCase);
                    }
                    xml = Regex.Replace(xml, contentType.Id.StringValue, contentTypeReplacement, RegexOptions.IgnoreCase);
                }

                if(!string.IsNullOrWhiteSpace(NewDocumentTemplatesJson))
                {
                    web.EnsureProperties(w => w.ServerRelativeUrl);
                    try
                    {
                        var JObjNewDocTemplate = JArray.Parse(NewDocumentTemplatesJson);
                        var fileUrls = JObjNewDocTemplate.SelectTokens("..url");
                        foreach (var templateFile in fileUrls)
                        {
                            var templateObj = templateFile.Parent.Parent as JObject;
                            string orgValue = (string)templateObj["url"];
                            if (!string.IsNullOrWhiteSpace(orgValue))
                            {
                                if (web.ServerRelativeUrl == "/")
                                {
                                    templateObj["url"] = $"{{site}}/{orgValue.TrimStart('/')}";
                                }
                                else
                                {
                                    templateObj["url"] = Regex.Replace(orgValue, web.ServerRelativeUrl.TrimEnd('/'), "{site}", RegexOptions.IgnoreCase); ;
                                }
                            }
                        }
                        NewDocumentTemplatesJson = JObjNewDocTemplate.ToString(Newtonsoft.Json.Formatting.None);
                    }
                    catch { }
                }

                string tokenizedXML = TokenizeXml(xml, web);
                if (!string.IsNullOrWhiteSpace(NewDocumentTemplatesJson))
                {
                    var viewSchema = System.Xml.Linq.XDocument.Parse(tokenizedXML);
                    var templateElement = viewSchema.Root.Elements().FirstOrDefault(element => element.Name.LocalName == "NewDocumentTemplates");
                    if (templateElement != null)
                    {
                        templateElement.Value = NewDocumentTemplatesJson;
                        tokenizedXML = viewSchema.ToString();
                    }
                }

                return tokenizedXML;
            }
        }

        /// <summary>
        /// Tokenize a XML snippet based attribute with {themecatalog} or {masterpagecatalog}
        /// </summary>
        /// <param name="xml">the XML snippet to tokenize as String</param>
        /// <param name="web">Web being used</param>
        /// <returns>tokenized xml as String</returns>
        protected string TokenizeXml(string xml, Web web = null)
        {
            if (string.IsNullOrEmpty(xml))
            {
                // nothing to tokenize...
                return string.Empty;
            }
            else
            {
                var subsite = false;
                if (web != null)
                {
                    subsite = web.IsSubSite();
                }
                web.EnsureProperty(w => w.ServerRelativeUrl);
                // Theme Catalog
                var themeRegex = new Regex(@"(?<theme>\/_catalogs\/theme)");
                xml = themeRegex.Replace(xml, subsite ? "{sitecollection}/_catalogs/theme" : "{themecatalog}");

                // Master Page Catalog
                var masterPageRegex = new Regex(@"(?<masterpage>\/_catalogs\/masterpage)");
                xml = masterPageRegex.Replace(xml, subsite ? "{sitecollection}/_catalogs/masterpage" : "{masterpagecatalog}");

                // Site
                var siteRegexReplacement = "{site}";
                // If we are in the root site collection with just / as ServerRelativeUrl then we cannot replace all / with {site}, otherwise the urls will look like "{site}_layouts/15/images"
                if (web.ServerRelativeUrl == "/")
                    siteRegexReplacement += "/";

                xml = Regex.Replace(xml, "(\"" + web.ServerRelativeUrl + ")(?!&)", "\"" + siteRegexReplacement, RegexOptions.IgnoreCase);
                xml = Regex.Replace(xml, "'" + web.ServerRelativeUrl, "'" + siteRegexReplacement, RegexOptions.IgnoreCase);
                xml = Regex.Replace(xml, ">" + web.ServerRelativeUrl, ">" + siteRegexReplacement, RegexOptions.IgnoreCase);

                return xml;
            }
        }

        /// <summary>
        /// Tokenize a template item url based attribute with {themecatalog} or {masterpagecatalog} or {site}+
        /// </summary>
        /// <param name="url">the url to tokenize as String</param>
        /// <param name="webUrl">web url of the actual web as String</param>
        /// <param name="web">Web being used</param>
        /// <returns>tokenized url as String</returns>
        protected string Tokenize(string url, string webUrl, Web web = null)
        {
            String result = null;

            if (string.IsNullOrEmpty(url))
            {
                // nothing to tokenize...
                result = String.Empty;
            }
            else
            {
                // Decode URL
                url = Uri.UnescapeDataString(url);
                // Try with theme catalog
                if (url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    var subsite = false;
                    if (web != null)
                    {
                        subsite = web.IsSubSite();
                    }
                    if (subsite)
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/theme", "{sitecollection}/_catalogs/theme");
                    }
                    else
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/theme", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/theme", "{themecatalog}");
                    }
                }

                // Try with master page catalog
                if (url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    var subsite = false;
                    if (web != null)
                    {
                        subsite = web.IsSubSite();
                    }
                    if (subsite)
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/masterpage", "{sitecollection}/_catalogs/masterpage");
                    }
                    else
                    {
                        result = url.Substring(url.IndexOf("/_catalogs/masterpage", StringComparison.InvariantCultureIgnoreCase)).Replace("/_catalogs/masterpage", "{masterpagecatalog}");
                    }
                }

                // Try with site URL
                if (result != null)
                {
                    url = result;
                }
                Uri uri;
                if (Uri.TryCreate(webUrl, UriKind.Absolute, out uri))
                {
                    string webUrlPathAndQuery = Uri.UnescapeDataString(uri.PathAndQuery);
                    // Don't do additional replacement when masterpagecatalog and themecatalog (see #675)
                    if (url.IndexOf(webUrlPathAndQuery, StringComparison.InvariantCultureIgnoreCase) > -1 && (url.IndexOf("{masterpagecatalog}") == -1) && (url.IndexOf("{themecatalog}") == -1))
                    {
                        result = (uri.PathAndQuery.Equals("/") && url.StartsWith(uri.PathAndQuery))
                            ? "{site}" + url // we need this for DocumentTemplate attribute of pnp:ListInstance also on a root site ("/") without managed path
                            : url.Replace(webUrlPathAndQuery, "{site}");
                    }
                }

                // Default action
                if (String.IsNullOrEmpty(result))
                {
                    result = url;
                }
            }

            return (result);
        }
    }
}
