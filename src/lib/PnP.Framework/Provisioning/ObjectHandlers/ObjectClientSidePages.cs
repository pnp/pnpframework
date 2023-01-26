﻿using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Core.Services;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.Extensions;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;
using PnP.Framework.Utilities;
using PnP.Framework.Utilities.CanvasControl;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using PnPCore = PnP.Core.Model.SharePoint;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectClientSidePages : ObjectHandlerBase
    {
        private PnPContext pnpContext;
        private PnPCore.IPage dummyPage;
        private const string ContentTypeIdField = "ContentTypeId";
        private const string FileRefField = "FileRef";
        private const string SPSitePageFlagsField = "_SPSitePageFlags";
        private static readonly Guid MultilingualPagesFeature = new Guid("24611c05-ee19-45da-955f-6602264abaf8");
        private static readonly Guid MixedRealityFeature = new Guid("2ac9c540-6db4-4155-892c-3273957f1926");

        public override string Name
        {
            get { return "ClientSidePages"; }
        }

        public override string InternalName => "ClientSidePages";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                pnpContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext);
                dummyPage = pnpContext.Web.NewPage();

                web.EnsureProperties(w => w.ServerRelativeUrl);

                // determine pages library
                string pagesLibrary = "SitePages";

                var pagesLibraryList = web.GetListByUrl(pagesLibrary, p => p.RootFolder);

                List<string> preCreatedPages = new List<string>();

                // Ensure the needed languages are enabled on the site
                EnsureWebLanguages(web, template, scope);
                // Ensure spaces is enabled
                EnsureSpaces(web, template, scope);

                var currentPageIndex = 0;
                // pre create the needed pages so we can fill the needed tokens which might be used later on when we put web parts on those pages
                foreach (var clientSidePage in template.ClientSidePages)
                {
                    var preCreatedPage = PreCreatePage(web, template, parser, clientSidePage, pagesLibrary, pagesLibraryList, ref currentPageIndex);
                    if (preCreatedPage != null)
                    {
                        preCreatedPages.Add(preCreatedPage);
                    }

                    if (clientSidePage.Translations.Any())
                    {
                        //Pages.ClientSidePage page = null;
                        PnPCore.IPage page = null;
                        string pageName = DeterminePageName(parser, clientSidePage);
                        if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                        {
                            // Get the existing template page
                            //page = web.LoadClientSidePage($"{Pages.ClientSidePage.GetTemplatesFolder(pagesLibraryList)}/{pageName}");
                            page = web.LoadClientSidePage($"{dummyPage.GetTemplatesFolder()}/{pageName}");
                        }
                        else
                        {
                            // Get the existing page
                            page = web.LoadClientSidePage(pageName);
                        }

                        if (page != null)
                        {

                            //Pages.TranslationStatusCollection availableTranslations = page.Translations();
                            var availableTranslations = page.GetPageTranslations();

                            // Trigger the creation of the translated pages
                            //Pages.TranslationStatusCreationRequest tscr = new Pages.TranslationStatusCreationRequest();
                            PnPCore.PageTranslationOptions tscr = new PnPCore.PageTranslationOptions();
                            foreach (var translatedClientSidePage in clientSidePage.Translations)
                            {
                                //if (availableTranslations.Items.Where(p => p.Culture.Equals(new CultureInfo(translatedClientSidePage.LCID).Name, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault() == null)
                                if (availableTranslations.TranslatedLanguages.Where(p => p.Culture.Equals(new CultureInfo(translatedClientSidePage.LCID).Name, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault() == null)
                                {
                                    tscr.AddLanguage(translatedClientSidePage.LCID);
                                }
                            }

                            // Pages.TranslationStatusCollection translationResults = null;
                            PnPCore.IPageTranslationStatusCollection translationResults = null;
                            if (tscr.LanguageCodes != null && tscr.LanguageCodes.Count > 0)
                            {
                                //translationResults = page.GenerateTranslations(tscr);
                                translationResults = page.TranslatePages(tscr);
                            }

                            //IEnumerable<Pages.TranslationStatus> combinedTranslationResults = new List<Pages.TranslationStatus>();
                            IEnumerable<PnPCore.IPageTranslationStatus> combinedTranslationResults = new List<PnPCore.IPageTranslationStatus>();

                            // Translation results will contain all available pages when ran
                            //if (translationResults != null && translationResults.Items.Count > 0)
                            if (translationResults != null && translationResults.TranslatedLanguages.Count > 0)
                            {
                                //combinedTranslationResults = combinedTranslationResults.Union(translationResults.Items);
                                combinedTranslationResults = combinedTranslationResults.Union(translationResults.TranslatedLanguages);
                            }
                            // No new translations generated, so take what we got as available translations
                            //else if (availableTranslations != null && availableTranslations.Items.Count > 0)
                            else if (availableTranslations != null && availableTranslations.TranslatedLanguages.Count > 0)
                            {
                                //combinedTranslationResults = combinedTranslationResults.Union(availableTranslations.Items);
                                combinedTranslationResults = combinedTranslationResults.Union(availableTranslations.TranslatedLanguages);
                            }

                            foreach (var createdTranslation in combinedTranslationResults)
                            {
                                //string url = UrlUtility.Combine(web.ServerRelativeUrl, createdTranslation.Path.DecodedUrl);
                                string url = UrlUtility.Combine(web.ServerRelativeUrl, createdTranslation.Path);
                                preCreatedPages.Add(url);
                                // Load up page tokens for these translations
                                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                                web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath);
                                web.Context.ExecuteQueryRetry();

                                // Fill token
                                var pageUrlForToken = file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray());
                                parser.AddToken(new PageUniqueIdToken(web, pageUrlForToken, file.UniqueId));
                                parser.AddToken(new PageUniqueIdEncodedToken(web, pageUrlForToken, file.UniqueId));
                            }
                        }
                    }
                }

                currentPageIndex = 0;
                // Iterate over the pages and create/update them
                foreach (var clientSidePage in template.ClientSidePages)
                {
                    CreatePage(web, template, parser, scope, clientSidePage, pagesLibrary, pagesLibraryList, ref currentPageIndex, preCreatedPages);

                    if (clientSidePage.Translations.Any())
                    {
                        foreach (var translatedClientSidePage in clientSidePage.Translations)
                        {
                            CreatePage(web, template, parser, scope, translatedClientSidePage, pagesLibrary, pagesLibraryList, ref currentPageIndex, preCreatedPages);
                        }
                    }
                }
            }

            WriteMessage("Done processing Client Side Pages", ProvisioningMessageType.Completed);
            return parser;
        }

        private static void EnsureSpaces(Web web, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            var spacesPages = template.ClientSidePages.Where(p => p.Layout != null && p.Layout.Equals(PnPCore.PageLayoutType.Spaces.ToString(), StringComparison.InvariantCultureIgnoreCase));
            if (spacesPages.Any())
            {
                try
                {
                    // Enable the MUI feature
                    web.ActivateFeature(ObjectClientSidePages.MixedRealityFeature);
                }
                catch (Exception ex)
                {
                    scope.LogError($"Mixed reality feature could not be enabled: {ex.Message}");
                    throw;
                }
            }
        }

        private static void EnsureWebLanguages(Web web, ProvisioningTemplate template, PnPMonitoredScope scope)
        {
            List<int> neededLanguages = new List<int>();
            int neededSourceLanguage = 0;

            foreach (var page in template.ClientSidePages.Where(p => p.Translations.Any()))
            {
                if (neededSourceLanguage == 0)
                {
                    neededSourceLanguage = page.LCID > 0 ? page.LCID : (template.RegionalSettings != null ? template.RegionalSettings.LocaleId : 0);
                }
                else
                {
                    // Source language should be the same for all pages in the template
                    if (neededSourceLanguage != page.LCID)
                    {
                        string error = "The pages in this template are based upon multiple source languages while all pages in a site must have the same source language";
                        scope.LogError(error);
                        throw new Exception(error);
                    }
                }

                foreach (var translatedPage in page.Translations)
                {
                    if (!neededLanguages.Contains(translatedPage.LCID))
                    {
                        neededLanguages.Add(translatedPage.LCID);
                    }
                }
            }

            // No translations found, bail out
            if (neededLanguages.Count == 0)
            {
                return;
            }

            try
            {
                // Enable the MUI feature
                web.ActivateFeature(ObjectClientSidePages.MultilingualPagesFeature);
            }
            catch (Exception ex)
            {
                scope.LogError($"Multilingual pages feature could not be enabled: {ex.Message}");
                throw;
            }

            // Check the "source" language
            web.EnsureProperties(p => p.Language, p => p.IsMultilingual);
            int sourceLanguage = Convert.ToInt32(web.Language);
            if (sourceLanguage != neededSourceLanguage)
            {
                string error = $"The web has source language {sourceLanguage} while the template expects {neededSourceLanguage}";
                scope.LogError(error);
                throw new Exception(error);
            }

            // Ensure the needed languages are available on this site
            if (!web.IsMultilingual)
            {
                web.IsMultilingual = true;
                web.Context.Load(web, w => w.SupportedUILanguageIds);
                web.Update();
            }
            else
            {
                web.Context.Load(web, w => w.SupportedUILanguageIds);
            }
            web.Context.ExecuteQueryRetry();

            var supportedLanguages = web.SupportedUILanguageIds;
            bool languageAdded = false;
            foreach (var language in neededLanguages)
            {
                if (!supportedLanguages.Contains(language))
                {
                    web.AddSupportedUILanguage(language);
                    languageAdded = true;
                }
            }

            if (languageAdded)
            {
                web.Update();
                web.Context.ExecuteQueryRetry();
            }
        }

        private void CreatePage(Web web, ProvisioningTemplate template, TokenParser parser, PnPMonitoredScope scope, BaseClientSidePage clientSidePage, string pagesLibrary, List pagesLibraryList, ref int currentPageIndex, List<string> preCreatedPages)
        {
            string pageName = DeterminePageName(parser, clientSidePage);
            string url = $"{pagesLibrary}/{pageName}";

            if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
            {
                if (clientSidePage is TranslatedClientSidePage)
                {
                    url = $"{pagesLibrary}/{pageName}";
                }
                else
                {
                    //url = $"{pagesLibrary}/{Pages.ClientSidePage.GetTemplatesFolder(pagesLibraryList)}/{pageName}";
                    url = $"{pagesLibrary}/{dummyPage.GetTemplatesFolder()}/{pageName}";
                }
            }

            // Write page level status messages, needed in case many pages are provisioned
            currentPageIndex++;
            int totalPages = 0;
            foreach (var p in template.ClientSidePages)
            {
                totalPages++;
                if (p.Translations.Any())
                {
                    totalPages += p.Translations.Count;
                }
            }
            WriteSubProgress("Provision ClientSidePage", pageName, currentPageIndex, totalPages);

            url = UrlUtility.Combine(web.ServerRelativeUrl, url);

            var exists = true;
            try
            {
                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                web.Context.Load(file);
                web.Context.ExecuteQueryRetry();
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    exists = false;
                }
            }

            PnPCore.IPage page = null;
            if (exists)
            {
                if (clientSidePage.Overwrite || preCreatedPages.Contains(url))
                {
                    if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                    {
                        // Get the existing template page
                        if (clientSidePage is TranslatedClientSidePage)
                        {
                            page = web.LoadClientSidePage($"{pageName}");
                        }
                        else
                        {
                            page = web.LoadClientSidePage($"{dummyPage.GetTemplatesFolder()}/{pageName}");
                        }
                    }
                    else
                    {
                        // Get the existing page
                        page = web.LoadClientSidePage(pageName);
                    }

                    // Clear the page
                    page.ClearPage();
                }
                else
                {
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePages_NoOverWrite, pageName);
                    return;
                }
            }
            else
            {
                // Create new client side page
                page = web.AddClientSidePage(pageName);
            }

            // Set page title
            string newTitle = parser.ParseString(clientSidePage.Title);
            if (page.PageTitle != newTitle)
            {
                page.PageTitle = newTitle;
            }

            // Set page layout
            if (!string.IsNullOrEmpty(clientSidePage.Layout))
            {
                page.LayoutType = (PnPCore.PageLayoutType)Enum.Parse(typeof(PnPCore.PageLayoutType), clientSidePage.Layout);
            }

            // Page Header
            if (clientSidePage.Header != null && page.LayoutType != PnPCore.PageLayoutType.Topic)
            {
                switch (clientSidePage.Header.Type)
                {
                    case ClientSidePageHeaderType.None:
                        {
                            page.RemovePageHeader();
                            break;
                        }
                    case ClientSidePageHeaderType.Default:
                        {
                            page.SetDefaultPageHeader();
                            break;
                        }
                    case ClientSidePageHeaderType.Custom:
                        {
                            var serverRelativeImageUrl = parser.ParseString(clientSidePage.Header.ServerRelativeImageUrl);
                            if (clientSidePage.Header.TranslateX.HasValue && clientSidePage.Header.TranslateY.HasValue)
                            {
                                page.SetCustomPageHeader(serverRelativeImageUrl, clientSidePage.Header.TranslateX.Value, clientSidePage.Header.TranslateY.Value);
                            }
                            else
                            {
                                page.SetCustomPageHeader(serverRelativeImageUrl);
                            }

                            page.PageHeader.LayoutType = (PnPCore.PageHeaderLayoutType)Enum.Parse(typeof(PnPCore.PageHeaderLayoutType), clientSidePage.Header.LayoutType.ToString());
                            page.PageHeader.TextAlignment = (PnPCore.PageHeaderTitleAlignment)Enum.Parse(typeof(PnPCore.PageHeaderTitleAlignment), clientSidePage.Header.TextAlignment.ToString());
                            page.PageHeader.ShowTopicHeader = clientSidePage.Header.ShowTopicHeader;
                            page.PageHeader.ShowPublishDate = clientSidePage.Header.ShowPublishDate;
                            page.PageHeader.TopicHeader = parser.ParseString(clientSidePage.Header.TopicHeader);
                            page.PageHeader.AlternativeText = parser.ParseString(clientSidePage.Header.AlternativeText);
                            page.PageHeader.Authors = clientSidePage.Header.Authors;
                            page.PageHeader.AuthorByLine = clientSidePage.Header.AuthorByLine;
                            page.PageHeader.AuthorByLineId = clientSidePage.Header.AuthorByLineId;
                            break;
                        }
                }
            }

            if (!string.IsNullOrEmpty(clientSidePage.ThumbnailUrl))
            {
                page.ThumbnailUrl = parser.ParseString(clientSidePage.ThumbnailUrl);
            }

            // Add content on the page, not needed for repost pages
            if (page.LayoutType != PnPCore.PageLayoutType.RepostPage)
            {
                // Load existing available controls
                var componentsToAdd = page.AvailablePageComponents();

                // if no section specified then add a default single column section
                if (!clientSidePage.Sections.Any())
                {
                    clientSidePage.Sections.Add(new CanvasSection() { Type = CanvasSectionType.OneColumn, Order = 10 });
                }

                int sectionCount = -1;
                // Apply the "layout" and content
                foreach (var section in clientSidePage.Sections)
                {
                    // Skip topic page header control section
                    if (section.Order == 999999)
                    {
                        continue;
                    }

                    sectionCount++;
                    switch (section.Type)
                    {
                        case CanvasSectionType.OneColumn:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.OneColumnFullWidth:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumnFullWidth, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.TwoColumn:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.ThreeColumn:
                            page.AddSection(PnPCore.CanvasSectionTemplate.ThreeColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnLeft:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnLeft, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnRight:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnRight, section.Order, (int)section.BackgroundEmphasis);
                            break;
                        case CanvasSectionType.OneColumnVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumnVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnLeftVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnLeftVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.TwoColumnRightVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.TwoColumnRightVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        case CanvasSectionType.ThreeColumnVerticalSection:
                            page.AddSection(PnPCore.CanvasSectionTemplate.ThreeColumnVerticalSection, section.Order, (int)section.BackgroundEmphasis, (int)section.VerticalSectionEmphasis);
                            break;
                        default:
                            page.AddSection(PnPCore.CanvasSectionTemplate.OneColumn, section.Order, (int)section.BackgroundEmphasis);
                            break;
                    }

                    // Configure collapsible section, if needed
                    if (section.Collapsible)
                    {
                        var targetSection = page.Sections[sectionCount];
                        targetSection.Collapsible = section.Collapsible;
                        targetSection.IsExpanded = section.IsExpanded;
                        targetSection.DisplayName = section.DisplayName;
                        targetSection.IconAlignment = (PnP.Core.Model.SharePoint.IconAlignment)Enum.Parse(
                            typeof(PnP.Core.Model.SharePoint.IconAlignment), 
                            section.IconAlignment.ToString());
                        targetSection.ShowDividerLine = section.ShowDividerLine;
                    }

                    // Add controls to the section
                    if (section.Controls.Any())
                    {
                        // Safety measure: reset column order to 1 for columns marked with 0 or lower
                        foreach (var control in section.Controls.Where(p => p.Column <= 0).ToList())
                        {
                            control.Column = 1;
                        }

                        foreach (CanvasControl control in section.Controls)
                        {
                            PnPCore.IPageComponent baseControl = null;

                            // Is it a text control?
                            if (control.Type == WebPartType.Text)
                            {
                                var textControl = page.NewTextPart();

                                if (control.ControlProperties.Any())
                                {
                                    var textProperty = control.ControlProperties.First();
                                    textControl.Text = parser.ParseString(textProperty.Value);
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        var json = JsonConvert.DeserializeObject<Dictionary<string, string>>(control.JsonControlData);

                                        if (json.Count > 0)
                                        {
                                            textControl.Text = parser.ParseString(json.First().Value);
                                        }
                                    }
                                }
                                // Reduce column number by 1 due 0 start indexing
                                page.AddControl(textControl, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                            }
                            // It is a web part
                            else
                            {
                                // apply token parsing on the web part properties
                                control.JsonControlData = parser.ParseString(control.JsonControlData);

                                // perform processing of web part properties (e.g. include listid property based list title property)
                                var webPartPostProcessor = CanvasControlPostProcessorFactory.Resolve(control);
                                webPartPostProcessor.Process(control, web.Context as ClientContext);

                                // Is a custom developed client side web part (3rd party)
                                if (control.Type == WebPartType.Custom)
                                {
                                    if (!string.IsNullOrEmpty(control.CustomWebPartName))
                                    {
                                        baseControl = componentsToAdd.FirstOrDefault(p => p.Name.Equals(control.CustomWebPartName, StringComparison.InvariantCultureIgnoreCase));
                                    }
                                    else if (control.ControlId != Guid.Empty)
                                    {
                                        baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals($"{{{control.ControlId}}}", StringComparison.CurrentCultureIgnoreCase));

                                        if (baseControl == null)
                                        {
                                            baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals(control.ControlId.ToString(), StringComparison.InvariantCultureIgnoreCase));
                                        }
                                    }
                                }
                                // Is an OOB client side web part (1st party)
                                else
                                {
                                    string webPartName = "";
                                    switch (control.Type)
                                    {
                                        case WebPartType.Image:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Image);
                                            break;
                                        case WebPartType.BingMap:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.BingMap);
                                            break;
                                        case WebPartType.Button:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Button);
                                            break;
                                        case WebPartType.CallToAction:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.CallToAction);
                                            break;
                                        case WebPartType.GroupCalendar:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.GroupCalendar);
                                            break;
                                        case WebPartType.News:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.News);
                                            break;
                                        case WebPartType.PowerBIReportEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.PowerBIReportEmbed);
                                            break;
                                        case WebPartType.Sites:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Sites);
                                            break;
                                        case WebPartType.MicrosoftForms:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.MicrosoftForms);
                                            break;
                                        case WebPartType.ClientWebPart:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ClientWebPart);
                                            break;
                                        case WebPartType.ContentEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ContentEmbed);
                                            break;
                                        case WebPartType.ContentRollup:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ContentRollup);
                                            break;
                                        case WebPartType.DocumentEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.DocumentEmbed);
                                            break;
                                        case WebPartType.Events:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Events);
                                            break;
                                        case WebPartType.Hero:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Hero);
                                            break;
                                        case WebPartType.ImageGallery:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.ImageGallery);
                                            break;
                                        case WebPartType.LinkPreview:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.LinkPreview);
                                            break;
                                        case WebPartType.List:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.List);
                                            break;
                                        case WebPartType.NewsFeed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.NewsFeed);
                                            break;
                                        case WebPartType.NewsReel:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.NewsReel);
                                            break;
                                        case WebPartType.PageTitle:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.PageTitle);
                                            break;
                                        case WebPartType.People:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.People);
                                            break;
                                        case WebPartType.QuickChart:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.QuickChart);
                                            break;
                                        case WebPartType.QuickLinks:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.QuickLinks);
                                            break;
                                        case WebPartType.SiteActivity:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SiteActivity);
                                            break;
                                        case WebPartType.VideoEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.VideoEmbed);
                                            break;
                                        case WebPartType.YammerEmbed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.YammerEmbed);
                                            break;
                                        case WebPartType.CustomMessageRegion:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.CustomMessageRegion);
                                            break;
                                        case WebPartType.Divider:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Divider);
                                            break;
                                        case WebPartType.Spacer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Spacer);
                                            break;
                                        case WebPartType.Kindle:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Kindle);
                                            break;
                                        case WebPartType.MyFeed:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.MyFeed);
                                            break;
                                        case WebPartType.OrgChart:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.OrgChart);
                                            break;
                                        case WebPartType.SavedForLater:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SavedForLater);
                                            break;
                                        case WebPartType.Twitter:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.Twitter);
                                            break;
                                        case WebPartType.WorldClock:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.WorldClock);
                                            break;
                                        case WebPartType.SpacesDocLib:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesDocLib);
                                            break;
                                        case WebPartType.SpacesFileViewer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesFileViewer);
                                            break;
                                        case WebPartType.SpacesImageViewer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesImageViewer);
                                            break;
                                        case WebPartType.SpacesModelViewer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesModelViewer);
                                            break;
                                        case WebPartType.SpacesImageThreeSixty:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesImageThreeSixty);
                                            break;
                                        case WebPartType.SpacesVideoThreeSixty:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesVideoThreeSixty);
                                            break;
                                        case WebPartType.SpacesText2D:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesText2D);
                                            break;
                                        case WebPartType.SpacesVideoPlayer:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesVideoPlayer);
                                            break;
                                        case WebPartType.SpacesPeople:
                                            webPartName = page.DefaultWebPartToWebPartId(PnPCore.DefaultWebPart.SpacesPeople);
                                            break;
                                    }

                                    baseControl = componentsToAdd.FirstOrDefault(p => p.Name.Equals(webPartName, StringComparison.InvariantCultureIgnoreCase));
                                }

                                if (baseControl != null)
                                {
                                    PnPCore.IPageWebPart myWebPart = page.NewWebPart(baseControl);
                                    myWebPart.Order = control.Order;
                                    //{
                                    //    Order = control.Order
                                    //};

                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        var json = JsonConvert.DeserializeObject<JObject>(control.JsonControlData);
                                        if (json["instanceId"] != null && json["instanceId"].Type != JTokenType.Null)
                                        {
                                            if (Guid.TryParse(json["instanceId"].Value<string>(), out Guid instanceId))
                                            {
                                                myWebPart.InstanceId = instanceId;
                                            }
                                        }
                                    }

                                    // Reduce column number by 1 due 0 start indexing
                                    page.AddControl(myWebPart, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                                    // set properties using json string
                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        myWebPart.PropertiesJson = control.JsonControlData;
                                    }

                                    //CHECK:
                                    // set using property collection
                                    //if (control.ControlProperties.Any())
                                    //{
                                    //    // grab the "default" properties so we can deduct their types, needed to correctly apply the set properties
                                    //    var controlManifest = JObject.Parse(baseControl.Manifest);
                                    //    JToken controlProperties = null;
                                    //    if (controlManifest != null)
                                    //    {
                                    //        controlProperties = controlManifest.SelectToken("preconfiguredEntries[0].properties");
                                    //    }

                                    //    foreach (var property in control.ControlProperties)
                                    //    {
                                    //        Type propertyType = typeof(string);

                                    //        if (controlProperties != null)
                                    //        {
                                    //            var defaultProperty = controlProperties.SelectToken(property.Key, false);
                                    //            if (defaultProperty != null)
                                    //            {
                                    //                propertyType = Type.GetType($"System.{defaultProperty.Type}");

                                    //                if (propertyType == null)
                                    //                {
                                    //                    if (defaultProperty.Type.ToString().Equals("integer", StringComparison.InvariantCultureIgnoreCase))
                                    //                    {
                                    //                        propertyType = typeof(int);
                                    //                    }
                                    //                }
                                    //            }
                                    //        }

                                    //        myWebPart.Properties[property.Key] = JToken.FromObject(Convert.ChangeType(parser.ParseString(property.Value), propertyType));
                                    //    }
                                    //}
                                }
                                else
                                {

                                    PnPCore.IPageWebPart myWebPart = page.NewWebPart();
                                    myWebPart.Order = control.Order;
                                    

                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        var json = JsonConvert.DeserializeObject<JObject>(control.JsonControlData);
                                        if (json["instanceId"] != null && json["instanceId"].Type != JTokenType.Null)
                                        {
                                            if (Guid.TryParse(json["instanceId"].Value<string>(), out Guid instanceId))
                                            {
                                                myWebPart.InstanceId = instanceId;
                                            }
                                        }
                                        if (json["id"] != null && json["id"].Type != JTokenType.Null)
                                        {
                                            if (Guid.TryParse(json["id"].Value<string>(), out Guid webPartId))
                                            {
                                                var pageWebPartType = typeof(PnPCore.IPageWebPart).Assembly.GetType("PnP.Core.Model.SharePoint.PageWebPart");

                                                PropertyInfo propertyInfo = pageWebPartType.GetProperty("WebPartId");
                                                if (propertyInfo != null)
                                                {
                                                    propertyInfo.SetValue(myWebPart, json["id"].Value<string>());
                                                }
                                            }
                                        }
                                    }

                                    // Reduce column number by 1 due 0 start indexing
                                    page.AddControl(myWebPart, page.Sections[sectionCount].Columns[control.Column - 1], control.Order);

                                    // set properties using json string
                                    if (!string.IsNullOrEmpty(control.JsonControlData))
                                    {
                                        myWebPart.PropertiesJson = control.JsonControlData;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Handle the header controls in the topic pages
            if (page.LayoutType == PnPCore.PageLayoutType.Topic)
            {
                var headerControlSection = clientSidePage.Sections.FirstOrDefault(p => p.Order == 999999);
                if (headerControlSection != null)
                {
                    // Ensure there's at least one default section available
                    if (!page.Sections.Any())
                    {
                        page.AddSection(PnPCore.CanvasSectionTemplate.OneColumn, 0);
                    }

                    // Clear existing header controls as they'll be overwritten
                    page.HeaderControls.Clear();                    

                    // Load existing available controls
                    var componentsToAdd = page.AvailablePageComponents();

                    int order = 1;
                    foreach (var headerControl in headerControlSection.Controls)
                    {
                        PnPCore.IPageComponent baseControl = null;

                        // apply token parsing on the web part properties
                        headerControl.JsonControlData = parser.ParseString(headerControl.JsonControlData);

                        if (headerControl.Type == WebPartType.Custom)
                        {
                            // Find the base control installed to the current site
                            baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals($"{{{headerControl.ControlId}}}", StringComparison.CurrentCultureIgnoreCase));
                            if (baseControl == null)
                            {
                                baseControl = componentsToAdd.FirstOrDefault(p => p.Id.Equals(headerControl.ControlId.ToString(), StringComparison.InvariantCultureIgnoreCase));
                            }

                            if (baseControl != null)
                            {
                                PnPCore.IPageWebPart myWebPart = page.NewWebPart(baseControl);

                                myWebPart.IsHeaderControl = true;

                                if (!string.IsNullOrEmpty(headerControl.JsonControlData))
                                {
                                    var json = JsonConvert.DeserializeObject<JObject>(headerControl.JsonControlData);
                                    if (json["instanceId"] != null && json["instanceId"].Type != JTokenType.Null)
                                    {
                                        if (Guid.TryParse(json["instanceId"].Value<string>(), out Guid instanceId))
                                        {
                                            myWebPart.InstanceId = instanceId;
                                        }
                                    }

                                    if (json["dataVersion"] != null && json["dataVersion"].Type != JTokenType.Null)
                                    {
                                        myWebPart.DataVersion = json["dataVersion"].Value<string>();
                                    }
                                }

                                // set properties using json string
                                if (!string.IsNullOrEmpty(headerControl.JsonControlData))
                                {
                                    myWebPart.PropertiesJson = headerControl.JsonControlData;
                                }

                                page.AddHeaderControl(myWebPart, order);
                                order++;
                            }
                            else
                            {
                                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePages_BaseControlNotFound, headerControl.ControlId, headerControl.CustomWebPartName);
                            }

                        }
                    }
                }
            }

            // Persist the page
            if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
            {
                page.SaveAsTemplate(pageName.Replace($"{dummyPage.GetTemplatesFolder()}/", ""));
            }
            else
            {
                page.Save(pageName);
            }

            // Load the page list item
            var fileAfterSave = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
            web.Context.Load(fileAfterSave, p=>p.ListItemAllFields);
            web.Context.ExecuteQueryRetry();

            // Update page content type
            bool isDirty = false;
            if (!string.IsNullOrEmpty(clientSidePage.ContentTypeID))
            {
                ContentTypeId bestMatchCT = fileAfterSave.ListItemAllFields.ParentList.BestMatchContentTypeId(clientSidePage.ContentTypeID);
                ContentTypeId currentCT = fileAfterSave.ListItemAllFields.FieldExistsAndUsed(ContentTypeIdField) ? ((ContentTypeId)fileAfterSave.ListItemAllFields[ContentTypeIdField]) : null;

                if (currentCT == null)
                {
                    fileAfterSave.ListItemAllFields[ContentTypeIdField] = bestMatchCT.StringValue;
                    isDirty = true;
                }
                else if (currentCT != null && !currentCT.IsChildOf(bestMatchCT))
                {
                    fileAfterSave.ListItemAllFields[ContentTypeIdField] = bestMatchCT.StringValue;
                    isDirty = true;
                }
            }

            if (clientSidePage.PromoteAsTemplate && page.LayoutType == PnPCore.PageLayoutType.Article)
            {
                // Choice field, currently there's only one value possible and that's Template
                fileAfterSave.ListItemAllFields[SPSitePageFlagsField] = ";#Template;#";
                isDirty = true;
            }

            if (isDirty)
            {
                if (exists)
                {
                    fileAfterSave.ListItemAllFields.SystemUpdate();
                }
                else
                {
                    fileAfterSave.ListItemAllFields.UpdateOverwriteVersion();
                }
                web.Context.Load(fileAfterSave.ListItemAllFields);
                web.Context.ExecuteQueryRetry();
            }

            if (clientSidePage.FieldValues != null && clientSidePage.FieldValues.Any())
            {
                ListItemUtilities.UpdateListItem(fileAfterSave.ListItemAllFields, parser, clientSidePage.FieldValues, ListItemUtilities.ListItemUpdateType.UpdateOverwriteVersion);
            }

            // Set page property bag values
            if (clientSidePage.Properties != null && clientSidePage.Properties.Any())
            {
                string pageFilePath = fileAfterSave.ListItemAllFields[FileRefField].ToString();
                var pageFile = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pageFilePath));
                web.Context.Load(pageFile, p => p.Properties);

                foreach (var pageProperty in clientSidePage.Properties)
                {
                    if (!string.IsNullOrEmpty(pageProperty.Key))
                    {
                        pageFile.Properties[pageProperty.Key] = pageProperty.Value;
                    }
                }

                pageFile.Update();
                web.Context.Load(fileAfterSave.ListItemAllFields);
                web.Context.ExecuteQueryRetry();
            }

            if (page.LayoutType != PnPCore.PageLayoutType.SingleWebPartAppPage)
            {
                // Set commenting, ignore on pages of the type Home or page templates
                if (page.LayoutType != PnPCore.PageLayoutType.Home && !clientSidePage.PromoteAsTemplate)
                {
                    // Make it a news page if requested
                    if (clientSidePage.PromoteAsNewsArticle)
                    {
                        page.PromoteAsNewsArticle();
                    }
                }

                if (page.LayoutType != PnPCore.PageLayoutType.RepostPage)
                {
                    if (clientSidePage.EnableComments)
                    {
                        page.EnableComments();
                    }
                    else
                    {
                        page.DisableComments();
                    }
                }
            }

            // Publish page, page templates cannot be published
            if (clientSidePage.Publish && !clientSidePage.PromoteAsTemplate)
            {
                page.Publish();
            }

            // Set any security on the page
            if (clientSidePage.Security != null && clientSidePage.Security.RoleAssignments.Count != 0)
            {
                web.Context.Load(fileAfterSave.ListItemAllFields);
                web.Context.ExecuteQueryRetry();
                fileAfterSave.ListItemAllFields.SetSecurity(parser, clientSidePage.Security, WriteMessage);
            }
        }

        private static string DeterminePageName(TokenParser parser, BaseClientSidePage clientSidePage)
        {
            string pageName;
            if (clientSidePage is ClientSidePage)
            {
                if (clientSidePage.PromoteAsTemplate)
                {
                    pageName = $"{System.IO.Path.GetFileNameWithoutExtension(parser.ParseString((clientSidePage as ClientSidePage).PageName))}.aspx";
                }
                else
                {
                    var parsedPageName = parser.ParseString((clientSidePage as ClientSidePage).PageName);
                    var pageNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(parsedPageName);
                    var pageFolder = System.IO.Path.GetDirectoryName(parsedPageName);
                    if (!string.IsNullOrEmpty(pageFolder))
                    {
                        pageFolder += "/";
                    }

                    pageName = $"{pageFolder}{pageNameWithoutExtension}.aspx";
                }
            }
            else
            {
                pageName = parser.ParseString((clientSidePage as TranslatedClientSidePage).PageName);
            }

            return pageName;
        }

        private string PreCreatePage(Web web, ProvisioningTemplate template, TokenParser parser, BaseClientSidePage clientSidePage, string pagesLibrary, List pagesLibraryList, ref int currentPageIndex)
        {
            string pageName = DeterminePageName(parser, clientSidePage);
            string url = $"{pagesLibrary}/{pageName}";

            if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
            {
                url = $"{pagesLibrary}/{dummyPage.GetTemplatesFolder()}/{pageName}";
            }

            // Write page level status messages, needed in case many pages are provisioned
            currentPageIndex++;
            WriteSubProgress("ClientSidePage", $"Create {pageName} stub", currentPageIndex, template.ClientSidePages.Count);

            url = UrlUtility.Combine(web.ServerRelativeUrl, url);

            var exists = true;
            try
            {
                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath, f => f.Exists);
                web.Context.ExecuteQueryRetry();

                // Fill token
                parser.AddToken(new PageUniqueIdToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));

                exists = file.Exists;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    exists = false;
                }
            }

            if (!exists)
            {
                // Pre-create the page    
                PnPCore.IPage page = web.AddClientSidePage(pageName);

                // Set page layout now, because once it's set, it can't be changed.
                if (!string.IsNullOrEmpty(clientSidePage.Layout))
                {
                    page.LayoutType = (PnPCore.PageLayoutType)Enum.Parse(typeof(PnPCore.PageLayoutType), clientSidePage.Layout);
                }

                string createdPageName;
                if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                {
                    createdPageName = page.SaveAsTemplate(pageName);
                }
                else
                {
                    createdPageName = page.Save(pageName);
                }

                url = $"{pagesLibrary}/{createdPageName}";
                if (clientSidePage.Layout == "Article" && clientSidePage.PromoteAsTemplate)
                {
                    url = $"{pagesLibrary}/{dummyPage.GetTemplatesFolder()}/{pageName}";
                }
                url = UrlUtility.Combine(web.ServerRelativeUrl, url);

                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(url));
                web.Context.Load(file, f => f.UniqueId, f => f.ServerRelativePath);
                web.Context.ExecuteQueryRetry();

                // Fill token
                parser.AddToken(new PageUniqueIdToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));
                parser.AddToken(new PageUniqueIdEncodedToken(web, file.ServerRelativePath.DecodedUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), file.UniqueId));

                // Track that we pre-added this page
                return url;
            }

            return null;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (new PnPMonitoredScope(this.Name))
            {
                // Impossible to return all files in the site currently

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.ClientSidePages.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }
}
