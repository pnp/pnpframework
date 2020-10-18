using Microsoft.SharePoint.Client;
using PnP.Framework.Pages;
using PnP.Framework.Utilities;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Extensions;
using PnP.Framework.Modernization.Pages;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Transform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace PnP.Framework.Modernization.Delve
{
    /// <summary>
    /// Transformator that convers a Delve blog page into a modern client side page
    /// </summary>
    public class DelvePageTransformator: BasePageTransformator
    {
        private readonly Regex invalidCharsRegex = new Regex(@"[\*\?\|\\\t/:""'<>#{}%~&]", RegexOptions.Compiled);
        private readonly Regex invalidRulesRegex = new Regex(@"\.{2,}", RegexOptions.Compiled);
        private readonly Regex startEndRegex = new Regex(@"^[\. ]|[\. ]$", RegexOptions.Compiled);
        private readonly Regex extraSpacesRegex = new Regex(" {2,}", RegexOptions.Compiled);
        private string pagesLibraryName = "pPg";

        #region Construction
        /// <summary>
        /// Creates a page transformator instance with a target destination of a target web e.g. Modern/Communication Site
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="targetClientContext">ClientContext of the site that will receive the modernized page</param>
        public DelvePageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext) : this(sourceClientContext, targetClientContext, "webpartmapping.xml")
        {

        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        public DelvePageTransformator(ClientContext sourceClientContext) : this(sourceClientContext, null, "webpartmapping.xml")
        {
        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="pageTransformationFile">Used page mapping file</param>
        public DelvePageTransformator(ClientContext sourceClientContext, string pageTransformationFile) : this(sourceClientContext, null, pageTransformationFile)
        {

        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="targetClientContext">ClientContext of the site that will receive the modernized page</param>
        /// <param name="pageTransformationFile">Used page mapping file</param>
        public DelvePageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, string pageTransformationFile)
        {
            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            using (Stream schema = typeof(PageTransformator).Assembly.GetManifestResourceStream("PnP.Framework.Modernization.webpartmapping.xsd"))
            {
                // Load xml mapping data
                XmlSerializer xmlMapping = new XmlSerializer(typeof(PageTransformation));
                using (var stream = new FileStream(pageTransformationFile, FileMode.Open))
                {
                    // Ensure the provided file complies with the current schema
                    ValidateSchema(schema, stream);

                    // All good so it seems
                    this.pageTransformation = (PageTransformation)xmlMapping.Deserialize(stream);
                }
            }
        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="pageTransformationModel">Page transformation model</param>
        public DelvePageTransformator(ClientContext sourceClientContext, PageTransformation pageTransformationModel) : this(sourceClientContext, null, pageTransformationModel)
        {

        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="targetClientContext">ClientContext of the site that will receive the modernized page</param>
        /// <param name="pageTransformationModel">Page transformation model</param>
        public DelvePageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, PageTransformation pageTransformationModel)
        {

            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            this.pageTransformation = pageTransformationModel;
        }
        #endregion

        /// <summary>
        /// Transform the page
        /// </summary>
        /// <param name="pageTransformationInformation">Information about the page to transform</param>
        /// <returns>The path to the created modern page</returns>
        public string Transform(DelvePageTransformationInformation pageTransformationInformation)
        {
            SetPageId(Guid.NewGuid().ToString());

            var logsForSettings = this.DetailSettingsAsLogEntries(pageTransformationInformation);
            logsForSettings?.ForEach(o => Log(o, LogLevel.Information));

            #region Check for Target Site Context
            var hasTargetContext = this.targetClientContext != null;
            #endregion

            #region Input validation
            string pageType = null;

            if (pageTransformationInformation.SourcePage == null)
            {
                LogError(LogStrings.Error_SourcePageNotFound, LogStrings.Heading_InputValidation);
                throw new ArgumentNullException(LogStrings.Error_SourcePageNotFound);
            }

            // Validate page and it's eligibility for transformation
            if (!pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileRefField) || !pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileLeafRefField))
            {
                LogError(LogStrings.Error_PageNotValidMissingFileRef, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_PageNotValidMissingFileRef);
            }

            pageType = pageTransformationInformation.SourcePage.PageType();
            LogInfo(string.Format(LogStrings.TransformationMode, pageType.FormatAsFriendlyTitle()), LogStrings.Heading_Summary, LogEntrySignificance.TransformMode);

            if (pageType.Equals("AspxPage", StringComparison.InvariantCultureIgnoreCase))
            {
                LogError(LogStrings.Error_BasicASPXPageCannotTransform, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_BasicASPXPageCannotTransform);
            }

            if (pageType.Equals("WikiPage", StringComparison.InvariantCultureIgnoreCase) || pageType.Equals("WebPartPage", StringComparison.InvariantCultureIgnoreCase))
            {
                LogError(LogStrings.Error_PageIsNotAPublishingPage, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_PageIsNotAPublishingPage);
            }

            if (IsClientSidePage(pageType))
            {
                LogError(LogStrings.Error_SourcePageIsModern, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_SourcePageIsModern);
            }

            if (IsAspxPage(pageType))
            {
                LogError(LogStrings.Error_BasicASPXPageCannotTransform, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_BasicASPXPageCannotTransform);
            }

            if (IsPublishingPage(pageType))
            {
                LogError(LogStrings.Error_PublishingPagesNotYetSupported, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_PublishingPagesNotYetSupported);
            }
            #endregion

            if (!hasTargetContext)
            {
                LogError(LogStrings.Error_BlogPageTransformationHasToBeCrossSite, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_BlogPageTransformationHasToBeCrossSite);
            }

            // Disable cross-farm item level permissions from copying
            CrossFarmTransformationValidation(pageTransformationInformation);

            LogDebug(LogStrings.ValidationChecksComplete, LogStrings.Heading_InputValidation);

            try
            {
                #region Telemetry
                DateTime transformationStartDateTime = DateTime.Now;

                LogDebug(LogStrings.LoadingClientContextObjects, LogStrings.Heading_SharePointConnection);
                LoadClientObject(sourceClientContext, false);

                LogInfo($"{sourceClientContext.Web.GetUrl()}", LogStrings.Heading_Summary, LogEntrySignificance.SourceSiteUrl);

                LogDebug(LogStrings.LoadingTargetClientContext, LogStrings.Heading_SharePointConnection);
                LoadClientObject(targetClientContext, true);

                PopulateGlobalProperties(sourceClientContext, targetClientContext);

                if (sourceClientContext.Site.Id.Equals(targetClientContext.Site.Id))
                {
                    // Oops, seems source and target point to the same site collection...that's a no go for publishing portal page transformation!                
                    LogError(LogStrings.Error_SameSiteTransferNoAllowedForPublishingPages, LogStrings.Heading_SharePointConnection);
                    throw new ArgumentNullException(LogStrings.Error_SameSiteTransferNoAllowedForPublishingPages);
                }

                LogInfo($"{targetClientContext.Web.GetUrl()}", LogStrings.Heading_Summary, LogEntrySignificance.TargetSiteUrl);

                // Need to add further validation for target template
                if (targetClientContext.Web.WebTemplate != "SITEPAGEPUBLISHING" && targetClientContext.Web.WebTemplate != "STS" && targetClientContext.Web.WebTemplate != "GROUP")
                {

                    LogError(LogStrings.Error_CrossSiteTransferTargetsNonModernSite);
                    throw new ArgumentException(LogStrings.Error_CrossSiteTransferTargetsNonModernSite, LogStrings.Heading_SharePointConnection);
                }

                // Ensure PostAsNews is used together with PagePublishing
                if (pageTransformationInformation.PostAsNews && !pageTransformationInformation.PublishCreatedPage)
                {
                    pageTransformationInformation.PublishCreatedPage = true;
                    LogWarning(LogStrings.Warning_PostingAPageAsNewsRequiresPagePublishing, LogStrings.Heading_Summary);
                }

                // Store the information of the source page we do want to retain
                if (pageTransformationInformation.KeepPageCreationModificationInformation)
                {
                    StoreSourcePageInformationToKeep(pageTransformationInformation.SourcePage);
                }

                LogInfo($"{pageTransformationInformation.SourcePage[Constants.FileRefField].ToString()}", LogStrings.Heading_Summary, LogEntrySignificance.SourcePage);

                var spVersion = pageTransformationInformation.SourceVersion;
                var exactSpVersion = pageTransformationInformation.SourceVersionNumber;
                LogInfo($"{spVersion.DisplaySharePointVersion()} ({exactSpVersion})", LogStrings.Heading_Summary, LogEntrySignificance.SharePointVersion);
                
                //Load User Mapping File
                InitializeUserMapping(pageTransformationInformation);
                #endregion

                #region Page creation
                // Detect if the page is living inside a folder
                LogDebug(LogStrings.DetectIfPageIsInFolder, LogStrings.Heading_PageCreation);
                string pageFolder = "";

                if (pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileDirRefField))
                {
                    var fileRefFieldValue = pageTransformationInformation.SourcePage[Constants.FileDirRefField].ToString();

                    if (fileRefFieldValue.ContainsIgnoringCasing($"/{this.pagesLibraryName}"))
                    {
                        string pagesLibraryRelativeUrl = $"{sourceClientContext.Web.ServerRelativeUrl.TrimEnd(new[] { '/' })}/{this.pagesLibraryName}";
                        pageFolder = fileRefFieldValue.Replace(pagesLibraryRelativeUrl, "", StringComparison.InvariantCultureIgnoreCase).Trim();
                    }
                    else
                    {
                        // Page was living in another list, leave the list name as that will be the folder hosting the modern file in SitePages.
                        // This convention is used to avoid naming conflicts
                        pageFolder = fileRefFieldValue.Replace($"{sourceClientContext.Web.ServerRelativeUrl}", "").Trim();
                    }

                    if (pageFolder.Length > 0 || !string.IsNullOrEmpty(pageTransformationInformation.TargetPageFolder))
                    {
                        if (pageFolder.StartsWith("/"))
                        {
                            if (pageFolder == "/")
                            {
                                pageFolder = "";
                            }
                            else
                            {
                                pageFolder = pageFolder.Substring(1);
                            }
                        }

                        // Add a trailing slash
                        pageFolder = pageFolder + "/";

                        if (!string.IsNullOrEmpty(pageTransformationInformation.TargetPageFolder))
                        {
                            if (pageTransformationInformation.TargetPageFolderOverridesDefaultFolder)
                            {
                                pageFolder = pageTransformationInformation.TargetPageFolder;
                            }
                            else
                            {
                                pageFolder = Path.Combine(pageFolder, pageTransformationInformation.TargetPageFolder);
                            }

                            if (!pageFolder.EndsWith("/"))
                            {
                                // Add a trailing slash
                                pageFolder = pageFolder + "/";
                            }
                        }

                        LogInfo(LogStrings.PageIsLocatedInFolder, LogStrings.Heading_PageCreation);
                    }
                }
                pageTransformationInformation.Folder = pageFolder;

                // If no targetname specified then we'll come up with one
                if (string.IsNullOrEmpty(pageTransformationInformation.TargetPageName))
                {
                    var generatedBlogPageName = $"{GetFieldValue(pageTransformationInformation, Constants.TitleField).Replace(" ", "-")}-{GetFieldValue(pageTransformationInformation, Constants.IDField)}.aspx";

                    // Based on this blog - http://www.simplyaprogrammer.com/2008/05/importing-files-into-sharepoint.html
                    string sanitizedName = extraSpacesRegex.Replace(invalidRulesRegex.Replace(invalidCharsRegex.Replace(input: generatedBlogPageName, replacement: string.Empty).Trim(), "."), " ");

                    while (startEndRegex.IsMatch(sanitizedName))
                    {
                        sanitizedName = startEndRegex.Replace(sanitizedName, string.Empty);
                    }

                    pageTransformationInformation.TargetPageName = sanitizedName;
                }

                // Check if page name is free to use

                bool pageExists = false;
                ClientSidePage targetPage = null;
                List pagesLibrary = null;
                Microsoft.SharePoint.Client.File existingFile = null;

                //The determines of the target client context has been specified and use that to generate the target page
                var context = targetClientContext;

                try
                {
                    LogDebug(LogStrings.LoadingExistingPageIfExists, LogStrings.Heading_PageCreation);

                    // Just try to load the page in the fastest possible manner, we only want to see if the page exists or not
                    existingFile = Load(sourceClientContext, pageTransformationInformation, pageType, out pagesLibrary, targetClientContext);
                    pageExists = true;
                }
                catch (Exception ex)
                {
                    if (ex is ArgumentException)
                    {
                        //Non-critical error generated 
                        LogInfo(LogStrings.CheckPageExistsError, LogStrings.Heading_PageCreation);
                    }
                    else
                    {
                        //Something else occurred
                        LogError(LogStrings.CheckPageExistsError, LogStrings.Heading_PageCreation, ex);
                    }
                }

                if (pageExists)
                {
                    LogInfo(LogStrings.PageAlreadyExistsInTargetLocation, LogStrings.Heading_PageCreation);

                    if (!pageTransformationInformation.Overwrite)
                    {
                        var message = $"{LogStrings.PageNotOverwriteIfExists}  {pageTransformationInformation.TargetPageName}.";
                        LogError(message, LogStrings.Heading_PageCreation);
                        throw new ArgumentException(message);
                    }
                }

                // Create the client side page

                targetPage = context.Web.AddClientSidePage($"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
                LogInfo($"{LogStrings.ModernPageCreated} ", LogStrings.Heading_PageCreation);
                #endregion

                LogInfo(LogStrings.TransformSourcePageAsArticlePage, LogStrings.Heading_ArticlePageHandling);

                #region Analysis of the source page

                LogInfo($"{LogStrings.TransformSourcePageIsDelvePage} - {LogStrings.TransformSourcePageAnalysing}", LogStrings.Heading_ArticlePageHandling);

                // Analyze the source page
                Tuple<PageLayout, List<WebPartEntity>> pageData = new DelvePage(pageTransformationInformation.SourcePage, pageTransformation, base.RegisteredLogObservers).AnalyzeAndTransform(pageTransformationInformation, targetPage);
                #endregion

                #region Layout transformation
                // Use the default layout transformator
                ILayoutTransformator layoutTransformator = new LayoutTransformator(targetPage);

                // Do we have an override?
                if (pageTransformationInformation.LayoutTransformatorOverride != null)
                {
                    LogInfo(LogStrings.TransformLayoutTransformatorOverride, LogStrings.Heading_ArticlePageHandling);
                    layoutTransformator = pageTransformationInformation.LayoutTransformatorOverride(targetPage);
                }

                // Apply the layout to the page
                layoutTransformator.Transform(pageData);
                #endregion

                #region Content transformation
                LogDebug(LogStrings.PreparingContentTransformation, LogStrings.Heading_ArticlePageHandling);

                // Use the default content transformator
                IContentTransformator contentTransformator = new ContentTransformator(sourceClientContext, targetPage, pageTransformation, pageTransformationInformation as BaseTransformationInformation, base.RegisteredLogObservers);

                // Do we have an override?
                if (pageTransformationInformation.ContentTransformatorOverride != null)
                {
                    LogInfo(LogStrings.TransformUsingContentTransformerOverride, LogStrings.Heading_ArticlePageHandling);

                    contentTransformator = pageTransformationInformation.ContentTransformatorOverride(targetPage, pageTransformation);
                }

                LogInfo(LogStrings.TransformingContentStart, LogStrings.Heading_ArticlePageHandling);

                // Run the content transformator
                contentTransformator.Transform(pageData.Item2.Where(c => !c.IsClosed).ToList());

                LogInfo(LogStrings.TransformingContentEnd, LogStrings.Heading_ArticlePageHandling);
                #endregion

                #region Text/Section/Column cleanup
                // Drop "empty" text parts. Wiki pages tend to have a lot of text parts just containing div's and BR's...no point in keep those as they generate to much whitespace
                RemoveEmptyTextParts(targetPage);

                // Remove empty sections and columns to optimize screen real estate
                if (pageTransformationInformation.RemoveEmptySectionsAndColumns)
                {
                    RemoveEmptySectionsAndColumns(targetPage);
                }
                #endregion

                #region Header Configuration
                if (pageTransformationInformation.SetAuthorInPageHeader)
                {
                    SetAuthorInPageHeader(targetPage);
                }
                #endregion

                #region Page persisting
                // Persist the client side page
                var pageName = $"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}";
                targetPage.Save(pageName);
                LogInfo($"{LogStrings.TransformSavedPageInCrossSiteCollection}: {pageName}", LogStrings.Heading_ArticlePageHandling);
                #endregion

                #region Page publishing
                // Tag the file with a page modernization version stamp
                string serverRelativePathForModernPage = targetPage.PageListItem[Constants.FileRefField].ToString();
                bool pageListItemWasReloaded = false;
                try
                {
                    var targetPageFile = context.Web.GetFileByServerRelativeUrl(serverRelativePathForModernPage);
                    context.Load(targetPageFile, p => p.Properties);
                    targetPageFile.Properties["sharepointpnp_pagemodernization"] = this.version;
                    targetPageFile.Update();

                    if (!pageTransformationInformation.KeepPageCreationModificationInformation &&
                        !pageTransformationInformation.PostAsNews &&
                        pageTransformationInformation.PublishCreatedPage)
                    {
                        // Try to publish, if publish is not needed/possible (e.g. when no minor/major versioning set) then this will return an error that we'll be ignoring
                        targetPageFile.Publish(LogStrings.PublishMessage);
                    }

                    // Ensure we've the most recent page list item loaded, must be last statement before calling ExecuteQuery
                    context.Load(targetPage.PageListItem);
                    // Send both the property update and publish as a single operation to SharePoint
                    context.ExecuteQueryRetry();
                    pageListItemWasReloaded = true;
                }
                catch (Exception)
                {
                    // Eat exceptions as this is not critical for the generated page
                    LogWarning(LogStrings.Warning_NonCriticalErrorDuringVersionStampAndPublish, LogStrings.Heading_ArticlePageHandling);
                }

                // Update flags field to indicate this is a "migrated" page
                try
                {
                    // If for some reason the reload batched with the previous request did not finish then do it again
                    if (!pageListItemWasReloaded)
                    {
                        context.Load(targetPage.PageListItem);
                        context.ExecuteQueryRetry();
                    }

                    // Only perform the update when the field was not yet set
                    bool skipSettingMigratedFromServerRendered = false;
                    if (targetPage.PageListItem[Constants.SPSitePageFlagsField] != null)
                    {
                        skipSettingMigratedFromServerRendered = (targetPage.PageListItem[Constants.SPSitePageFlagsField] as string[]).Contains("MigratedFromServerRendered");
                    }

                    if (!skipSettingMigratedFromServerRendered)
                    {
                        targetPage.PageListItem[Constants.SPSitePageFlagsField] = ";#MigratedFromServerRendered;#";
                        //targetPage.PageListItem.Update();
                        targetPage.PageListItem.UpdateOverwriteVersion();
                        context.Load(targetPage.PageListItem);
                        context.ExecuteQueryRetry();
                    }
                }
                catch (Exception)
                {
                    // Eat any exception
                }

                // Disable page comments on the create page, if needed
                if (pageTransformationInformation.DisablePageComments)
                {
                    targetPage.DisableComments();
                    LogInfo(LogStrings.TransformDisablePageComments, LogStrings.Heading_ArticlePageHandling);
                }
                #endregion

                ListItem finalListItemToUpdate = targetPage.PageListItem;

                #region Restore page author/editor/created/modified
                if ((pageTransformationInformation.SourcePage != null && pageTransformationInformation.KeepPageCreationModificationInformation && this.SourcePageAuthor != null && this.SourcePageEditor != null) ||
                    pageTransformationInformation.PostAsNews)
                {
                    UpdateTargetPageWithSourcePageInformation(finalListItemToUpdate, pageTransformationInformation, finalListItemToUpdate[Constants.FileRefField].ToString(), hasTargetContext);
                }
                #endregion

                // NO page updates are allowed anymore past this point as otherwise the set page usage information and published/posted state will be impacted!

                #region Telemetry
                if (!pageTransformationInformation.SkipTelemetry && this.pageTelemetry != null)
                {
                    TimeSpan duration = DateTime.Now.Subtract(transformationStartDateTime);
                    this.pageTelemetry.LogTransformationDone(duration, pageType, pageTransformationInformation);
                    this.pageTelemetry.Flush();
                }

                LogInfo(LogStrings.TransformComplete, LogStrings.Heading_PageCreation);
                #endregion

                #region Closing
                CacheManager.Instance.SetLastUsedTransformator(this);
                LogInfo($"{finalListItemToUpdate[Constants.FileRefField].ToString()}", LogStrings.Heading_Summary, LogEntrySignificance.TargetPage);
                return Uri.EscapeUriString(finalListItemToUpdate[Constants.FileRefField].ToString());
                #endregion

            }
            catch (Exception ex)
            {
                LogError(LogStrings.CriticalError_ErrorOccurred, LogStrings.Heading_Summary, ex, isCriticalException: true);

                // Throw exception if there's no registered log observers
                if (base.RegisteredLogObservers.Count == 0)
                {
                    throw;
                }
            }

            return string.Empty;
        }

        #region Helper methods
        /// <summary>
        /// Use reflection to read the object properties and detail the values
        /// </summary>
        /// <param name="pti">PageTransformationInformation object</param>
        /// <returns></returns>
        private List<LogEntry> DetailSettingsAsLogEntries(DelvePageTransformationInformation pti)
        {
            List<LogEntry> logs = new List<LogEntry>();

            try
            {
                // Add version 
                logs.Add(new LogEntry()
                {
                    Heading = LogStrings.Heading_PageTransformationInfomation,
                    Message = $"Engine version {LogStrings.KeyValueSeperatorToken} {this.version ?? "Not Specified"}"
                });

                var properties = pti.GetType().GetProperties();
                foreach (var property in properties)
                {
                    if (property.PropertyType == typeof(String) ||
                        property.PropertyType == typeof(bool))
                    {
                        var propVal = property.GetValue(pti);

                        logs.Add(new LogEntry()
                        {
                            Heading = LogStrings.Heading_PageTransformationInfomation,
                            Message = $"{property.Name.FormatAsFriendlyTitle()} {LogStrings.KeyValueSeperatorToken} {propVal ?? "Not Specified"}"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                logs.Add(new LogEntry() { Message = "Failed to convert object properties for reporting", Exception = ex, Heading = LogStrings.Heading_PageTransformationInfomation });
            }

            return logs;

        }

        private Microsoft.SharePoint.Client.File Load(ClientContext sourceContext, DelvePageTransformationInformation pageTransformationInformation, string pageType, out List pagesLibrary, ClientContext targetContext = null)
        {
            sourceContext.Web.EnsureProperty(w => w.ServerRelativeUrl);
            var listServerRelativeUrl = UrlUtility.Combine(sourceContext.Web.ServerRelativeUrl, this.pagesLibraryName);
            pagesLibrary = sourceContext.Web.GetList(listServerRelativeUrl);

            sourceContext.Web.Context.Load(pagesLibrary, l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title,
                                                l => l.Hidden, l => l.EffectiveBasePermissions, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);

            var contextForFile = targetClientContext;
            var sitePagesServerRelativeUrl = UrlUtility.Combine(contextForFile.Web.ServerRelativeUrl, "sitepages");

            var file = contextForFile.Web.GetFileByServerRelativeUrl($"{sitePagesServerRelativeUrl}/{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
            contextForFile.Web.Context.Load(file, f => f.Exists, f => f.ListItemAllFields);
            contextForFile.ExecuteQueryRetry();

            if (pageTransformationInformation.KeepPageSpecificPermissions)
            {
                sourceContext.Load(pageTransformationInformation.SourcePage, p => p.HasUniqueRoleAssignments);
            }

            try
            {
                sourceContext.ExecuteQueryRetry();
            }
            catch (ServerException se)
            {
                if (se.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    pagesLibrary = null;
                }
                else
                {
                    throw;
                }
            }

            if (pagesLibrary == null)
            {
                LogError(LogStrings.Error_MissingSitePagesLibrary, LogStrings.Heading_Load);
                throw new ArgumentException(LogStrings.Error_MissingSitePagesLibrary);
            }

            if (!file.Exists)
            {
                LogInfo(LogStrings.TransformPageDoesNotExistInWeb, LogStrings.Heading_Load);
                throw new ArgumentException($"{pageTransformationInformation.TargetPageName} - {LogStrings.TransformPageDoesNotExistInWeb}");
            }

            return file;
        }

        private void ValidateSchema(Stream schema, FileStream stream)
        {
            // Load the template into an XDocument
            XDocument xml = XDocument.Load(stream);

            // Prepare the XML Schema Set
            XmlSchemaSet schemas = new XmlSchemaSet();
            schema.Seek(0, SeekOrigin.Begin);
            schemas.Add(Constants.PageTransformationSchema, new XmlTextReader(schema));

            // Set stream back to start
            stream.Seek(0, SeekOrigin.Begin);

            xml.Validate(schemas, (o, e) =>
            {
                LogError(string.Format(LogStrings.Error_WebPartMappingSchemaValidation, e.Message), LogStrings.Heading_PageTransformationInfomation, e.Exception);
                throw new Exception(string.Format(LogStrings.Error_MappingFileSchemaValidation, e.Message));
            });
        }
        #endregion

    }
}
