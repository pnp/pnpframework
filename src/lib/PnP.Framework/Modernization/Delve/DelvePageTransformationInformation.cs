using Microsoft.SharePoint.Client;
using PnP.Framework.Modernization.Transform;
using System.Collections.Generic;

namespace PnP.Framework.Modernization.Delve
{
    /// <summary>
    /// Information to initiate the transformation of a Delve blog page
    /// </summary>
    public class DelvePageTransformationInformation: BaseTransformationInformation
    {
        #region Construction
        /// <summary>
        /// For internal use only
        /// </summary>
        public DelvePageTransformationInformation()
        {
            // constructor used exclusively for mocking
        }

        /// <summary>
        /// Instantiates the page transformation class
        /// </summary>
        /// <param name="sourcePage">Page we want to transform</param>
        public DelvePageTransformationInformation(ListItem sourcePage) : this(sourcePage, false)
        {
        }

        /// <summary>
        /// Instantiates the page transformation class
        /// </summary>
        /// <param name="sourcePage">Page we want to transform</param>
        /// <param name="overwrite">Do we overwrite the target page if it already exists</param>
        public DelvePageTransformationInformation(ListItem sourcePage, bool overwrite)
        {
            SourcePage = sourcePage;
            Overwrite = overwrite;
            HandleWikiImagesAndVideos = true;
            AddTableListImageAsImageWebPart = true;
            KeepPageSpecificPermissions = true;
            SkipTelemetry = false;
            RemoveEmptySectionsAndColumns = true;
            PublishCreatedPage = true;
            KeepPageCreationModificationInformation = false;
            PostAsNews = false;
            DisablePageComments = false;
            SkipUserMapping = false;
            TargetPageFolderOverridesDefaultFolder = false;
            SetAuthorInPageHeader = false;
            KeepSubTitle = false;
            // Populate with OOB mapping properties
            MappingProperties = new Dictionary<string, string>(5)
            {
                { Constants.UseCommunityScriptEditorMappingProperty, "false" },
                { Constants.SummaryLinksToQuickLinksMappingProperty, "true" }
            };
        }
        #endregion

        /// <summary>
        /// Sets the page author in the page header similar to the original page author
        /// </summary>
        public bool SetAuthorInPageHeader { get; set; }

        /// <summary>
        /// Converts the sub title of a Delve blog page as modern page topic header
        /// </summary>
        public bool KeepSubTitle { get; set; }
    }
}
