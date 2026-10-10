using Microsoft.SharePoint.Client;
using PnP.Framework.Modernization.Transform;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Modernization.Publishing
{
    /// <summary>
    /// Information used to configure the publishing page transformation process
    /// </summary>
    public class PublishingPageTransformationInformation : BaseTransformationInformation
    {
        #region Construction
        /// <summary>
        /// For internal use only
        /// </summary>
        public PublishingPageTransformationInformation() 
        {
            // constructor used exclusively for mocking
        }

        /// <summary>
        /// Instantiates the page transformation class
        /// </summary>
        /// <param name="sourcePage">Page we want to transform</param>
        public PublishingPageTransformationInformation(ListItem sourcePage) : this(sourcePage, false)
        {
        }

        /// <summary>
        /// Instantiates the page transformation class
        /// </summary>
        /// <param name="sourcePage">Page we want to transform</param>
        /// <param name="overwrite">Do we overwrite the target page if it already exists</param>
        public PublishingPageTransformationInformation(ListItem sourcePage, bool overwrite)
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
            // Populate with OOB mapping properties
            MappingProperties = new Dictionary<string, string>(5)
            {
                { Constants.UseCommunityScriptEditorMappingProperty, "false" },
                { Constants.SummaryLinksToQuickLinksMappingProperty, "true" }
            };
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Given a collection of field names will return the first one which contains a value which is not null or empty.
        /// </summary>
        /// <param name="fieldNames">A array of Internal field names to check</param>
        /// <returns>An Internal field name</returns>
        public string GetFirstNonEmptyFieldName(string[] fieldNames)
        {
            // trim each entry before continuing
            // and ignore empty or null values
            fieldNames = fieldNames.Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToArray();

            string fieldNameToReturn = string.Empty;

            foreach(string fieldName in fieldNames)
            {
                if(IsFieldUsed(fieldName))
                {
                    // use this field
                    fieldNameToReturn = fieldName;
                    break;
                }
            }

            return fieldNameToReturn;
        }

        /// <summary>
        /// Checks if a field is in use
        /// </summary>
        /// <param name="fieldName">Internal field name</param>
        /// <returns>True or False</returns>
        public virtual bool IsFieldUsed(string fieldName)
        {
            return this.SourcePage.FieldExistsAndUsed(fieldName);
        }
        #endregion

        #region Page Properties
        #endregion
    }
}
