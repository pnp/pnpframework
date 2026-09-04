using System;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageIngredientIds
    {
        public const string Runtime = "runtime:page";

        public const string PageArtifact = "artifact:page";

        public const string Layout = "layout:page";

        public const string ContentType = "content-type:page";

        public const string PublishingContent = "content:publishing-page-content";

        public const string Security = "security:page";

        public const string Lifecycle = "lifecycle:page";

        public static string Field(string internalName)
        {
            return "field:" + (internalName ?? string.Empty);
        }

        public static string TaxonomyRelationship(Guid sourceFieldId, Guid sourceTermId, int sourceWssId)
        {
            return "taxonomy-relationship:" + sourceFieldId.ToString("D") + "/" + sourceTermId.ToString("D") + "/" + sourceWssId;
        }

        public static string WebPart(Guid id)
        {
            return "webpart:" + id.ToString("D");
        }

        public static string List(Guid sourceWebId, Guid sourceListId)
        {
            return "list:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D");
        }

        public static string PlatformFeature(Guid sourceSiteId, Guid featureId)
        {
            return "platform-feature:" + sourceSiteId.ToString("D") + "/" + featureId.ToString("D");
        }

        public static string View(Guid sourceWebId, Guid sourceListId, Guid sourceViewId)
        {
            return "view:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D") + "/" + sourceViewId.ToString("D");
        }

        public static string ViewRenderingResource(Guid sourceSiteId, string resourceId)
        {
            return "view-rendering-resource:" + sourceSiteId.ToString("D") + "/" + (resourceId ?? string.Empty);
        }

        public static string Web(Guid sourceSiteId, Guid sourceWebId)
        {
            return "web:" + sourceSiteId.ToString("D") + "/" + sourceWebId.ToString("D");
        }

        public static string LayoutResource(string sourceReference)
        {
            return "layout-resource:" + (sourceReference ?? string.Empty);
        }

        public static string PageContentTypeField(Guid sourceFieldId)
        {
            return "page-content-type-field:" + sourceFieldId.ToString("D");
        }

        public static string SiteContentType(string sourceScope, string sourceContentTypeId)
        {
            return "site-content-type:" + NormalizeScope(sourceScope) + "/" + (sourceContentTypeId ?? string.Empty);
        }

        public static string SiteField(string sourceScope, Guid sourceFieldId)
        {
            return "site-field:" + NormalizeScope(sourceScope) + "/" + sourceFieldId.ToString("D");
        }

        public static string ListField(Guid sourceWebId, Guid sourceListId, Guid sourceFieldId)
        {
            return "list-field:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D") + "/" + sourceFieldId.ToString("D");
        }

        public static string ListContentType(Guid sourceWebId, Guid sourceListId, string sourceContentTypeId)
        {
            return "list-content-type:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D") + "/" + (sourceContentTypeId ?? string.Empty);
        }

        public static string ListItem(Guid sourceWebId, Guid sourceListId, int sourceItemId)
        {
            return "list-item:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D") + "/" + sourceItemId;
        }

        public static string ListDocument(Guid sourceWebId, Guid sourceListId, int sourceItemId)
        {
            return "list-document:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D") + "/" + sourceItemId;
        }

        public static string ListDocumentInformationProtection(Guid sourceWebId, Guid sourceListId, int sourceItemId)
        {
            return "list-document-information-protection:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D") + "/" + sourceItemId;
        }

        public static string ListAttachment(Guid sourceWebId, Guid sourceListId, int sourceItemId, string fileName)
        {
            return "list-attachment:" + sourceWebId.ToString("D") + "/" + sourceListId.ToString("D") + "/" + sourceItemId + "/" + (fileName ?? string.Empty);
        }

        public static string Reference(string id)
        {
            return "reference:" + (id ?? string.Empty);
        }

        private static string NormalizeScope(string value)
        {
            var normalized = (value ?? string.Empty).Replace('\\', '/').TrimEnd('/');
            return normalized.Length == 0 ? "/" : normalized;
        }
    }
}
