using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutSchemaMaterializer
    {
        public static ContentTypeMaterializationDisposition Ensure(
            ClientContext context,
            PublishingPageLayoutMaterializationPlan plan,
            PublishingPageLayoutTargetAdmission admission)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            if (plan.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock)
            {
                return ContentTypeMaterializationDisposition.ReuseOwned;
            }

            if (admission?.ContentTypeSchema == null || plan.ContentTypeSchema == null)
            {
                throw new InvalidOperationException("A custom Page Layout requires an admitted associated content type closure.");
            }

            return ContentTypeMaterializer.Ensure(
                context,
                context.Site.RootWeb,
                plan.ContentTypeSchema,
                admission.ContentTypeSchema);
        }
    }
}
