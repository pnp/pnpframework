using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutTargetInspector
    {
        public static PublishingPageLayoutTargetProbe Inspect(
            ClientContext context,
            PublishingPageLayoutMaterializationPlan plan)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            try
            {
                var rootWeb = context.Site.RootWeb;
                var pageWeb = context.Web;
                context.Load(rootWeb,
                    value => value.Url,
                    value => value.ServerRelativeUrl,
                    value => value.EffectiveBasePermissions);
                context.Load(pageWeb,
                    value => value.Url,
                    value => value.ServerRelativeUrl,
                    value => value.EffectiveBasePermissions);
                context.Load(rootWeb.AvailableContentTypes, values => values.Include(value => value.Id, value => value.Name));
                context.Load(rootWeb.AvailableFields, values => values.Include(value => value.Id, value => value.InternalName));
                context.ExecuteQueryRetry();

                var diagnostics = new List<string>();
                var contentTypeCandidates = rootWeb.AvailableContentTypes
                    .Where(value => string.Equals(value.Name, plan.AssociatedContentTypeName, StringComparison.OrdinalIgnoreCase))
                    .ToArray();
                var associatedContentType = contentTypeCandidates.FirstOrDefault(value =>
                        string.Equals(value.Id.StringValue, plan.AssociatedContentTypeId, StringComparison.OrdinalIgnoreCase))
                    ?? (contentTypeCandidates.Length == 1 ? contentTypeCandidates[0] : null);
                if (contentTypeCandidates.Length > 1 && associatedContentType == null)
                {
                    diagnostics.Add($"Target exposes {contentTypeCandidates.Length} content types named '{plan.AssociatedContentTypeName}', but none has the exact source ID.");
                }

                var availableFields = new HashSet<string>(
                    rootWeb.AvailableFields
                        .SelectMany(field => new[] { field.Id.ToString("D"), field.InternalName })
                        .Where(value => !string.IsNullOrWhiteSpace(value)),
                    StringComparer.OrdinalIgnoreCase);
                var missingFields = plan.RequiredFieldBindings
                    .Where(binding => !availableFields.Contains(binding.Trim().Trim('{', '}')))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var result = ProbeLayoutFile(
                    context,
                    rootWeb,
                    plan,
                    associatedContentType,
                    missingFields,
                    diagnostics);
                result.ContentTypeSchema = plan.ContentTypeSchema == null
                    ? null
                    : ContentTypeTargetInspector.Inspect(context, rootWeb, plan.ContentTypeSchema);
                result.Resources = plan.ResourceMaterializations
                    .Where(value => value.Disposition == PublishingPageLayoutResourceMaterializationDisposition.CreateOrReuseOwned)
                    .Where(value => !string.IsNullOrWhiteSpace(value.TargetServerRelativeUrl))
                    .GroupBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                    .Select(group => PublishingPageLayoutResourceTargetInspector.Inspect(context, pageWeb, rootWeb, group.Key))
                    .ToList();
                return result;
            }
            catch (ServerException exception)
            {
                return new PublishingPageLayoutTargetProbe
                {
                    TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                    Availability = EvidenceAvailability.Unavailable,
                    Diagnostics = new List<string>
                    {
                        $"Target Page Layout inspection failed: {exception.Message}"
                    }
                };
            }
        }

        private static PublishingPageLayoutTargetProbe ProbeLayoutFile(
            ClientContext context,
            Web rootWeb,
            PublishingPageLayoutMaterializationPlan plan,
            ContentType associatedContentType,
            IList<string> missingFields,
            IList<string> diagnostics)
        {
            try
            {
                var file = rootWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(plan.TargetServerRelativeUrl));
                var stream = file.OpenBinaryStream();
                context.Load(file, value => value.Exists);
                context.Load(file.ListItemAllFields);
                context.ExecuteQueryRetry();
                if (!file.Exists)
                {
                    return Missing(plan, associatedContentType, missingFields, diagnostics, rootWeb);
                }

                if (stream.Value == null)
                {
                    diagnostics.Add("Target Page Layout exists, but its binary stream was unavailable.");
                    return new PublishingPageLayoutTargetProbe
                    {
                        TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                        FileExists = true,
                        AssociatedContentTypeAvailable = associatedContentType != null,
                        ResolvedAssociatedContentTypeId = associatedContentType?.Id.StringValue,
                        MissingFieldBindings = missingFields,
                        CanAddAndCustomizePages = rootWeb.EffectiveBasePermissions.Has(PermissionKind.AddAndCustomizePages),
                        Availability = EvidenceAvailability.Partial,
                        Diagnostics = diagnostics
                    };
                }

                string digest;
                using (stream.Value)
                {
                    digest = MigrationDigest.ComputeSha256(stream.Value);
                }

                var association = ParseAssociation(FieldString(file.ListItemAllFields, "PublishingAssociatedContentType"));
                return new PublishingPageLayoutTargetProbe
                {
                    TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                    FileExists = true,
                    ExistingBytesSha256 = digest,
                    ExistingAssociatedContentTypeName = association?.Name,
                    ExistingAssociatedContentTypeId = association?.Id,
                    AssociatedContentTypeAvailable = associatedContentType != null,
                    ResolvedAssociatedContentTypeId = associatedContentType?.Id.StringValue,
                    MissingFieldBindings = missingFields,
                    CanAddAndCustomizePages = rootWeb.EffectiveBasePermissions.Has(PermissionKind.AddAndCustomizePages),
                    Availability = EvidenceAvailability.Captured,
                    Diagnostics = diagnostics
                };
            }
            catch (ServerException exception) when (IsFileNotFound(exception))
            {
                return Missing(plan, associatedContentType, missingFields, diagnostics, rootWeb);
            }
        }

        private static PublishingPageLayoutTargetProbe Missing(
            PublishingPageLayoutMaterializationPlan plan,
            ContentType associatedContentType,
            IList<string> missingFields,
            IList<string> diagnostics,
            Web rootWeb)
        {
            return new PublishingPageLayoutTargetProbe
            {
                TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                FileExists = false,
                AssociatedContentTypeAvailable = associatedContentType != null,
                ResolvedAssociatedContentTypeId = associatedContentType?.Id.StringValue,
                MissingFieldBindings = missingFields,
                CanAddAndCustomizePages = rootWeb.EffectiveBasePermissions.Has(PermissionKind.AddAndCustomizePages),
                Availability = EvidenceAvailability.Captured,
                Diagnostics = diagnostics
            };
        }

        private static string FieldString(ListItem item, string name)
        {
            object value;
            return item.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value) : null;
        }

        private static (string Name, string Id)? ParseAssociation(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var parts = value.Split(new[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);
            for (var index = 0; index + 1 < parts.Length; index++)
            {
                if (parts[index + 1].Trim().StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                {
                    return (parts[index].Trim(), parts[index + 1].Trim());
                }
            }

            var comma = value.LastIndexOf(", 0x", StringComparison.OrdinalIgnoreCase);
            return comma <= 0 ? ((string, string)?)null : (value.Substring(0, comma).Trim(), value.Substring(comma + 2).Trim());
        }

        private static bool IsFileNotFound(ServerException exception)
        {
            return string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }
    }
}
