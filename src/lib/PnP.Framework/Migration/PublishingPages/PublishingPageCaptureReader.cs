using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace PnP.Framework.Migration.PublishingPages
{
    internal static class PublishingPageCaptureReader
    {
        public static CapturedPublishingPage Read(
            ClientContext context,
            string pagePath,
            PublishingPageExportOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var web = context.Web;
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            var item = file.ListItemAllFields;
            var contentType = item.ContentType;
            context.Load(web, value => value.Url, value => value.ServerRelativeUrl);
            context.Load(file,
                value => value.Exists,
                value => value.UniqueId,
                value => value.ServerRelativeUrl,
                value => value.UIVersionLabel,
                value => value.Length,
                value => value.TimeLastModified,
                value => value.CheckOutType,
                value => value.Level,
                value => value.TimeCreated);
            context.Load(item);
            context.Load(contentType, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            context.Load(item, value => value.Id, value => value.HasUniqueRoleAssignments);
            context.ExecuteQueryRetry();
            if (!file.Exists)
            {
                throw new FileNotFoundException("The source publishing page was not found.", pagePath);
            }

            var content = GetFieldString(item, "PublishingPageContent") ?? string.Empty;
            var layout = item.FieldValues.TryGetValue("PublishingPageLayout", out var layoutValue)
                ? layoutValue as FieldUrlValue
                : null;
            if (string.IsNullOrWhiteSpace(content))
            {
                warnings.Add("PublishingPageContent is empty.");
            }

            if (layout == null || string.IsNullOrWhiteSpace(layout.Url))
            {
                blockers.Add("PublishingPageLayout is unavailable on the source page.");
            }

            var identity = new PublishingPageIdentity
            {
                WebUrl = web.Url.TrimEnd('/'),
                WebServerRelativeUrl = web.ServerRelativeUrl,
                PageServerRelativeUrl = file.ServerRelativeUrl,
                ListItemId = item.Id,
                FileUniqueId = file.UniqueId,
                ContentTypeId = contentType.Id.StringValue,
                ContentTypeName = contentType.Name,
                VersionLabel = file.UIVersionLabel,
                Length = file.Length,
                ModifiedUtc = file.TimeLastModified.ToUniversalTime(),
                Title = GetFieldString(item, "Title") ?? PublishingPagePath.GetFileName(pagePath),
                PageLayoutUrl = layout?.Url ?? string.Empty,
                PageLayoutDescription = layout?.Description
            };

            return new CapturedPublishingPage
            {
                Identity = identity,
                PublishingPageContent = content,
                Fields = PageFieldSnapshotReader.Read(context, item, warnings),
                WebParts = options.IncludeWebParts
                    ? PageWebPartSnapshotReader.Read(web, pagePath, blockers)
                    : new List<PageWebPartSnapshot>(),
                Security = PageSecuritySnapshotReader.Read(context, item, warnings),
                Lifecycle = new PageLifecycleSnapshot
                {
                    CheckOutType = file.CheckOutType.ToString(),
                    Level = file.Level.ToString(),
                    ModerationStatus = TryGetInt32(item, "_ModerationStatus"),
                    CreatedUtc = file.TimeCreated.ToUniversalTime(),
                    ModifiedUtc = file.TimeLastModified.ToUniversalTime()
                },
                SourceFence = SourcePageFenceReader.FromFile(file)
            };
        }

        internal static string GetFieldString(ListItem item, string internalName)
        {
            return item.FieldValues.TryGetValue(internalName, out var value)
                ? Convert.ToString(value, CultureInfo.InvariantCulture)
                : null;
        }

        internal static int? TryGetInt32(ListItem item, string internalName)
        {
            if (!item.FieldValues.TryGetValue(internalName, out var value) || value == null)
            {
                return null;
            }

            return int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out var result)
                ? result
                : (int?)null;
        }
    }
}
