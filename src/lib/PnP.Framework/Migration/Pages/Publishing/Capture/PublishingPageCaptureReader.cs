using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Security;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Markup;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace PnP.Framework.Migration.Pages.Publishing.Capture
{
    internal static class PublishingPageCaptureReader
    {
        public static CapturedPublishingPage Read(
            ClientContext context,
            string pagePath,
            PageCaptureOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            return Read(context, pagePath, options, null, blockers, warnings);
        }

        public static CapturedPublishingPage Read(
            ClientContext context,
            string pagePath,
            PageCaptureOptions options,
            IMigrationArtifactStore artifactStore,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var web = context.Web;
            var site = context.Site;
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            var item = file.ListItemAllFields;
            var contentType = item.ContentType;
            context.Load(site, value => value.Id);
            context.Load(web, value => value.Id, value => value.Url, value => value.ServerRelativeUrl);
            context.Load(file,
                value => value.Exists,
                value => value.Name,
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

            var layoutSnapshot = PublishingPageLayoutSnapshotReader.Read(
                context,
                layout?.Url,
                layout?.Description,
                artifactStore,
                blockers,
                warnings);
            var pageArtifact = PageArtifactSnapshotReader.Read(context, file, artifactStore, blockers);

            var identity = new PageIdentity
            {
                SiteId = site.Id,
                WebId = web.Id,
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
                Title = GetFieldString(item, "Title") ?? PagePath.GetFileName(pagePath)
            };

            return new CapturedPublishingPage
            {
                Identity = identity,
                PageArtifact = pageArtifact,
                Layout = layoutSnapshot,
                PublishingPageContent = content,
                Fields = PageFieldSnapshotReader.Read(context, item, warnings),
                WebParts = options.IncludeWebParts
                    ? ClassicWebPartSnapshotReader.Read(web, pagePath, blockers)
                    : new List<ClassicWebPartSnapshot>(),
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
