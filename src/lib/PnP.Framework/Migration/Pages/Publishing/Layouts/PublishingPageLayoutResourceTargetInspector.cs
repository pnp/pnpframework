using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutResourceTargetInspector
    {
        public static PublishingPageLayoutResourceTargetProbe Inspect(
            ClientContext context,
            Web pageWeb,
            Web rootWeb,
            string targetServerRelativeUrl)
        {
            var owner = ResolveOwner(pageWeb, rootWeb, targetServerRelativeUrl);
            var canWrite = owner.EffectiveBasePermissions.Has(PermissionKind.AddListItems)
                && owner.EffectiveBasePermissions.Has(PermissionKind.AddAndCustomizePages);
            try
            {
                var file = owner.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(targetServerRelativeUrl));
                var stream = file.OpenBinaryStream();
                context.Load(file, value => value.Exists);
                context.ExecuteQueryRetry();
                if (!file.Exists || stream.Value == null)
                {
                    return Missing(targetServerRelativeUrl, canWrite);
                }

                using (stream.Value)
                {
                    return new PublishingPageLayoutResourceTargetProbe
                    {
                        TargetServerRelativeUrl = targetServerRelativeUrl,
                        FileExists = true,
                        ExistingBytesSha256 = MigrationDigest.ComputeSha256(stream.Value),
                        CanWrite = canWrite,
                        Availability = EvidenceAvailability.Captured
                    };
                }
            }
            catch (ServerException exception) when (IsFileNotFound(exception))
            {
                return Missing(targetServerRelativeUrl, canWrite);
            }
            catch (ServerException exception)
            {
                return new PublishingPageLayoutResourceTargetProbe
                {
                    TargetServerRelativeUrl = targetServerRelativeUrl,
                    CanWrite = canWrite,
                    Availability = EvidenceAvailability.Unavailable,
                    Diagnostics = new List<string>
                    {
                        $"Target layout resource inspection failed: {exception.Message}"
                    }
                };
            }
        }

        private static Web ResolveOwner(Web pageWeb, Web rootWeb, string path)
        {
            var pagePath = pageWeb.ServerRelativeUrl.TrimEnd('/');
            var rootPath = rootWeb.ServerRelativeUrl == "/" ? string.Empty : rootWeb.ServerRelativeUrl.TrimEnd('/');
            if (!string.IsNullOrEmpty(pagePath)
                && path.StartsWith(pagePath + "/", StringComparison.OrdinalIgnoreCase))
            {
                return pageWeb;
            }

            if (string.IsNullOrEmpty(rootPath) || path.StartsWith(rootPath + "/", StringComparison.OrdinalIgnoreCase))
            {
                return rootWeb;
            }

            throw new InvalidOperationException($"Target layout resource escapes the target site collection: {path}");
        }

        private static PublishingPageLayoutResourceTargetProbe Missing(string path, bool canWrite)
        {
            return new PublishingPageLayoutResourceTargetProbe
            {
                TargetServerRelativeUrl = path,
                FileExists = false,
                CanWrite = canWrite,
                Availability = EvidenceAvailability.Captured
            };
        }

        private static bool IsFileNotFound(ServerException exception)
        {
            return string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }
    }
}
