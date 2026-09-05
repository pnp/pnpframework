using Microsoft.SharePoint.Client;
using System;
using System.IO;
using CsomFile = Microsoft.SharePoint.Client.File;

namespace PnP.Framework.Migration.Pages.Capture
{
    internal static class SourcePageFenceReader
    {
        public static SourcePageFence Read(ClientContext context, string pagePath)
        {
            var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            context.Load(file, value => value.Exists, value => value.UniqueId, value => value.UIVersionLabel, value => value.Length, value => value.TimeLastModified);
            context.ExecuteQueryRetry();
            if (!file.Exists)
            {
                throw new FileNotFoundException("The source page disappeared during capture.", pagePath);
            }

            return FromFile(file);
        }

        public static SourcePageFence FromFile(CsomFile file)
        {
            return new SourcePageFence
            {
                FileUniqueId = file.UniqueId,
                VersionLabel = file.UIVersionLabel,
                Length = file.Length,
                ModifiedUtc = file.TimeLastModified.ToUniversalTime()
            };
        }

        public static bool Equals(SourcePageFence left, SourcePageFence right)
        {
            return left != null
                && right != null
                && left.FileUniqueId == right.FileUniqueId
                && left.Length == right.Length
                && left.ModifiedUtc == right.ModifiedUtc
                && string.Equals(left.VersionLabel, right.VersionLabel, StringComparison.Ordinal);
        }
    }
}
