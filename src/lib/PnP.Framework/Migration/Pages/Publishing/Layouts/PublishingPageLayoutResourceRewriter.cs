using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutResourceRewriter
    {
        public static byte[] Rewrite(
            byte[] sourceBytes,
            IEnumerable<PublishingPageLayoutResourceRewrite> rewrites)
        {
            if (sourceBytes == null)
            {
                throw new ArgumentNullException(nameof(sourceBytes));
            }

            if (rewrites == null)
            {
                throw new ArgumentNullException(nameof(rewrites));
            }

            var text = PublishingPageLayoutEncoding.Decode(sourceBytes);
            foreach (var rewrite in rewrites.OrderByDescending(value => value.SourceReference?.Length ?? 0))
            {
                if (rewrite == null
                    || string.IsNullOrWhiteSpace(rewrite.SourceReference)
                    || string.IsNullOrWhiteSpace(rewrite.TargetReference))
                {
                    throw new ArgumentException("Every Page Layout resource rewrite requires source and target references.", nameof(rewrites));
                }

                var before = text;
                text = text.Replace(rewrite.SourceReference, rewrite.TargetReference);
                text = text.Replace(WebUtility.HtmlEncode(rewrite.SourceReference), WebUtility.HtmlEncode(rewrite.TargetReference));
                if (string.Equals(before, text, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException($"Page Layout resource reference was not present in source bytes: {rewrite.SourceReference}");
                }
            }

            return PublishingPageLayoutEncoding.Encode(text, sourceBytes);
        }
    }
}
