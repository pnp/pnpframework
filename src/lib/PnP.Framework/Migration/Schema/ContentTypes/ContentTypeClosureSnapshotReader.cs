using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.ContentTypes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    internal static class ContentTypeClosureSnapshotReader
    {
        public static IList<ContentTypeSchemaSnapshot> Read(
            ClientContext context,
            Web sourceWeb,
            IEnumerable<ListContentTypeSnapshot> listContentTypes,
            ICollection<string> diagnostics)
        {
            var pending = new Queue<string>((listContentTypes ?? Enumerable.Empty<ListContentTypeSnapshot>())
                .Select(value => value.ParentId)
                .Where(value => !string.IsNullOrWhiteSpace(value) && !ContentTypeRuntimeCatalog.IsTargetRuntime(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
            var observed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<ContentTypeSchemaSnapshot>();
            while (pending.Count > 0)
            {
                var contentTypeId = pending.Dequeue();
                if (!observed.Add(contentTypeId) || ContentTypeRuntimeCatalog.IsTargetRuntime(contentTypeId))
                {
                    continue;
                }

                var contentType = sourceWeb.AvailableContentTypes.GetById(contentTypeId);
                context.Load(contentType, value => value.Id, value => value.Scope);
                context.Load(contentType.Parent, value => value.Id);
                context.Load(contentType.FieldLinks, values => values.Include(value => value.Name));
                context.ExecuteQueryRetry();
                if (contentType.ServerObjectIsNull.GetValueOrDefault(true))
                {
                    throw new InvalidOperationException("Source site content type is unavailable: " + contentTypeId + ".");
                }

                var localDiagnostics = new List<string>();
                var snapshot = ContentTypeSchemaSnapshotReader.Read(
                    context,
                    sourceWeb,
                    contentTypeId,
                    contentType.FieldLinks.Select(value => value.Name),
                    localDiagnostics);
                snapshot.SourceScope = contentType.Scope;
                snapshot.SourceWebUrl = ScopeOwnerUrl(contentType.Scope, sourceWeb.Url);
                result.Add(snapshot);
                foreach (var diagnostic in localDiagnostics)
                {
                    diagnostics?.Add("Site content type '" + contentTypeId + "': " + diagnostic);
                }

                var parentId = contentType.Parent == null || contentType.Parent.ServerObjectIsNull.GetValueOrDefault(true)
                    ? null
                    : contentType.Parent.Id.StringValue;
                if (!string.IsNullOrWhiteSpace(parentId)
                    && !ContentTypeRuntimeCatalog.IsTargetRuntime(parentId)
                    && !observed.Contains(parentId))
                {
                    pending.Enqueue(parentId);
                }
            }
            return result.OrderBy(value => value.ContentTypeId.Length)
                .ThenBy(value => value.ContentTypeId, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string ScopeOwnerUrl(string scope, string fallbackWebUrl)
        {
            if (string.IsNullOrWhiteSpace(scope))
            {
                return fallbackWebUrl.TrimEnd('/');
            }
            Uri absolute;
            if (Uri.TryCreate(scope, UriKind.Absolute, out absolute)
                && (string.Equals(absolute.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(absolute.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
            {
                return absolute.AbsoluteUri.TrimEnd('/');
            }
            var origin = new Uri(fallbackWebUrl).GetLeftPart(UriPartial.Authority).TrimEnd('/');
            return new Uri(origin + "/" + scope.Trim('/')).AbsoluteUri.TrimEnd('/');
        }
    }
}
