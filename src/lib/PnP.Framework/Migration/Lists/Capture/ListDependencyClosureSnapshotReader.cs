using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Capture
{
    internal sealed class ListDependencyClosureCaptureResult
    {
        public IList<ListDependencySnapshot> Dependencies { get; set; } = new List<ListDependencySnapshot>();

        public IList<ListLookupDependency> LookupDependencies { get; set; } = new List<ListLookupDependency>();

        public IList<Guid> RequiredSourceWebIds { get; set; } = new List<Guid>();
    }

    internal static class ListDependencyClosureSnapshotReader
    {
        public static ListDependencyClosureCaptureResult Read(
            ClientContext context,
            IEnumerable<ClassicListWebPartBindingSnapshot> bindings,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var result = new ListDependencyClosureCaptureResult();
            var queue = new Queue<ListIdentity>((bindings ?? Enumerable.Empty<ClassicListWebPartBindingSnapshot>())
                .Select(value => new ListIdentity(value.SourceListWebId, value.SourceListId)));
            var captured = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (queue.Count > 0)
            {
                var identity = queue.Dequeue();
                var key = identity.WebId.ToString("D") + ":" + identity.ListId.ToString("D");
                if (!captured.Add(key))
                {
                    continue;
                }

                try
                {
                    var web = context.Site.OpenWebById(identity.WebId);
                    var dependency = ListDependencySnapshotReader.Read(context, web, identity.ListId, maximumBytes, artifactStore, warnings);
                    result.Dependencies.Add(dependency);
                    result.RequiredSourceWebIds.Add(dependency.SourceWebId);
                    foreach (var field in dependency.Fields.Where(ShouldFollowLookupDependency))
                    {
                        var lookupWebId = field.SourceLookupWebId ?? dependency.SourceWebId;
                        var edge = new ListLookupDependency
                        {
                            SourceListId = dependency.SourceListId,
                            LookupListId = field.SourceLookupListId.Value,
                            FieldId = field.Id,
                            FieldInternalName = field.InternalName
                        };
                        result.LookupDependencies.Add(edge);
                        queue.Enqueue(new ListIdentity(lookupWebId, field.SourceLookupListId.Value));
                    }
                }
                catch (Exception exception) when (exception is ServerException || exception is InvalidOperationException || exception is IOException)
                {
                    blockers.Add("List dependency " + identity.ListId.ToString("D") + " in source Web " + identity.WebId.ToString("D") + " could not be captured: " + exception.Message);
                }
            }

            result.Dependencies = result.Dependencies.OrderBy(value => value.SourceWebId).ThenBy(value => value.SourceListId).ToList();
            result.LookupDependencies = result.LookupDependencies.OrderBy(value => value.LookupListId).ThenBy(value => value.SourceListId).ThenBy(value => value.FieldId).ToList();
            result.RequiredSourceWebIds = result.RequiredSourceWebIds.Distinct().OrderBy(value => value).ToList();
            return result;
        }

        internal static bool ShouldFollowLookupDependency(ListFieldSnapshot field)
        {
            if (field?.SourceLookupListId.HasValue != true)
            {
                return false;
            }

            // A taxonomy field's List attribute points at the site-local
            // TaxonomyHiddenList WssId cache. That cache is platform-owned and
            // must not be copied as a cross-site business List dependency;
            // SharePoint allocates target WssIds when taxonomy values are set.
            // The field binding and source term identifiers remain fully
            // captured by TaxonomyFieldBindingSnapshot.
            return !field.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.OrdinalIgnoreCase)
                && !ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field);
        }

        private sealed class ListIdentity
        {
            public ListIdentity(Guid webId, Guid listId)
            {
                WebId = webId;
                ListId = listId;
            }

            public Guid WebId { get; }

            public Guid ListId { get; }
        }
    }
}
