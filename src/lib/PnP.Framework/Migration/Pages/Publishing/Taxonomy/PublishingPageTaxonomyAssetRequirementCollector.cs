using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Taxonomy.Assets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Taxonomy
{
    /// <summary>
    /// Projects page and List snapshots into one shared taxonomy dependency
    /// closure. Complete field evidence remains in each page package while the
    /// expensive tenant-wide taxonomy asset capture can be deduplicated per run.
    /// </summary>
    public static class PublishingPageTaxonomyAssetRequirementCollector
    {
        public static IList<TaxonomyTermSetCaptureRequest> Collect(
            IEnumerable<PublishingPageCaptureBundle> snapshots)
        {
            var requests = new Dictionary<string, MutableRequest>(StringComparer.Ordinal);
            foreach (var snapshot in (snapshots ?? Enumerable.Empty<PublishingPageCaptureBundle>()).Where(value => value != null))
            {
                AddPageFields(
                    snapshot.Fields,
                    snapshot.Source?.WebUrl,
                    snapshot.Source?.WebId ?? Guid.Empty,
                    snapshot.Source?.PageServerRelativeUrl,
                    requests);
                AddLayoutFields(snapshot.Layout?.AssociatedContentTypeSchema, requests);
                foreach (var list in (snapshot.ListDependencies ?? new List<ListDependencySnapshot>()).Where(value => value != null))
                {
                    AddListFields(list, requests);
                    AddSiteFields(list, requests);
                }
            }
            return requests.Values
                .OrderBy(value => value.StoreId)
                .ThenBy(value => value.SetId)
                .Select(value => new TaxonomyTermSetCaptureRequest
                {
                    SourceTermStoreId = value.StoreId,
                    SourceTermSetId = value.SetId,
                    SourceWebUrls = value.SourceWebUrls.OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToList(),
                    RequiredTermIds = value.TermIds.OrderBy(item => item).ToList(),
                    Consumers = value.Consumers.OrderBy(item => item, StringComparer.Ordinal).ToList()
                })
                .ToList();
        }

        private static void AddLayoutFields(
            ContentTypeSchemaSnapshot contentType,
            IDictionary<string, MutableRequest> requests)
        {
            if (contentType == null)
            {
                return;
            }
            foreach (var field in (contentType.RequiredFieldClosure ?? new List<FieldSchemaSnapshot>())
                         .Where(value => value?.Taxonomy != null))
            {
                var consumer = "page-layout-field:" + NormalizeScope(contentType.SourceScope, contentType.SourceWebUrl)
                    + "/" + field.Id.ToString("D") + "/content-type:" + contentType.ContentTypeId;
                Add(
                    requests,
                    field.Taxonomy.SourceTermStoreId,
                    field.Taxonomy.SourceTermSetId,
                    field.Taxonomy.AnchorTermId != Guid.Empty ? field.Taxonomy.AnchorTermId : ParseAnchor(field.SchemaXml),
                    contentType.SourceWebUrl,
                    consumer);
            }
        }

        private static void AddPageFields(
            IEnumerable<PageFieldValueSnapshot> fields,
            string sourceWebUrl,
            Guid sourceWebId,
            string sourcePageServerRelativeUrl,
            IDictionary<string, MutableRequest> requests)
        {
            foreach (var field in (fields ?? Enumerable.Empty<PageFieldValueSnapshot>())
                         .Where(value => value?.TaxonomyBinding != null))
            {
                var binding = field.TaxonomyBinding;
                var consumer = "page-field:" + sourceWebId.ToString("D")
                    + "/" + NormalizeScope(sourcePageServerRelativeUrl, sourceWebUrl)
                    + "/" + field.Id.ToString("D") + "/" + field.InternalName;
                Add(
                    requests,
                    binding.TermStoreId,
                    binding.BoundTermSetId,
                    binding.AnchorTermId != Guid.Empty ? binding.AnchorTermId : ParseAnchor(field.SchemaXml),
                    sourceWebUrl,
                    consumer);
                foreach (var value in (field.TaxonomyValues ?? new List<PageTaxonomyValueSnapshot>()).Where(item => item?.Relationship != null))
                {
                    if (!Guid.TryParse(value.TermGuid, out var termId) || termId == Guid.Empty)
                    {
                        continue;
                    }
                    if (value.Relationship.State == TaxonomyRelationshipState.LiveInBoundTermSet)
                    {
                        Add(requests, binding.TermStoreId, binding.BoundTermSetId, termId, sourceWebUrl, consumer + "/value:" + value.WssId);
                    }
                    else if (value.Relationship.State == TaxonomyRelationshipState.LiveOutsideBoundTermSet
                        && value.Relationship.LiveTermSetId.HasValue)
                    {
                        Add(requests, binding.TermStoreId, value.Relationship.LiveTermSetId.Value, termId, sourceWebUrl, consumer + "/value:" + value.WssId);
                    }
                }
            }
        }

        private static void AddListFields(
            ListDependencySnapshot list,
            IDictionary<string, MutableRequest> requests)
        {
            foreach (var field in (list.Fields ?? new List<ListFieldSnapshot>())
                         .Where(value => value?.Taxonomy != null))
            {
                var consumer = "list-field:" + list.SourceWebId.ToString("D") + "/" + list.SourceListId.ToString("D") + "/" + field.Id.ToString("D");
                var termIds = (list.Items ?? new List<ListItemSnapshot>())
                    .Where(item => item != null)
                    .SelectMany(item => item.Values ?? new List<ListItemValueSnapshot>())
                    .Where(value => value != null
                        && string.Equals(value.InternalName, field.InternalName, StringComparison.OrdinalIgnoreCase))
                    .SelectMany(value => value.TaxonomyValues ?? new List<ListItemTaxonomyValueSnapshot>())
                    .Select(value => ParseGuid(value?.TermGuid))
                    .Where(value => value.HasValue)
                    .Select(value => value.Value);
                Add(
                    requests,
                    field.Taxonomy.SourceTermStoreId,
                    field.Taxonomy.SourceTermSetId,
                    new[] { field.Taxonomy.AnchorTermId != Guid.Empty ? field.Taxonomy.AnchorTermId : ParseAnchor(field.SchemaXml) }
                        .Where(value => value != Guid.Empty)
                        .Concat(termIds),
                    list.SourceWebUrl,
                    consumer);
            }
        }

        private static void AddSiteFields(
            ListDependencySnapshot list,
            IDictionary<string, MutableRequest> requests)
        {
            foreach (var contentType in (list.SiteContentTypes ?? new List<ContentTypeSchemaSnapshot>())
                         .Where(value => value != null))
            {
                foreach (var field in (contentType.RequiredFieldClosure ?? new List<FieldSchemaSnapshot>())
                             .Where(value => value?.Taxonomy != null))
                {
                    var consumer = "site-field:" + NormalizeScope(contentType.SourceScope, contentType.SourceWebUrl)
                        + "/" + field.Id.ToString("D") + "/content-type:" + contentType.ContentTypeId;
                    Add(
                        requests,
                        field.Taxonomy.SourceTermStoreId,
                        field.Taxonomy.SourceTermSetId,
                        field.Taxonomy.AnchorTermId != Guid.Empty ? field.Taxonomy.AnchorTermId : ParseAnchor(field.SchemaXml),
                        contentType.SourceWebUrl ?? list.SourceWebUrl,
                        consumer);
                }
            }
        }

        private static void Add(
            IDictionary<string, MutableRequest> requests,
            Guid storeId,
            Guid setId,
            Guid termId,
            string sourceWebUrl,
            string consumer)
        {
            Add(requests, storeId, setId, termId == Guid.Empty ? Enumerable.Empty<Guid>() : new[] { termId }, sourceWebUrl, consumer);
        }

        private static void Add(
            IDictionary<string, MutableRequest> requests,
            Guid storeId,
            Guid setId,
            IEnumerable<Guid> termIds,
            string sourceWebUrl,
            string consumer)
        {
            if (storeId == Guid.Empty || setId == Guid.Empty)
            {
                return;
            }
            var key = storeId.ToString("D") + "/" + setId.ToString("D");
            if (!requests.TryGetValue(key, out var request))
            {
                request = new MutableRequest(storeId, setId);
                requests[key] = request;
            }
            foreach (var termId in (termIds ?? Enumerable.Empty<Guid>()).Where(value => value != Guid.Empty))
            {
                request.TermIds.Add(termId);
            }
            if (!string.IsNullOrWhiteSpace(sourceWebUrl))
            {
                request.SourceWebUrls.Add(sourceWebUrl);
            }
            if (!string.IsNullOrWhiteSpace(consumer))
            {
                request.Consumers.Add(consumer);
            }
        }

        private static Guid ParseAnchor(string schemaXml)
        {
            try
            {
                var property = XDocument.Parse(schemaXml ?? string.Empty)
                    .Descendants()
                    .Where(value => string.Equals(value.Name.LocalName, "Property", StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault(value => string.Equals(
                        value.Elements().FirstOrDefault(item => string.Equals(item.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))?.Value,
                        "AnchorId",
                        StringComparison.OrdinalIgnoreCase));
                return ParseGuid(property?.Elements().FirstOrDefault(value => string.Equals(value.Name.LocalName, "Value", StringComparison.OrdinalIgnoreCase))?.Value)
                    ?? Guid.Empty;
            }
            catch (System.Xml.XmlException)
            {
                return Guid.Empty;
            }
        }

        private static Guid? ParseGuid(string value)
        {
            return Guid.TryParse((value ?? string.Empty).Trim().Trim('{', '}'), out var result) && result != Guid.Empty
                ? result
                : (Guid?)null;
        }

        private static string NormalizeScope(string scope, string webUrl)
        {
            var value = scope;
            if (string.IsNullOrWhiteSpace(value) && Uri.TryCreate(webUrl, UriKind.Absolute, out var uri))
            {
                value = uri.AbsolutePath;
            }
            var normalized = Uri.UnescapeDataString(value ?? string.Empty).Replace('\\', '/').TrimEnd('/');
            return normalized.Length == 0 ? "/" : normalized;
        }

        private sealed class MutableRequest
        {
            public MutableRequest(Guid storeId, Guid setId)
            {
                StoreId = storeId;
                SetId = setId;
            }

            public Guid StoreId { get; }

            public Guid SetId { get; }

            public ISet<Guid> TermIds { get; } = new HashSet<Guid>();

            public ISet<string> SourceWebUrls { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public ISet<string> Consumers { get; } = new HashSet<string>(StringComparer.Ordinal);
        }
    }
}
