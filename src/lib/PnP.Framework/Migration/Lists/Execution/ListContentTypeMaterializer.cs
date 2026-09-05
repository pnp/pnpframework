using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListContentTypeMaterializer
    {
        public static IDictionary<string, string> EnsureMembership(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source)
        {
            context.Load(context.Web.AvailableContentTypes, values => values.Include(value => value.Id, value => value.Name));
            LoadListContentTypes(context, targetList);
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sourceContentType in source.ContentTypes.OrderBy(value => value.Id, StringComparer.OrdinalIgnoreCase))
            {
                var target = FindTarget(targetList.ContentTypes, sourceContentType);
                if (target == null)
                {
                    if (string.IsNullOrWhiteSpace(sourceContentType.ParentId))
                    {
                        throw new InvalidDataException("Source List content type has no site-content-type parent: " + sourceContentType.Id + ".");
                    }
                    var available = context.Web.AvailableContentTypes.AsEnumerable().SingleOrDefault(value =>
                        string.Equals(value.Id.StringValue, sourceContentType.ParentId, StringComparison.OrdinalIgnoreCase));
                    if (available == null)
                    {
                        throw new InvalidDataException("Target Web does not expose required site content type '" + sourceContentType.ParentId + "'.");
                    }
                    target = targetList.ContentTypes.AddExistingContentType(available);
                    context.Load(target, value => value.Id);
                    context.ExecuteQueryRetry();
                    var targetId = target.Id.StringValue;
                    // AddExistingContentType mutates the loaded collection with
                    // an object whose Parent identity is not initialized. The
                    // next source content type search walks every collection
                    // member, so refresh the collection and all Parent IDs before
                    // continuing the membership transaction.
                    LoadListContentTypes(context, targetList);
                    target = LoadContentType(context, targetList, targetId);
                }
                if (target.Parent == null
                    || !string.Equals(target.Parent.Id.StringValue, sourceContentType.ParentId, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("Target List content type parent differs from source mapping: " + sourceContentType.Id + ".");
                }
                result[sourceContentType.Id] = target.Id.StringValue;
            }
            if (result.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != result.Count)
            {
                throw new InvalidDataException("Multiple source List content types resolved to the same target List content type.");
            }
            return result;
        }

        public static void EnsureFieldLinks(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            IDictionary<string, string> contentTypeIds)
        {
            var retainedFieldIds = new HashSet<Guid>(plan.Fields
                .Where(value => value.Disposition != ListFieldMaterializationDisposition.EvidenceOnly)
                .Select(value => value.SourceFieldId));
            foreach (var sourceContentType in source.ContentTypes.OrderBy(value => value.Id, StringComparer.OrdinalIgnoreCase))
            {
                string targetId;
                if (!contentTypeIds.TryGetValue(sourceContentType.Id, out targetId))
                {
                    throw new InvalidDataException("Source List content type has no target identity mapping: " + sourceContentType.Id + ".");
                }
                var target = LoadContentType(context, targetList, targetId);
                if (target.ReadOnly || target.Sealed)
                {
                    target.ReadOnly = false;
                    target.Sealed = false;
                    target.Update(false);
                    context.ExecuteQueryRetry();
                    target = LoadContentType(context, targetList, targetId);
                }
                var links = target.FieldLinks.AsEnumerable().ToDictionary(value => value.Id);
                var fieldsToAdd = new List<Field>();
                var retainedLinks = sourceContentType.FieldLinks
                    .Where(value => retainedFieldIds.Contains(value.FieldId))
                    .OrderBy(value => value.FieldId)
                    .ToArray();
                foreach (var expected in retainedLinks)
                {
                    if (!links.ContainsKey(expected.FieldId))
                    {
                        var field = targetList.Fields.GetById(expected.FieldId);
                        context.Load(field, value => value.Id);
                        fieldsToAdd.Add(field);
                    }
                }
                if (fieldsToAdd.Count > 0)
                {
                    context.ExecuteQueryRetry();
                    foreach (var field in fieldsToAdd)
                    {
                        target.FieldLinks.Add(new FieldLinkCreationInformation { Field = field });
                    }
                    target.Update(false);
                    context.ExecuteQueryRetry();
                    target = LoadContentType(context, targetList, targetId);
                    links = target.FieldLinks.AsEnumerable().ToDictionary(value => value.Id);
                }
                foreach (var expected in retainedLinks)
                {
                    FieldLink link;
                    if (!links.TryGetValue(expected.FieldId, out link))
                    {
                        throw new InvalidDataException("Target List content type did not bind field '" + expected.InternalName + "' (" + expected.FieldId.ToString("D") + ").");
                    }
                    link.DisplayName = expected.DisplayName;
                    link.Required = expected.Required;
                    link.Hidden = expected.Hidden;
                    link.ReadOnly = expected.ReadOnly;
                }
                target.Name = sourceContentType.Name;
                target.Description = sourceContentType.Description;
                target.Group = sourceContentType.Group;
                target.Hidden = sourceContentType.Hidden;
                target.ReadOnly = sourceContentType.ReadOnly;
                target.Sealed = sourceContentType.Sealed;
                target.Update(false);
                context.ExecuteQueryRetry();
                VerifyContentType(context, targetList, targetId, sourceContentType, retainedFieldIds);
            }
        }

        public static void EnsureOrder(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            IDictionary<string, string> contentTypeIds)
        {
            LoadListContentTypes(context, targetList);
            var targetById = targetList.ContentTypes.AsEnumerable()
                .ToDictionary(value => value.Id.StringValue, StringComparer.OrdinalIgnoreCase);
            IList<ContentTypeId> desired = null;
            if (source.HasExplicitUniqueContentTypeOrder)
            {
                desired = new List<ContentTypeId>();
                foreach (var sourceId in source.UniqueContentTypeOrder)
                {
                    string targetId;
                    ContentType target;
                    if (!contentTypeIds.TryGetValue(sourceId, out targetId)
                        || !targetById.TryGetValue(targetId, out target))
                    {
                        throw new InvalidDataException("Source List content type order has no target identity mapping: " + sourceId + ".");
                    }
                    if (target.GetIsAllowedInContentTypeOrder())
                    {
                        desired.Add(target.Id);
                    }
                }
            }

            context.Load(targetList.RootFolder, value => value.UniqueContentTypeOrder);
            context.ExecuteQueryRetry();
            var observed = targetList.RootFolder.UniqueContentTypeOrder == null
                ? null
                : targetList.RootFolder.UniqueContentTypeOrder.Select(value => value.StringValue).ToArray();
            var expected = desired == null ? null : desired.Select(value => value.StringValue).ToArray();
            if (!SameOrder(observed, expected))
            {
                targetList.RootFolder.UniqueContentTypeOrder = desired;
                targetList.RootFolder.Update();
                context.ExecuteQueryRetry();
            }
            context.Load(targetList.RootFolder, value => value.UniqueContentTypeOrder);
            context.ExecuteQueryRetry();
            observed = targetList.RootFolder.UniqueContentTypeOrder == null
                ? null
                : targetList.RootFolder.UniqueContentTypeOrder.Select(value => value.StringValue).ToArray();
            if (!SameOrder(observed, expected))
            {
                throw new InvalidDataException("Fresh target List unique content type order differs from the sealed source order.");
            }
        }

        private static void LoadListContentTypes(ClientContext context, List list)
        {
            context.Load(list.ContentTypes, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.Parent));
            context.ExecuteQueryRetry();
            foreach (var contentType in list.ContentTypes)
            {
                if (contentType.Parent != null)
                {
                    context.Load(contentType.Parent, value => value.Id);
                }
            }
            context.ExecuteQueryRetry();
        }

        private static ContentType FindTarget(ContentTypeCollection contentTypes, ListContentTypeSnapshot source)
        {
            var exact = contentTypes.AsEnumerable().SingleOrDefault(value =>
                string.Equals(value.Id.StringValue, source.Id, StringComparison.OrdinalIgnoreCase));
            if (exact != null)
            {
                return exact;
            }
            var candidates = contentTypes.AsEnumerable().Where(value => value.Parent != null
                    && string.Equals(value.Parent.Id.StringValue, source.ParentId, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (candidates.Length > 1)
            {
                throw new InvalidDataException("Target List exposes ambiguous content type candidates for source '" + source.Id + "'.");
            }
            return candidates.SingleOrDefault();
        }

        private static bool SameOrder(IList<string> left, IList<string> right)
        {
            if (left == null || right == null)
            {
                return left == null && right == null;
            }
            return left.SequenceEqual(right, StringComparer.OrdinalIgnoreCase);
        }

        private static ContentType LoadContentType(ClientContext context, List targetList, string targetId)
        {
            var target = targetList.ContentTypes.GetById(targetId);
            context.Load(target,
                value => value.Id,
                value => value.Name,
                value => value.Description,
                value => value.Group,
                value => value.Hidden,
                value => value.ReadOnly,
                value => value.Sealed,
                value => value.Parent);
            context.Load(target.Parent, value => value.Id);
            context.Load(target.FieldLinks, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.DisplayName,
                value => value.Required,
                value => value.Hidden,
                value => value.ReadOnly));
            context.ExecuteQueryRetry();
            return target;
        }

        private static void VerifyContentType(
            ClientContext context,
            List targetList,
            string targetId,
            ListContentTypeSnapshot source,
            ISet<Guid> retainedFieldIds)
        {
            var target = LoadContentType(context, targetList, targetId);
            if (!string.Equals(target.Name, source.Name, StringComparison.Ordinal)
                || !string.Equals(target.Description ?? string.Empty, source.Description ?? string.Empty, StringComparison.Ordinal)
                || !string.Equals(target.Group ?? string.Empty, source.Group ?? string.Empty, StringComparison.Ordinal)
                || target.Hidden != source.Hidden
                || target.ReadOnly != source.ReadOnly
                || target.Sealed != source.Sealed
                || target.Parent == null
                || !string.Equals(target.Parent.Id.StringValue, source.ParentId, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Fresh target List content type metadata differs: " + source.Id + ".");
            }
            var actual = target.FieldLinks.AsEnumerable().ToDictionary(value => value.Id);
            foreach (var expected in source.FieldLinks.Where(value => retainedFieldIds.Contains(value.FieldId)))
            {
                FieldLink link;
                if (!actual.TryGetValue(expected.FieldId, out link)
                    || !string.Equals(link.DisplayName ?? string.Empty, expected.DisplayName ?? string.Empty, StringComparison.Ordinal)
                    || link.Required != expected.Required
                    || link.Hidden != expected.Hidden
                    || link.ReadOnly != expected.ReadOnly)
                {
                    throw new InvalidDataException("Fresh target List content type field-link readback differs: " + source.Id + "/" + expected.FieldId.ToString("D") + ".");
                }
            }
        }
    }
}
