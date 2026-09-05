using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListSchemaVerifier
    {
        public static void Verify(
            ClientContext context,
            List list,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationReceipt receipt,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            ListMaterializationExecutionScope.ListSelection selection,
            ICollection<string> diagnostics)
        {
            context.Load(list.Fields, values => values.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.TypeAsString,
                value => value.SchemaXml));
            context.Load(list.ContentTypes, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.Description,
                value => value.Group,
                value => value.Hidden,
                value => value.ReadOnly,
                value => value.Sealed,
                value => value.Parent));
            context.Load(list.Views, values => values.Include(
                value => value.Id,
                value => value.Title,
                value => value.PersonalView,
                value => value.ViewType,
                value => value.ViewQuery,
                value => value.RowLimit,
                value => value.Paged,
                value => value.JSLink));
            context.ExecuteQueryRetry();
            foreach (var contentType in list.ContentTypes)
            {
                if (contentType.Parent != null)
                {
                    context.Load(contentType.Parent, value => value.Id);
                }
                context.Load(contentType.FieldLinks, values => values.Include(
                    value => value.Id,
                    value => value.DisplayName,
                    value => value.Required,
                    value => value.Hidden,
                    value => value.ReadOnly));
            }
            foreach (var view in list.Views)
            {
                context.Load(view.ViewFields);
            }
            context.ExecuteQueryRetry();

            VerifyFields(list, plan, dependencyReceipts, receipt, diagnostics);
            VerifyContentTypes(list, source, plan, receipt, selection, diagnostics);
            receipt.VerifiedViewRenderingResourceCount = ListViewRenderingResourceMaterializer.Verify(context, plan, diagnostics);
            VerifyViews(list, plan, receipt, diagnostics);
        }

        private static void VerifyFields(
            List list,
            ListMaterializationPlan plan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            ListMaterializationReceipt receipt,
            ICollection<string> diagnostics)
        {
            var actual = list.Fields.AsEnumerable().ToDictionary(value => value.Id);
            foreach (var fieldPlan in plan.Fields.Where(value => value.Disposition != ListFieldMaterializationDisposition.EvidenceOnly
                         && value.Disposition != ListFieldMaterializationDisposition.Block))
            {
                Field field;
                if (!actual.TryGetValue(fieldPlan.SourceFieldId, out field))
                {
                    diagnostics.Add("Target field is missing: " + fieldPlan.InternalName + " (" + fieldPlan.SourceFieldId.ToString("D") + ").");
                    continue;
                }
                if (!string.Equals(field.InternalName, fieldPlan.InternalName, StringComparison.OrdinalIgnoreCase)
                    || !ListFieldTypeCompatibility.IsCompatibleRuntimeType(field.TypeAsString, fieldPlan.TypeAsString))
                {
                    diagnostics.Add("Target field identity/type differs: " + fieldPlan.InternalName + ".");
                    continue;
                }
                if (fieldPlan.Disposition != ListFieldMaterializationDisposition.RequireTargetRuntime
                    && fieldPlan.Disposition != ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue)
                {
                    var expectedSchema = fieldPlan.TargetSchemaXml;
                    if (fieldPlan.Disposition == ListFieldMaterializationDisposition.MapLookup)
                    {
                        ListMaterializationReceipt lookup;
                        if (!fieldPlan.SourceLookupListId.HasValue
                            || !dependencyReceipts.TryGetValue(fieldPlan.SourceLookupListId.Value, out lookup))
                        {
                            diagnostics.Add("Lookup field has no target dependency receipt: " + fieldPlan.InternalName + ".");
                            continue;
                        }
                        expectedSchema = FieldSchemaCanonicalizer.RewriteLookupForTarget(fieldPlan.SourceSchemaXml, lookup.TargetWebId, lookup.TargetListId);
                    }
                    if (string.IsNullOrWhiteSpace(expectedSchema)
                        || !string.Equals(FieldSchemaCanonicalizer.PortableDigest(field.SchemaXml), FieldSchemaCanonicalizer.PortableDigest(expectedSchema), StringComparison.OrdinalIgnoreCase))
                    {
                        diagnostics.Add("Target field portable schema differs: " + fieldPlan.InternalName + ".");
                        continue;
                    }
                }
                receipt.VerifiedFieldCount++;
            }
        }

        private static void VerifyContentTypes(
            List list,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationReceipt receipt,
            ListMaterializationExecutionScope.ListSelection selection,
            ICollection<string> diagnostics)
        {
            var retainedFieldIds = new HashSet<Guid>(plan.Fields
                .Where(value => value.Disposition != ListFieldMaterializationDisposition.EvidenceOnly)
                .Select(value => value.SourceFieldId));
            var actual = list.ContentTypes.AsEnumerable().ToDictionary(value => value.Id.StringValue, StringComparer.OrdinalIgnoreCase);
            if (receipt.TargetContentTypeIds.Count != source.ContentTypes.Count)
            {
                diagnostics.Add("Receipt content type mapping count differs from the source List content type count.");
            }
            foreach (var sourceContentType in source.ContentTypes)
            {
                string targetId;
                ContentType target;
                if (!receipt.TargetContentTypeIds.TryGetValue(sourceContentType.Id, out targetId)
                    || !actual.TryGetValue(targetId, out target))
                {
                    diagnostics.Add("Target List content type mapping is missing: " + sourceContentType.Id + ".");
                    continue;
                }
                if (!string.Equals(target.Name, sourceContentType.Name, StringComparison.Ordinal)
                    || !string.Equals(target.Description ?? string.Empty, sourceContentType.Description ?? string.Empty, StringComparison.Ordinal)
                    || !string.Equals(target.Group ?? string.Empty, sourceContentType.Group ?? string.Empty, StringComparison.Ordinal)
                    || target.Hidden != sourceContentType.Hidden
                    || target.ReadOnly != sourceContentType.ReadOnly
                    || target.Sealed != sourceContentType.Sealed
                    || target.Parent == null
                    || !string.Equals(target.Parent.Id.StringValue, sourceContentType.ParentId, StringComparison.OrdinalIgnoreCase))
                {
                    diagnostics.Add("Target List content type name/parent differs: " + sourceContentType.Id + ".");
                    continue;
                }
                var links = target.FieldLinks.AsEnumerable().ToDictionary(value => value.Id);
                var linkMismatch = sourceContentType.FieldLinks
                    .Where(expected => retainedFieldIds.Contains(expected.FieldId))
                    .Any(expected =>
                {
                    FieldLink observed;
                    return !links.TryGetValue(expected.FieldId, out observed)
                        || !string.Equals(observed.DisplayName ?? string.Empty, expected.DisplayName ?? string.Empty, StringComparison.Ordinal)
                        || observed.Required != expected.Required
                        || observed.Hidden != expected.Hidden
                        || observed.ReadOnly != expected.ReadOnly;
                });
                if (linkMismatch)
                {
                    diagnostics.Add("Target List content type FieldLinks differ: " + sourceContentType.Id + ".");
                    continue;
                }
                receipt.VerifiedContentTypeCount++;
            }

            if (selection == null || selection.ExactContentTypeInventory)
            {
                var expectedOrder = ExpectedContentTypeOrder(source, receipt, actual);
                var observedOrder = list.RootFolder.UniqueContentTypeOrder == null
                    ? null
                    : list.RootFolder.UniqueContentTypeOrder.Select(value => value.StringValue).ToArray();
                if (!SameOrder(observedOrder, expectedOrder))
                {
                    diagnostics.Add("Target List unique content type order differs from the sealed source order.");
                }
            }
        }

        private static string[] ExpectedContentTypeOrder(
            ListDependencySnapshot source,
            ListMaterializationReceipt receipt,
            IDictionary<string, ContentType> actual)
        {
            if (!source.HasExplicitUniqueContentTypeOrder)
            {
                return null;
            }
            var result = new List<string>();
            foreach (var sourceId in source.UniqueContentTypeOrder)
            {
                string targetId;
                ContentType target;
                if (receipt.TargetContentTypeIds.TryGetValue(sourceId, out targetId)
                    && actual.TryGetValue(targetId, out target)
                    && target.GetIsAllowedInContentTypeOrder())
                {
                    result.Add(targetId);
                }
            }
            return result.ToArray();
        }

        private static void VerifyViews(
            List list,
            ListMaterializationPlan plan,
            ListMaterializationReceipt receipt,
            ICollection<string> diagnostics)
        {
            var expected = plan.Views.Where(value => value.Disposition == ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView
                    || value.Disposition == ListViewMaterializationDisposition.CreateOrReuseWebPartView)
                .ToArray();
            var actual = list.Views.AsEnumerable().ToDictionary(value => value.Id);
            if (receipt.TargetViewIds.Count != expected.Length)
            {
                diagnostics.Add("Receipt View mapping count differs from the executable source View count.");
            }
            foreach (var viewPlan in expected)
            {
                Guid targetId;
                View target;
                if (!receipt.TargetViewIds.TryGetValue(viewPlan.SourceViewId, out targetId)
                    || !actual.TryGetValue(targetId, out target))
                {
                    diagnostics.Add("Target View mapping is missing: " + viewPlan.SourceViewId.ToString("D") + ".");
                    continue;
                }
                var source = viewPlan.Source;
                var expectedJsLink = ListViewRenderingResourceMaterializer.RewriteJsLink(
                    source.JsLink,
                    source,
                    plan.ViewRenderingResources);
                if (target.PersonalView
                    || !string.Equals(target.Title, ListViewMaterializer.TargetTitle(viewPlan), StringComparison.Ordinal)
                    || !string.Equals(target.ViewType, source.ViewType, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(target.ViewQuery ?? string.Empty, source.ViewQuery ?? string.Empty, StringComparison.Ordinal)
                    || target.RowLimit != source.RowLimit
                    || target.Paged != source.Paged
                    || !string.Equals(target.JSLink ?? string.Empty, expectedJsLink ?? string.Empty, StringComparison.Ordinal)
                    || !target.ViewFields.SequenceEqual(source.ViewFields, StringComparer.OrdinalIgnoreCase))
                {
                    diagnostics.Add("Target View readback differs: " + viewPlan.SourceViewId.ToString("D") + ".");
                    continue;
                }
                receipt.VerifiedViewCount++;
            }
        }

        private static bool SameOrder(IList<string> left, IList<string> right)
        {
            if (left == null || right == null)
            {
                return left == null && right == null;
            }
            return left.SequenceEqual(right, StringComparer.OrdinalIgnoreCase);
        }
    }
}
