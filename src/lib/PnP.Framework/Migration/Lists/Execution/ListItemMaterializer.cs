using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListItemMaterializer
    {
        private static readonly Guid OriginalItemIdFieldId = new Guid("6cf2fb94-16be-41d2-8905-81c4112825d4");
        private static readonly Guid OriginalItemDigestFieldId = new Guid("132755e6-5e26-4a22-b4fd-55ad3f80be67");
        internal const string OriginalItemIdFieldName = "PnPMigrationOriginalItemId";
        internal const string OriginalItemDigestFieldName = "PnPMigrationOriginalItemDigest";

        public static IDictionary<int, int> Ensure(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            IDictionary<string, string> contentTypeIds,
            IMigrationArtifactStore artifactStore)
        {
            var selection = new ListMaterializationExecutionScope.ListSelection
            {
                SourceListId = source.SourceListId,
                IncludeListObject = true,
                ExactContentTypeInventory = true,
                ExactItemInventory = true
            };
            foreach (var item in source.Items.Where(value => value != null))
            {
                selection.ItemIds.Add(item.SourceItemId);
                if (item.Document != null)
                {
                    selection.DocumentItemIds.Add(item.SourceItemId);
                }
                foreach (var attachment in item.Attachments.Where(value => value != null))
                {
                    selection.AddAttachment(item.SourceItemId, attachment.FileName);
                }
                selection.ExactAttachmentInventoryItemIds.Add(item.SourceItemId);
            }
            return Ensure(
                context,
                targetList,
                source,
                plan,
                selection,
                dependencyReceipts,
                contentTypeIds,
                artifactStore);
        }

        public static IDictionary<int, int> Ensure(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationExecutionScope.ListSelection selection,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            IDictionary<string, string> contentTypeIds,
            IMigrationArtifactStore artifactStore)
        {
            if (selection == null || selection.SourceListId != source.SourceListId)
            {
                throw new ArgumentException("The List item execution selection does not match the source List.", nameof(selection));
            }
            EnsureReservedFields(context, targetList);
            context.Load(targetList.RootFolder, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var existingBySourceId = ReadExisting(context, targetList);
            var targetItems = new Dictionary<int, ListItem>();
            foreach (var sourceItem in OrderItems(source.Items))
            {
                var includeItem = selection.ItemIds.Contains(sourceItem.SourceItemId);
                var expectedDigest = includeItem ? ComputeItemDigest(sourceItem) : null;
                ListItem targetItem;
                if (existingBySourceId.TryGetValue(sourceItem.SourceItemId, out targetItem))
                {
                    var observed = ReadString(targetItem, OriginalItemDigestFieldName);
                    if (includeItem
                        && !string.IsNullOrWhiteSpace(observed)
                        && !string.Equals(observed, expectedDigest, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException("Target List item provenance collision for source item " + sourceItem.SourceItemId + ".");
                    }
                }
                else
                {
                    targetItem = ListDocumentMaterializer.CreateItem(context, targetList, source, plan, sourceItem, artifactStore);
                }
                if (includeItem)
                {
                    if (ListItemValueWriter.ApplyContentType(context, targetItem, sourceItem, contentTypeIds))
                    {
                        // Get a clean CSOM object after changing ContentTypeId.
                        // Otherwise later Update calls can resend the dirty
                        // ContentTypeId and re-apply defaults over authored values.
                        targetItem = targetList.GetItemById(targetItem.Id);
                        context.Load(targetItem, value => value.Id);
                        context.ExecuteQueryRetry();
                    }
                    ListItemValueWriter.Apply(context, targetList, targetItem, sourceItem, plan, dependencyReceipts, contentTypeIds, false);
                }
                targetItems[sourceItem.SourceItemId] = targetItem;
                if (selection.AttachmentNamesByItemId.ContainsKey(sourceItem.SourceItemId))
                {
                    ListAttachmentMaterializer.Ensure(context, targetItem, sourceItem.Attachments, artifactStore);
                }
            }

            foreach (var sourceItem in source.Items
                         .Where(value => selection.ItemIds.Contains(value.SourceItemId))
                         .OrderBy(value => value.SourceItemId))
            {
                var targetItem = targetItems[sourceItem.SourceItemId];
                ListItemValueWriter.Apply(context, targetList, targetItem, sourceItem, plan, dependencyReceipts, contentTypeIds, true);
                targetItem[OriginalItemIdFieldName] = sourceItem.SourceItemId;
                targetItem[OriginalItemDigestFieldName] = ComputeItemDigest(sourceItem);
                targetItem.Update();
                context.ExecuteQueryRetry();
            }
            return targetItems.ToDictionary(value => value.Key, value => value.Value.Id);
        }

        private static void EnsureReservedFields(ClientContext context, List list)
        {
            context.Load(list.Fields, values => values.Include(value => value.Id, value => value.InternalName, value => value.TypeAsString));
            context.ExecuteQueryRetry();
            EnsureReservedField(context, list, OriginalItemIdFieldId, OriginalItemIdFieldName, "Integer",
                "<Field ID=\"{" + OriginalItemIdFieldId.ToString("D") + "}\" Name=\"" + OriginalItemIdFieldName + "\" DisplayName=\"PnP migration source item ID\" Type=\"Integer\" Hidden=\"TRUE\" Required=\"FALSE\" />");
            EnsureReservedField(context, list, OriginalItemDigestFieldId, OriginalItemDigestFieldName, "Text",
                "<Field ID=\"{" + OriginalItemDigestFieldId.ToString("D") + "}\" Name=\"" + OriginalItemDigestFieldName + "\" DisplayName=\"PnP migration source item digest\" Type=\"Text\" Hidden=\"TRUE\" Required=\"FALSE\" MaxLength=\"64\" />");
        }

        private static void EnsureReservedField(
            ClientContext context,
            List list,
            Guid id,
            string internalName,
            string type,
            string schemaXml)
        {
            var byId = list.Fields.AsEnumerable().SingleOrDefault(value => value.Id == id);
            var byName = list.Fields.AsEnumerable().SingleOrDefault(value => string.Equals(value.InternalName, internalName, StringComparison.OrdinalIgnoreCase));
            if (byId != null)
            {
                if (byName == null || byName.Id != id
                    || !string.Equals(byId.InternalName, internalName, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(byId.TypeAsString, type, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("Target List migration provenance field collides with a different field: " + internalName + ".");
                }
                return;
            }
            if (byName != null)
            {
                throw new InvalidDataException("Target List migration provenance field name is already used by a different GUID: " + internalName + ".");
            }
            var created = list.Fields.AddFieldAsXml(
                schemaXml,
                false,
                AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToNoContentType);
            context.Load(created, value => value.Id, value => value.InternalName, value => value.TypeAsString);
            context.ExecuteQueryRetry();
            if (created.Id != id
                || !string.Equals(created.InternalName, internalName, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(created.TypeAsString, type, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Fresh target List migration provenance field differs: " + internalName + ".");
            }
        }

        private static IDictionary<int, ListItem> ReadExisting(ClientContext context, List list)
        {
            var result = new Dictionary<int, ListItem>();
            ListItemCollectionPosition position = null;
            do
            {
                var query = new CamlQuery
                {
                    ViewXml = "<View Scope='RecursiveAll'><ViewFields><FieldRef Name='ID'/><FieldRef Name='" + OriginalItemIdFieldName
                        + "'/><FieldRef Name='" + OriginalItemDigestFieldName + "'/></ViewFields><RowLimit Paged='TRUE'>2000</RowLimit></View>",
                    ListItemCollectionPosition = position
                };
                var items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQueryRetry();
                foreach (var item in items)
                {
                    object value;
                    int sourceId;
                    if (item.FieldValues.TryGetValue(OriginalItemIdFieldName, out value)
                        && int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out sourceId))
                    {
                        if (result.ContainsKey(sourceId))
                        {
                            throw new InvalidDataException("Target List contains duplicate PnP source item identity " + sourceId + ".");
                        }
                        result[sourceId] = item;
                    }
                }
                position = items.ListItemCollectionPosition;
            }
            while (position != null);
            return result;
        }

        private static IEnumerable<ListItemSnapshot> OrderItems(IEnumerable<ListItemSnapshot> items)
        {
            return items.OrderBy(value => value.Document == null ? 2 : value.Document.Kind == ListDocumentObjectKind.Folder ? 0 : 1)
                .ThenBy(value => value.Document == null ? 0 : (value.Document.ServerRelativeUrl ?? string.Empty).Count(character => character == '/'))
                .ThenBy(value => value.SourceItemId);
        }

        internal static string ComputeItemDigest(ListItemSnapshot item)
        {
            return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(
                new ListItemCoreDigestContract
                {
                    SourceItemId = item.SourceItemId,
                    SourceUniqueId = item.SourceUniqueId,
                    Values = item.Values.Where(value => value != null).ToList()
                }));
        }

        private static string ReadString(ListItem item, string field)
        {
            object value;
            return item.FieldValues.TryGetValue(field, out value) ? Convert.ToString(value, CultureInfo.InvariantCulture) : null;
        }

        private sealed class ListItemCoreDigestContract
        {
            public int SourceItemId { get; set; }

            public Guid? SourceUniqueId { get; set; }

            public IList<ListItemValueSnapshot> Values { get; set; } = new List<ListItemValueSnapshot>();
        }
    }
}
