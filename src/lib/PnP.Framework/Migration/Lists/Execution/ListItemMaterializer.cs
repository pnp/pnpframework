using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
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
        private const string OriginalItemIdFieldName = "PnPMigrationOriginalItemId";
        private const string OriginalItemDigestFieldName = "PnPMigrationOriginalItemDigest";

        public static IDictionary<int, int> Ensure(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            IMigrationArtifactStore artifactStore)
        {
            EnsureReservedFields(context, targetList);
            context.Load(targetList.RootFolder, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var existingBySourceId = ReadExisting(context, targetList);
            var targetItems = new Dictionary<int, ListItem>();
            foreach (var sourceItem in OrderItems(source.Items))
            {
                var expectedDigest = ItemDigest(sourceItem);
                ListItem targetItem;
                if (existingBySourceId.TryGetValue(sourceItem.SourceItemId, out targetItem))
                {
                    var observed = ReadString(targetItem, OriginalItemDigestFieldName);
                    if (!string.IsNullOrWhiteSpace(observed) && !string.Equals(observed, expectedDigest, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException("Target List item provenance collision for source item " + sourceItem.SourceItemId + ".");
                    }
                }
                else
                {
                    targetItem = CreateItem(context, targetList, source, plan, sourceItem, artifactStore);
                }
                targetItems[sourceItem.SourceItemId] = targetItem;
                ApplyRecognizedValues(context, targetList, targetItem, sourceItem, plan, dependencyReceipts, false);
                EnsureAttachments(context, targetItem, sourceItem.Attachments, artifactStore);
            }

            foreach (var sourceItem in source.Items.OrderBy(value => value.SourceItemId))
            {
                var targetItem = targetItems[sourceItem.SourceItemId];
                ApplyRecognizedValues(context, targetList, targetItem, sourceItem, plan, dependencyReceipts, true);
                targetItem[OriginalItemIdFieldName] = sourceItem.SourceItemId;
                targetItem[OriginalItemDigestFieldName] = ItemDigest(sourceItem);
                targetItem.Update();
                context.ExecuteQueryRetry();
            }
            return targetItems.ToDictionary(value => value.Key, value => value.Value.Id);
        }

        private static void EnsureReservedFields(ClientContext context, List list)
        {
            context.Load(list.Fields, values => values.Include(value => value.Id, value => value.InternalName));
            context.ExecuteQueryRetry();
            if (!list.Fields.AsEnumerable().Any(value => value.Id == OriginalItemIdFieldId))
            {
                list.Fields.AddFieldAsXml(
                    "<Field ID=\"{" + OriginalItemIdFieldId.ToString("D") + "}\" Name=\"" + OriginalItemIdFieldName + "\" DisplayName=\"PnP migration source item ID\" Type=\"Integer\" Hidden=\"TRUE\" Required=\"FALSE\" />",
                    true,
                    AddFieldOptions.AddFieldInternalNameHint);
                context.ExecuteQueryRetry();
            }
            if (!list.Fields.AsEnumerable().Any(value => value.Id == OriginalItemDigestFieldId))
            {
                list.Fields.AddFieldAsXml(
                    "<Field ID=\"{" + OriginalItemDigestFieldId.ToString("D") + "}\" Name=\"" + OriginalItemDigestFieldName + "\" DisplayName=\"PnP migration source item digest\" Type=\"Text\" Hidden=\"TRUE\" Required=\"FALSE\" MaxLength=\"64\" />",
                    true,
                    AddFieldOptions.AddFieldInternalNameHint);
                context.ExecuteQueryRetry();
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

        private static ListItem CreateItem(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListItemSnapshot sourceItem,
            IMigrationArtifactStore artifactStore)
        {
            if (sourceItem.Document == null)
            {
                var item = targetList.AddItem(new ListItemCreationInformation());
                item[OriginalItemIdFieldName] = sourceItem.SourceItemId;
                item.Update();
                context.ExecuteQueryRetry();
                context.Load(item, value => value.Id);
                context.ExecuteQueryRetry();
                return item;
            }

            var targetPath = MapDocumentPath(sourceItem.Document.ServerRelativeUrl, source.RootFolderServerRelativeUrl, plan.TargetRootFolderServerRelativeUrl);
            if (sourceItem.Document.Kind == ListDocumentObjectKind.Folder)
            {
                var relative = targetPath.Substring(plan.TargetRootFolderServerRelativeUrl.TrimEnd('/').Length).Trim('/');
                var folder = context.Web.EnsureFolder(targetList.RootFolder, relative, value => value.ServerRelativeUrl, value => value.ListItemAllFields);
                var item = folder.ListItemAllFields;
                context.Load(item, value => value.Id);
                context.ExecuteQueryRetry();
                item[OriginalItemIdFieldName] = sourceItem.SourceItemId;
                item.Update();
                context.ExecuteQueryRetry();
                return item;
            }

            var bytes = MigrationArtifact.ReadAllBytes(sourceItem.Document.Content.Artifact, sourceItem.Document.Content.ContentBase64, artifactStore);
            var directory = targetPath.Substring(0, targetPath.LastIndexOf('/'));
            var relativeDirectory = directory.Substring(plan.TargetRootFolderServerRelativeUrl.TrimEnd('/').Length).Trim('/');
            var folderTarget = string.IsNullOrEmpty(relativeDirectory)
                ? targetList.RootFolder
                : context.Web.EnsureFolder(targetList.RootFolder, relativeDirectory, value => value.ServerRelativeUrl);
            Microsoft.SharePoint.Client.File file;
            if (TryGetFile(context, targetPath, out file))
            {
                VerifyTargetFile(context, file, sourceItem.Document.Content.Artifact);
            }
            else
            {
                using (var stream = new MemoryStream(bytes, false))
                {
                    file = folderTarget.UploadFile(sourceItem.Document.Name, stream, false);
                }
            }
            var listItem = file.ListItemAllFields;
            context.Load(listItem, value => value.Id);
            context.ExecuteQueryRetry();
            listItem[OriginalItemIdFieldName] = sourceItem.SourceItemId;
            listItem.Update();
            context.ExecuteQueryRetry();
            return listItem;
        }

        private static void ApplyRecognizedValues(
            ClientContext context,
            List targetList,
            ListItem targetItem,
            ListItemSnapshot sourceItem,
            ListMaterializationPlan plan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            bool lookupPhase)
        {
            var plans = plan.Fields.ToDictionary(value => value.InternalName, StringComparer.OrdinalIgnoreCase);
            var changed = false;
            foreach (var value in sourceItem.Values)
            {
                ListFieldMaterializationPlan fieldPlan;
                if (!plans.TryGetValue(value.InternalName, out fieldPlan) || value.Kind == ListItemValueKind.Null)
                {
                    continue;
                }
                var isLookup = fieldPlan.Disposition == ListFieldMaterializationDisposition.MapLookup;
                if (isLookup != lookupPhase)
                {
                    continue;
                }
                if (fieldPlan.Disposition != ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue
                    && fieldPlan.Disposition != ListFieldMaterializationDisposition.CreateOrReuseOwnedAndCopyValue
                    && fieldPlan.Disposition != ListFieldMaterializationDisposition.MapLookup
                    && fieldPlan.Disposition != ListFieldMaterializationDisposition.MapTaxonomy)
                {
                    continue;
                }
                if (fieldPlan.Disposition == ListFieldMaterializationDisposition.MapTaxonomy)
                {
                    var field = context.CastTo<TaxonomyField>(targetList.Fields.GetById(fieldPlan.SourceFieldId));
                    field.SetFieldValueByLabelGuidPair(targetItem, string.Join(";", value.TaxonomyValues.Select(term => term.Label + "|" + term.TermGuid)));
                    continue;
                }
                targetItem[fieldPlan.InternalName] = ToTargetValue(value, fieldPlan, dependencyReceipts);
                changed = true;
            }
            if (changed)
            {
                targetItem.Update();
                context.ExecuteQueryRetry();
            }
        }

        private static object ToTargetValue(
            ListItemValueSnapshot value,
            ListFieldMaterializationPlan fieldPlan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts)
        {
            switch (value.Kind)
            {
                case ListItemValueKind.String: return value.ScalarValue;
                case ListItemValueKind.StringCollection: return value.StringValues.ToArray();
                case ListItemValueKind.Boolean: return string.Equals(value.ScalarValue, "true", StringComparison.OrdinalIgnoreCase);
                case ListItemValueKind.Number:
                    decimal number;
                    return decimal.TryParse(value.ScalarValue, NumberStyles.Any, CultureInfo.InvariantCulture, out number) ? (object)number : value.ScalarValue;
                case ListItemValueKind.DateTime:
                    DateTime date;
                    return DateTime.TryParse(value.ScalarValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date) ? (object)date : value.ScalarValue;
                case ListItemValueKind.Guid:
                    Guid guid;
                    return Guid.TryParse(value.ScalarValue, out guid) ? (object)guid : value.ScalarValue;
                case ListItemValueKind.Url:
                    return new FieldUrlValue { Url = value.UrlValue == null ? null : value.UrlValue.Url, Description = value.UrlValue == null ? null : value.UrlValue.Description };
                case ListItemValueKind.Lookup:
                case ListItemValueKind.LookupCollection:
                    ListMaterializationReceipt receipt;
                    if (!fieldPlan.SourceLookupListId.HasValue || !dependencyReceipts.TryGetValue(fieldPlan.SourceLookupListId.Value, out receipt))
                    {
                        throw new InvalidDataException("Lookup value has no target item-ID catalog for field '" + fieldPlan.InternalName + "'.");
                    }
                    var mapped = value.LookupValues.Select(source =>
                    {
                        int targetId;
                        if (!receipt.TargetItemIds.TryGetValue(source.LookupId, out targetId))
                        {
                            throw new InvalidDataException("Lookup source item " + source.LookupId + " has no target item mapping for field '" + fieldPlan.InternalName + "'.");
                        }
                        return new FieldLookupValue { LookupId = targetId };
                    }).ToArray();
                    return value.Kind == ListItemValueKind.Lookup ? (object)mapped.Single() : mapped;
                default:
                    throw new InvalidDataException("Value kind '" + value.Kind + "' is not approved for List item replay.");
            }
        }

        private static void EnsureAttachments(ClientContext context, ListItem item, IEnumerable<ListAttachmentSnapshot> attachments, IMigrationArtifactStore artifactStore)
        {
            var values = attachments.ToArray();
            if (values.Length == 0)
            {
                return;
            }
            context.Load(item.AttachmentFiles, files => files.Include(value => value.FileName, value => value.ServerRelativeUrl));
            context.ExecuteQueryRetry();
            var existing = item.AttachmentFiles.ToDictionary(value => value.FileName, StringComparer.OrdinalIgnoreCase);
            foreach (var attachment in values)
            {
                Attachment existingAttachment;
                if (existing.TryGetValue(attachment.FileName, out existingAttachment))
                {
                    Microsoft.SharePoint.Client.File existingFile;
                    if (!TryGetFile(context, existingAttachment.ServerRelativeUrl, out existingFile))
                    {
                        throw new InvalidDataException("Existing target attachment could not be opened for verification: " + existingAttachment.ServerRelativeUrl);
                    }
                    VerifyTargetFile(context, existingFile, attachment.Content.Artifact);
                    continue;
                }
                var bytes = MigrationArtifact.ReadAllBytes(attachment.Content.Artifact, attachment.Content.ContentBase64, artifactStore);
                using (var stream = new MemoryStream(bytes, false))
                {
                    item.AttachmentFiles.Add(new AttachmentCreationInformation { FileName = attachment.FileName, ContentStream = stream });
                    context.ExecuteQueryRetry();
                }
            }
        }

        private static IEnumerable<ListItemSnapshot> OrderItems(IEnumerable<ListItemSnapshot> items)
        {
            return items.OrderBy(value => value.Document == null ? 2 : value.Document.Kind == ListDocumentObjectKind.Folder ? 0 : 1)
                .ThenBy(value => value.Document == null ? 0 : (value.Document.ServerRelativeUrl ?? string.Empty).Count(character => character == '/'))
                .ThenBy(value => value.SourceItemId);
        }

        private static string ItemDigest(ListItemSnapshot item)
        {
            return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(item));
        }

        private static string MapDocumentPath(string sourcePath, string sourceRoot, string targetRoot)
        {
            if (!sourcePath.StartsWith(sourceRoot.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Document path is outside its captured List root: " + sourcePath);
            }
            return targetRoot.TrimEnd('/') + sourcePath.Substring(sourceRoot.TrimEnd('/').Length);
        }

        private static bool TryGetFile(ClientContext context, string path, out Microsoft.SharePoint.Client.File file)
        {
            file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(path));
            try
            {
                context.Load(file, value => value.Exists, value => value.Length);
                context.ExecuteQueryRetry();
                return file.Exists;
            }
            catch (ServerException)
            {
                return false;
            }
        }

        private static void VerifyTargetFile(ClientContext context, Microsoft.SharePoint.Client.File file, ArtifactReference expected)
        {
            var stream = file.OpenBinaryStream();
            context.ExecuteQueryRetry();
            using (stream.Value)
            using (var buffer = new MemoryStream())
            {
                stream.Value.CopyTo(buffer);
                var bytes = buffer.ToArray();
                if (bytes.LongLength != expected.Length || !string.Equals(MigrationDigest.ComputeSha256(bytes), expected.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("Existing target document bytes differ from the sealed source artifact: " + file.ServerRelativeUrl);
                }
            }
        }

        private static string ReadString(ListItem item, string field)
        {
            object value;
            return item.FieldValues.TryGetValue(field, out value) ? Convert.ToString(value, CultureInfo.InvariantCulture) : null;
        }
    }
}
