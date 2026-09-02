using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.Ingredients;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageListContentIngredientActionProjector
    {
        public static void Project(
            ListDependencySnapshot source,
            ListMaterializationPlan listPlan,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions)
        {
            var fieldPlans = listPlan.Fields.ToDictionary(value => value.InternalName, StringComparer.OrdinalIgnoreCase);
            foreach (var item in source.Items.Where(value => value != null))
            {
                AddItem(source, listPlan, item, fieldPlans, listBlocked, actions);
                if (item.Document != null)
                {
                    AddDocument(source, listPlan, item, listBlocked, actions);
                }
                foreach (var attachment in item.Attachments.Where(value => value != null))
                {
                    AddAttachment(source, listPlan, item, attachment, listBlocked, actions);
                }
            }
        }

        private static void AddItem(
            ListDependencySnapshot source,
            ListMaterializationPlan listPlan,
            ListItemSnapshot item,
            IDictionary<string, ListFieldMaterializationPlan> fieldPlans,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions)
        {
            var mapping = MapItem(item, fieldPlans, listBlocked);
            var targetIdentity = listPlan.TargetRootFolderServerRelativeUrl + "#source-item:" + item.SourceItemId;
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.ListItem(source.SourceWebId, source.SourceListId, item.SourceItemId),
                mapping.Capability,
                mapping.Disposition,
                mapping.Realization,
                "policy.list-item.current-state",
                mapping.Reason,
                mapping.Disposition == IngredientDisposition.Block ? null : targetIdentity,
                mapping.Disposition == IngredientDisposition.Block
                    ? null
                    : $"The List receipt contains a source-to-target item ID mapping for source item '{item.SourceItemId}'.",
                mapping.Disposition == IngredientDisposition.Block
                    ? null
                    : "Fresh readback verifies every approved value and the item provenance digest."));
        }

        private static void AddDocument(
            ListDependencySnapshot source,
            ListMaterializationPlan listPlan,
            ListItemSnapshot item,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions)
        {
            var document = item.Document;
            var binaryUnavailable = document.Kind == ListDocumentObjectKind.File
                && (document.Content == null
                    || document.Content.Availability != EvidenceAvailability.Captured
                    || document.Content.Artifact == null);
            var blocked = listBlocked || binaryUnavailable;
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.ListDocument(source.SourceWebId, source.SourceListId, item.SourceItemId),
                blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                blocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                blocked
                    ? "none"
                    : document.Kind == ListDocumentObjectKind.Folder ? "create-or-reuse-folder" : "copy-exact-bytes-create-only",
                "policy.list-document.current-state",
                blocked
                    ? "The document object cannot be replayed because its owning List is blocked or exact file bytes are unavailable."
                    : "Materialize the captured current document or folder object under the target-owned List path.",
                blocked ? null : MapListOwnedPath(source, listPlan, document.ServerRelativeUrl),
                blocked
                    ? null
                    : document.Kind == ListDocumentObjectKind.Folder
                        ? "Fresh readback verifies the target folder path and item provenance."
                        : $"Fresh readback verifies target file bytes with SHA-256 '{document.Content?.Artifact?.Sha256}'."));
        }

        private static void AddAttachment(
            ListDependencySnapshot source,
            ListMaterializationPlan listPlan,
            ListItemSnapshot item,
            ListAttachmentSnapshot attachment,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions)
        {
            var blocked = listBlocked
                || attachment.Content == null
                || attachment.Content.Availability != EvidenceAvailability.Captured
                || attachment.Content.Artifact == null;
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.ListAttachment(source.SourceWebId, source.SourceListId, item.SourceItemId, attachment.FileName),
                blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                blocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                blocked ? "none" : "copy-exact-bytes-create-only",
                "policy.list-attachment.current-state",
                blocked
                    ? "The attachment cannot be replayed because its owning List is blocked or exact bytes are unavailable."
                    : "Copy the exact captured attachment bytes to the materialized target item.",
                blocked ? null : listPlan.TargetRootFolderServerRelativeUrl + "#source-item:" + item.SourceItemId + "/attachment:" + attachment.FileName,
                blocked ? null : $"Fresh readback verifies attachment bytes with SHA-256 '{attachment.Content?.Artifact?.Sha256}'."));
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization, string Reason) MapItem(
            ListItemSnapshot item,
            IDictionary<string, ListFieldMaterializationPlan> fieldPlans,
            bool listBlocked)
        {
            if (listBlocked)
            {
                return (IngredientCapability.Incompatible, IngredientDisposition.Block, "none", "The owning List has no executable materialization plan.");
            }

            var transformed = false;
            foreach (var value in item.Values.Where(value => value != null && value.Kind != ListItemValueKind.Null))
            {
                if (string.Equals(value.InternalName, "ContentTypeId", StringComparison.OrdinalIgnoreCase))
                {
                    transformed = true;
                    continue;
                }
                if (!fieldPlans.TryGetValue(value.InternalName, out var fieldPlan)
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.Block
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.EvidenceOnly)
                {
                    return (
                        IngredientCapability.Incompatible,
                        IngredientDisposition.Block,
                        "none",
                        $"Captured value '{value.InternalName}' has no approved replay or substitution action.");
                }
                if (fieldPlan.Disposition == ListFieldMaterializationDisposition.MapLookup
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.MapTaxonomy
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.RequireTargetRuntime
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.CreateOrReuseOwnedCalculated
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.CreateOrReuseOwnedSchemaOnly)
                {
                    transformed = true;
                }
            }

            return transformed
                ? (IngredientCapability.Available, IngredientDisposition.Transform, "replay-approved-values-and-substitute-runtime-values",
                    "Replay recognized business values while remapping identity-bound values and allowing reviewed target-runtime values to be regenerated.")
                : (IngredientCapability.Available, IngredientDisposition.Preserve, "replay-approved-current-values",
                    "Replay every nonempty captured value through an approved lossless field action.");
        }

        private static string MapListOwnedPath(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            string sourcePath)
        {
            if (!string.IsNullOrWhiteSpace(sourcePath)
                && !string.IsNullOrWhiteSpace(source.RootFolderServerRelativeUrl)
                && sourcePath.StartsWith(source.RootFolderServerRelativeUrl.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase))
            {
                return plan.TargetRootFolderServerRelativeUrl.TrimEnd('/')
                    + sourcePath.Substring(source.RootFolderServerRelativeUrl.TrimEnd('/').Length);
            }
            return plan.TargetRootFolderServerRelativeUrl;
        }
    }
}
