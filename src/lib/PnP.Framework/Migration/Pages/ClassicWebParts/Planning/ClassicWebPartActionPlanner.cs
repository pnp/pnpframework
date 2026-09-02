using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Planning
{
    internal static class ClassicWebPartActionPlanner
    {
        public static IList<ClassicWebPartAction> Build(
            IEnumerable<ClassicWebPartSnapshot> webParts,
            IEnumerable<ClassicListWebPartBindingSnapshot> bindings,
            ListMigrationPlanSet listMigration,
            ICollection<string> blockers)
        {
            var bindingByWebPart = (bindings ?? Array.Empty<ClassicListWebPartBindingSnapshot>())
                .ToDictionary(value => value.SourceWebPartId);
            var listPlans = listMigration == null
                ? new Dictionary<Guid, ListMaterializationPlan>()
                : listMigration.Lists.ToDictionary(value => value.SourceListId);
            var actions = new List<ClassicWebPartAction>();
            foreach (var webPart in webParts ?? Array.Empty<ClassicWebPartSnapshot>())
            {
                var portabilityBlocker = ClassicWebPartReplayCapabilityPolicy.GetBlocker(webPart.ExportXml);
                if (!string.IsNullOrWhiteSpace(portabilityBlocker))
                {
                    actions.Add(new ClassicWebPartAction
                    {
                        SourceWebPartId = webPart.Id,
                        Disposition = ClassicWebPartDisposition.Block,
                        Reason = portabilityBlocker
                    });
                    blockers.Add("WebPartCapabilityBlocked: Web Part '" + (webPart.Title ?? webPart.Id.ToString("D")) + "' cannot be preserved: " + portabilityBlocker + ".");
                    continue;
                }

                if (!bindingByWebPart.TryGetValue(webPart.Id, out var binding))
                {
                    actions.Add(new ClassicWebPartAction
                    {
                        SourceWebPartId = webPart.Id,
                        Disposition = ClassicWebPartDisposition.CopyCaptured,
                        Reason = "Copy the portable shared Web Part export after approved text rewrites."
                    });
                    continue;
                }

                if (!listPlans.TryGetValue(binding.SourceListId, out var listPlan) || !listPlan.IsExecutable)
                {
                    actions.Add(Block(webPart, binding, "The bound source List has no executable target materialization plan."));
                    blockers.Add("ListWebPartBindingBlocked: Web Part '" + (webPart.Title ?? webPart.Id.ToString("D")) + "' has no executable target List plan.");
                    continue;
                }

                if (!binding.SourceViewId.HasValue
                    || !listPlan.Views.Any(value => value.SourceViewId == binding.SourceViewId.Value
                        && (value.Disposition == ListViewMaterializationDisposition.CreateOrReuseWebPartView
                            || value.Disposition == ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView)))
                {
                    actions.Add(Block(webPart, binding, "The captured List Web Part View has no executable target View plan."));
                    blockers.Add("ListWebPartViewMappingBlocked: Web Part '" + (webPart.Title ?? webPart.Id.ToString("D")) + "' has no captured executable source View identity.");
                    continue;
                }

                actions.Add(new ClassicWebPartAction
                {
                    SourceWebPartId = webPart.Id,
                    Disposition = ClassicWebPartDisposition.RebindListAfterMaterialization,
                    SourceListWebId = binding.SourceListWebId,
                    SourceListId = binding.SourceListId,
                    SourceViewId = binding.SourceViewId,
                    TargetWebUrl = listPlan.TargetWebUrl,
                    TargetListServerRelativeUrl = listPlan.TargetRootFolderServerRelativeUrl,
                    Reason = "Resolve target Web/List/View runtime IDs from the verified materialization receipt, then rewrite the sealed Web Part export."
                });
            }

            return actions.OrderBy(value => value.SourceWebPartId).ToList();
        }

        private static ClassicWebPartAction Block(
            ClassicWebPartSnapshot webPart,
            ClassicListWebPartBindingSnapshot binding,
            string reason)
        {
            return new ClassicWebPartAction
            {
                SourceWebPartId = webPart.Id,
                Disposition = ClassicWebPartDisposition.Block,
                SourceListWebId = binding.SourceListWebId,
                SourceListId = binding.SourceListId,
                SourceViewId = binding.SourceViewId,
                Reason = reason
            };
        }
    }
}
