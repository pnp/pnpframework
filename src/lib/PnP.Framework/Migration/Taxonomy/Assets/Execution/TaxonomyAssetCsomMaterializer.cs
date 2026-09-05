using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Taxonomy.Assets.Execution
{
    internal static class TaxonomyAssetCsomMaterializer
    {
        public static TermStore GetStore(ClientContext context, Guid targetTermStoreId)
        {
            var store = TaxonomySession.GetTaxonomySession(context).TermStores.GetById(targetTermStoreId);
            context.Load(store, value => value.Id);
            context.ExecuteQueryRetry();
            if (store.ServerObjectIsNull.GetValueOrDefault(true))
            {
                throw new InvalidOperationException("Target taxonomy TermStore is unavailable: " + targetTermStoreId.ToString("D") + ".");
            }
            return store;
        }

        public static bool EnsureOwnedGroup(
            ClientContext context,
            TermStore store,
            Guid groupId,
            string groupName)
        {
            var group = store.GetGroup(groupId);
            context.Load(group, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            if (!group.ServerObjectIsNull.GetValueOrDefault(true))
            {
                if (!string.Equals(group.Name, groupName, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Target taxonomy group GUID is occupied by a different group: " + groupId.ToString("D") + ".");
                }
                return false;
            }

            group = store.CreateGroup(groupName, groupId);
            store.CommitAll();
            context.ExecuteQueryRetry();
            var readback = store.GetGroup(groupId);
            context.Load(readback, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            if (readback.ServerObjectIsNull.GetValueOrDefault(true)
                || readback.Id != groupId
                || !string.Equals(readback.Name, groupName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Fresh taxonomy group readback differs from the reviewed identity.");
            }
            return true;
        }

        public static bool EnsureTermSet(
            ClientContext context,
            TermStore store,
            TaxonomyTermSetMaterializationPlan plan,
            TaxonomyTermSetTargetProbe preflight)
        {
            if (preflight.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned)
            {
                return false;
            }
            if (preflight.Disposition == TaxonomyAssetTargetDisposition.CreateMissing)
            {
                var group = store.GetGroup(plan.TargetGroupId);
                context.Load(group, value => value.Id, value => value.Name);
                context.ExecuteQueryRetry();
                if (group.ServerObjectIsNull.GetValueOrDefault(true)
                    || !string.Equals(group.Name, plan.TargetGroupName, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("The migration-owned taxonomy group is unavailable or differs from the reviewed plan.");
                }
                var created = group.CreateTermSet(plan.TargetTermSetName, plan.PreferredTargetTermSetId, plan.Language);
                created.IsOpenForTermCreation = plan.IsOpenForTermCreation;
                created.IsAvailableForTagging = plan.IsAvailableForTagging;
                created.SetCustomProperty(plan.OriginalIdentifierPropertyName, plan.OriginalIdentifier);
                store.CommitAll();
                context.ExecuteQueryRetry();
                return true;
            }
            if (preflight.Disposition == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift)
            {
                if (!preflight.ResolvedTargetTermSetId.HasValue)
                {
                    throw new InvalidOperationException("The owned TermSet drift probe has no target identity.");
                }
                var existing = store.GetTermSet(preflight.ResolvedTargetTermSetId.Value);
                context.Load(
                    existing,
                    value => value.Id,
                    value => value.Name,
                    value => value.CustomProperties,
                    value => value.IsOpenForTermCreation,
                    value => value.IsAvailableForTagging);
                context.ExecuteQueryRetry();
                AssertOwned(existing.CustomProperties, plan.OriginalIdentifierPropertyName, plan.OriginalIdentifier, "TermSet");
                existing.Name = plan.TargetTermSetName;
                existing.IsOpenForTermCreation = plan.IsOpenForTermCreation;
                existing.IsAvailableForTagging = plan.IsAvailableForTagging;
                existing.SetCustomProperty(plan.OriginalIdentifierPropertyName, plan.OriginalIdentifier);
                store.CommitAll();
                context.ExecuteQueryRetry();
                return true;
            }
            throw new InvalidOperationException("TermSet disposition is not mutable: " + preflight.Disposition + ".");
        }

        public static bool EnsureTerm(
            ClientContext context,
            TermStore store,
            TaxonomyTermMaterializationPlan plan,
            TaxonomyTermTargetProbe preflight)
        {
            if (preflight.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned)
            {
                return false;
            }
            if (preflight.Disposition == TaxonomyAssetTargetDisposition.CreateMissing
                || preflight.Disposition == TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval)
            {
                var targetSet = RequireTermSet(context, store, plan.TargetTermSetId);
                var parent = ResolveParent(context, store, targetSet, plan.TargetTermSetId, plan.TargetParentTermId);
                var created = parent.CreateTerm(plan.Name, plan.Language, plan.PreferredTargetTermId);
                created.IsAvailableForTagging = plan.IsAvailableForTagging;
                created.SetCustomProperty(plan.OriginalIdentifierPropertyName, plan.OriginalIdentifier);
                store.CommitAll();
                context.ExecuteQueryRetry();
                return true;
            }
            if (preflight.Disposition == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift)
            {
                if (!preflight.ResolvedTargetTermId.HasValue)
                {
                    throw new InvalidOperationException("The owned Term drift probe has no target identity.");
                }
                var targetSet = RequireTermSet(context, store, plan.TargetTermSetId);
                var existing = store.GetTermInTermSet(plan.TargetTermSetId, preflight.ResolvedTargetTermId.Value);
                context.Load(
                    existing,
                    value => value.Id,
                    value => value.Name,
                    value => value.CustomProperties,
                    value => value.IsAvailableForTagging,
                    value => value.Parent,
                    value => value.TermSet);
                context.ExecuteQueryRetry();
                if (existing.ServerObjectIsNull.GetValueOrDefault(true))
                {
                    throw new InvalidOperationException("The provenance-owned target Term disappeared after admission.");
                }
                AssertOwned(existing.CustomProperties, plan.OriginalIdentifierPropertyName, plan.OriginalIdentifier, "Term");
                context.Load(existing.TermSet, value => value.Id);
                if (existing.Parent != null && !existing.Parent.ServerObjectIsNull.GetValueOrDefault(true))
                {
                    context.Load(existing.Parent, value => value.Id);
                }
                context.ExecuteQueryRetry();
                if (existing.TermSet.Id != plan.TargetTermSetId)
                {
                    throw new InvalidOperationException("The provenance-owned Term moved outside the reviewed TermSet.");
                }
                var observedParentId = existing.Parent != null
                    && !existing.Parent.ServerObjectIsNull.GetValueOrDefault(true)
                    && existing.Parent.Id != Guid.Empty
                        ? existing.Parent.Id
                        : (Guid?)null;
                if (observedParentId != plan.TargetParentTermId)
                {
                    existing.Move(ResolveParent(context, store, targetSet, plan.TargetTermSetId, plan.TargetParentTermId));
                }
                existing.Name = plan.Name;
                existing.IsAvailableForTagging = plan.IsAvailableForTagging;
                existing.SetCustomProperty(plan.OriginalIdentifierPropertyName, plan.OriginalIdentifier);
                store.CommitAll();
                context.ExecuteQueryRetry();
                return true;
            }
            throw new InvalidOperationException("Term disposition is not mutable: " + preflight.Disposition + ".");
        }

        private static TermSet RequireTermSet(ClientContext context, TermStore store, Guid termSetId)
        {
            var set = store.GetTermSet(termSetId);
            context.Load(set, value => value.Id);
            context.ExecuteQueryRetry();
            if (set.ServerObjectIsNull.GetValueOrDefault(true))
            {
                throw new InvalidOperationException("Target TermSet is missing: " + termSetId.ToString("D") + ".");
            }
            return set;
        }

        private static TermSetItem ResolveParent(
            ClientContext context,
            TermStore store,
            TermSet targetSet,
            Guid targetTermSetId,
            Guid? parentTermId)
        {
            if (!parentTermId.HasValue)
            {
                return targetSet;
            }
            var parent = store.GetTermInTermSet(targetTermSetId, parentTermId.Value);
            context.Load(parent, value => value.Id, value => value.TermSet);
            context.ExecuteQueryRetry();
            if (parent.ServerObjectIsNull.GetValueOrDefault(true))
            {
                throw new InvalidOperationException("Target parent Term is missing: " + parentTermId.Value.ToString("D") + ".");
            }
            context.Load(parent.TermSet, value => value.Id);
            context.ExecuteQueryRetry();
            if (parent.TermSet.Id != targetTermSetId)
            {
                throw new InvalidOperationException("Target parent Term belongs to a different TermSet.");
            }
            return parent;
        }

        private static void AssertOwned(
            IDictionary<string, string> properties,
            string propertyName,
            string expectedIdentity,
            string objectKind)
        {
            if (properties == null
                || !properties.TryGetValue(propertyName, out var actual)
                || !string.Equals(actual, expectedIdentity, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(objectKind + " ownership provenance changed after admission.");
            }
        }
    }
}
