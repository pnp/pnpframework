using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy.Assets.Verification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    /// <summary>
    /// Performs read-only target inspection for a sealed taxonomy asset plan.
    /// Missing assets become CreateMissing actions; foreign or ambiguous assets
    /// remain mitigation/review work and are never reported as authorization stops.
    /// </summary>
    public static class TaxonomyAssetTargetInspector
    {
        public static Guid ResolveSingleOnlineTermStoreId(ClientContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            var stores = TaxonomySession.GetTaxonomySession(context).TermStores;
            stores.RefreshLoad();
            context.Load(stores, values => values.Include(value => value.Id, value => value.IsOnline));
            context.ExecuteQueryRetry();
            var online = stores.AsEnumerable().Where(value => value.IsOnline).ToArray();
            if (online.Length != 1)
            {
                throw new InvalidOperationException(online.Length == 0
                    ? "The target exposes no online taxonomy TermStore."
                    : "The target exposes multiple online taxonomy TermStores; an explicit target store is required.");
            }
            return online[0].Id;
        }

        public static TaxonomyAssetReviewPlan Inspect(
            ClientContext context,
            TaxonomyAssetReviewPlan plan)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            if (plan.TargetTermStoreId == Guid.Empty)
            {
                throw new ArgumentException("The taxonomy plan has no target TermStore identity.", nameof(plan));
            }

            var store = GetStore(context, plan.TargetTermStoreId);
            plan.TermGroupProbes.Clear();
            plan.TermSetProbes.Clear();
            plan.TermProbes.Clear();
            plan.MappingCandidates.Clear();
            plan.Issues = plan.Issues
                .Where(value => value != null && !string.Equals(value.Ingredient, "Taxonomy.Target", StringComparison.Ordinal))
                .ToList();

            var groupProbes = new Dictionary<string, TaxonomyTermGroupTargetProbe>(StringComparer.Ordinal);
            var ownedTermSetsByGroup = new Dictionary<Guid, IReadOnlyList<TermSet>>();
            foreach (var groupPlan in plan.TermGroups
                         .OrderBy(value => value.Source.TenantId)
                         .ThenBy(value => value.Source.TermStoreId))
            {
                var probe = ProbeTermGroup(context, groupPlan, store);
                var key = TaxonomyAssetApprovalFactory.GroupKey(
                    groupPlan.Source.TenantId,
                    groupPlan.Source.TermStoreId);
                groupProbes[key] = probe;
                plan.TermGroupProbes.Add(probe);
                AddTermGroupIssue(plan.Issues, probe);
                if (probe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned)
                {
                    ownedTermSetsByGroup[groupPlan.PreferredTargetGroupId] = ReadOwnedTermSets(
                        context,
                        store,
                        groupPlan.PreferredTargetGroupId);
                }
            }

            var termPlansBySet = plan.Terms
                .Where(value => value != null && value.Source != null)
                .GroupBy(value => SetKey(value.Source.TermStoreId, value.Source.TermSetId), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
            foreach (var termSetPlan in plan.TermSets.OrderBy(value => value.Source.TermStoreId).ThenBy(value => value.Source.TermSetId))
            {
                var groupKey = TaxonomyAssetApprovalFactory.GroupKey(
                    termSetPlan.Source.TenantId,
                    termSetPlan.Source.TermStoreId);
                groupProbes.TryGetValue(groupKey, out var groupProbe);
                var ownedTermSets = groupProbe != null
                    && groupProbe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                    && ownedTermSetsByGroup.TryGetValue(termSetPlan.TargetGroupId, out var groupSets)
                        ? groupSets
                        : Array.Empty<TermSet>();
                var probe = ProbePreferredTermSet(context, termSetPlan, store, ownedTermSets);
                var key = SetKey(termSetPlan.Source.TermStoreId, termSetPlan.Source.TermSetId);
                if (probe.Disposition == TaxonomyAssetTargetDisposition.CreateMissing
                    && termPlansBySet.TryGetValue(key, out var requiredTerms)
                    && requiredTerms.Length > 0)
                {
                    var external = ProbeExternalTermSetCandidate(context, store, termSetPlan, requiredTerms);
                    if (external != null)
                    {
                        probe = external;
                    }
                }
                if (probe.Disposition == TaxonomyAssetTargetDisposition.CreateMissing
                    && (groupProbe == null
                        || groupProbe.Disposition == TaxonomyAssetTargetDisposition.ResolveCollision
                        || groupProbe.Disposition == TaxonomyAssetTargetDisposition.AuthorizationBlocked
                        || groupProbe.Disposition == TaxonomyAssetTargetDisposition.RetryRequired
                        || groupProbe.Disposition == TaxonomyAssetTargetDisposition.TargetInspectionRequired))
                {
                    probe.Disposition = groupProbe == null
                        ? TaxonomyAssetTargetDisposition.RetryRequired
                        : groupProbe.Disposition;
                    probe.AuthorizationEvidence = groupProbe?.AuthorizationEvidence;
                    probe.Issues.Add(Issue(
                        "TargetTaxonomyOwnershipGroupUnavailable",
                        "The deterministic target TermGroup has no usable action yet; resolve or retry the TermGroup before creating this TermSet."));
                }
                plan.TermSetProbes.Add(probe);
                AddTermSetIssue(plan.Issues, probe);
                if (probe.ResolvedTargetTermSetId.HasValue
                    || probe.Disposition == TaxonomyAssetTargetDisposition.CreateMissing)
                {
                    var targetSetId = probe.ResolvedTargetTermSetId ?? termSetPlan.PreferredTargetTermSetId;
                    plan.MappingCandidates.Add(new TaxonomyAssetMappingCandidate
                    {
                        SourceTermStoreId = termSetPlan.Source.TermStoreId,
                        SourceTermSetId = termSetPlan.Source.TermSetId,
                        TargetTermStoreId = termSetPlan.TargetTermStoreId,
                        TargetTermSetId = targetSetId,
                        Disposition = probe.Disposition,
                        RequiresReview = probe.Disposition != TaxonomyAssetTargetDisposition.ReuseOwned,
                        EvidenceSha256 = ProbeDigest(probe),
                        VerificationAssertions = VerificationAssertions(probe.Disposition, targetSetId)
                    });
                }
            }

            var setProbeByKey = plan.TermSetProbes.ToDictionary(
                value => SetKey(value.SourceTermStoreId, value.SourceTermSetId),
                StringComparer.Ordinal);
            foreach (var sourceTermPlan in OrderTerms(plan.Terms))
            {
                var setKey = SetKey(sourceTermPlan.Source.TermStoreId, sourceTermPlan.Source.TermSetId);
                if (!setProbeByKey.TryGetValue(setKey, out var setProbe)
                    || setProbe.Disposition == TaxonomyAssetTargetDisposition.ResolveCollision
                    || setProbe.Disposition == TaxonomyAssetTargetDisposition.AuthorizationBlocked
                    || setProbe.Disposition == TaxonomyAssetTargetDisposition.RetryRequired)
                {
                    var unresolved = CreateUnresolvedTermProbe(sourceTermPlan, setProbe);
                    plan.TermProbes.Add(unresolved);
                    AddTermIssue(plan.Issues, unresolved);
                    continue;
                }

                var targetSetId = setProbe.ResolvedTargetTermSetId
                    ?? plan.MappingCandidates.Single(value => value.SourceTermStoreId == sourceTermPlan.Source.TermStoreId
                        && value.SourceTermSetId == sourceTermPlan.Source.TermSetId).TargetTermSetId;
                sourceTermPlan.TargetTermSetId = targetSetId;
                sourceTermPlan.PlanDigest = TaxonomyAssetIdentity.ComputePlanDigest(sourceTermPlan);

                TaxonomyTermTargetProbe termProbe;
                if (setProbe.Disposition == TaxonomyAssetTargetDisposition.CreateMissing)
                {
                    termProbe = ProbeTermForMissingSet(context, store, sourceTermPlan);
                }
                else
                {
                    termProbe = ProbeTerm(context, store, sourceTermPlan, setProbe.Disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse);
                }
                plan.TermProbes.Add(termProbe);
                AddTermIssue(plan.Issues, termProbe);
            }

            plan.TermGroupProbes = plan.TermGroupProbes
                .OrderBy(value => value.SourceTenantId)
                .ThenBy(value => value.SourceTermStoreId)
                .ToList();
            plan.TermSetProbes = plan.TermSetProbes
                .OrderBy(value => value.SourceTermStoreId)
                .ThenBy(value => value.SourceTermSetId)
                .ToList();
            plan.TermProbes = plan.TermProbes
                .OrderBy(value => value.SourceTermStoreId)
                .ThenBy(value => value.SourceTermSetId)
                .ThenBy(value => value.SourceTermId)
                .ToList();
            plan.MappingCandidates = plan.MappingCandidates
                .OrderBy(value => value.SourceTermStoreId)
                .ThenBy(value => value.SourceTermSetId)
                .ToList();
            foreach (var candidate in plan.MappingCandidates)
            {
                candidate.VerificationAssertions = VerificationAssertions(
                    candidate.Disposition,
                    candidate.TargetTermSetId,
                    plan.Terms.Where(value => value.Source.TermStoreId == candidate.SourceTermStoreId
                        && value.Source.TermSetId == candidate.SourceTermSetId),
                    plan.TermProbes.Where(value => value.SourceTermStoreId == candidate.SourceTermStoreId
                        && value.SourceTermSetId == candidate.SourceTermSetId));
            }
            plan.Issues = plan.Issues
                .OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Subject, StringComparer.Ordinal)
                .ToList();
            plan.PlanDigest = TaxonomyAssetPlanner.ComputeDigest(plan);
            return plan;
        }

        public static TaxonomyTermGroupTargetProbe ProbeTermGroup(
            ClientContext context,
            TaxonomyTermGroupMaterializationPlan plan)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (plan == null || plan.Source == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            return ProbeTermGroup(context, plan, GetStore(context, plan.TargetTermStoreId));
        }

        private static TaxonomyTermGroupTargetProbe ProbeTermGroup(
            ClientContext context,
            TaxonomyTermGroupMaterializationPlan plan,
            TermStore store)
        {
            var group = store.GetGroup(plan.PreferredTargetGroupId);
            context.Load(group, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            var exists = !group.ServerObjectIsNull.GetValueOrDefault(true);
            var result = new TaxonomyTermGroupTargetProbe
            {
                SourceTenantId = plan.Source.TenantId,
                SourceTermStoreId = plan.Source.TermStoreId,
                TargetTermStoreId = plan.TargetTermStoreId,
                PreferredIdExists = exists,
                Disposition = TaxonomyAssetTargetDisposition.TargetInspectionRequired
            };
            if (!exists)
            {
                result.Disposition = TaxonomyAssetTargetDisposition.CreateMissing;
                return result;
            }

            result.ResolvedTargetGroupId = group.Id;
            result.ExistingName = group.Name;
            if (group.Id == plan.PreferredTargetGroupId
                && string.Equals(group.Name, plan.TargetGroupName, StringComparison.Ordinal))
            {
                result.Disposition = TaxonomyAssetTargetDisposition.ReuseOwned;
                return result;
            }

            result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
            result.Issues.Add(Issue(
                "TaxonomyTermGroupIdentityCollision",
                "The deterministic target TermGroup GUID is occupied by a differently named group; do not rename or reuse it without a new reviewed identity."));
            return result;
        }

        public static TaxonomyTermSetTargetProbe ProbePreferredTermSet(
            ClientContext context,
            TaxonomyTermSetMaterializationPlan plan)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (plan == null || plan.Source == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            var store = GetStore(context, plan.TargetTermStoreId);
            return ProbePreferredTermSet(
                context,
                plan,
                store,
                ReadOwnedTermSets(context, store, plan.TargetGroupId));
        }

        private static TaxonomyTermSetTargetProbe ProbePreferredTermSet(
            ClientContext context,
            TaxonomyTermSetMaterializationPlan plan,
            TermStore store,
            IReadOnlyList<TermSet> ownedTermSets)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (plan == null || plan.Source == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var matches = ownedTermSets
                .Where(value => string.Equals(
                    Property(value.CustomProperties, plan.OriginalIdentifierPropertyName),
                    plan.OriginalIdentifier,
                    StringComparison.Ordinal))
                .ToArray();
            var preferred = store.GetTermSet(plan.PreferredTargetTermSetId);
            context.Load(
                preferred,
                value => value.Id,
                value => value.Name,
                value => value.CustomProperties,
                value => value.IsOpenForTermCreation,
                value => value.IsAvailableForTagging);
            context.ExecuteQueryRetry();
            var preferredExists = !preferred.ServerObjectIsNull.GetValueOrDefault(true);
            var ids = matches.Select(value => value.Id).Distinct().OrderBy(value => value).ToArray();
            var result = BaseSetProbe(plan, ids, preferredExists);
            if (ids.Length > 1)
            {
                result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                result.Issues.Add(Issue(
                    "DuplicateTaxonomyTermSetOwnership",
                    "More than one target TermSet carries the same source identity; deduplicate the owned assets and re-probe."));
                return result;
            }
            if (ids.Length == 1)
            {
                var match = matches.Single(value => value.Id == ids[0]);
                PopulateSet(result, match);
                result.ResolvedTargetTermSetId = match.Id;
                result.Disposition = ExactSetShape(match, plan)
                    ? TaxonomyAssetTargetDisposition.ReuseOwned
                    : TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift;
                if (result.Disposition == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift)
                {
                    result.Issues.Add(Issue(
                        "OwnedTaxonomyTermSetPlanDrift",
                        "The provenance-matched TermSet differs in name, open state, or tagging availability; reconcile it to the reviewed plan."));
                }
                return result;
            }
            if (!preferredExists)
            {
                result.Disposition = TaxonomyAssetTargetDisposition.CreateMissing;
                return result;
            }

            PopulateSet(result, preferred);
            result.ResolvedTargetTermSetId = preferred.Id;
            if (ExactExternalSetShape(preferred, plan)
                && string.IsNullOrWhiteSpace(Property(preferred.CustomProperties, plan.OriginalIdentifierPropertyName)))
            {
                result.Disposition = TaxonomyAssetTargetDisposition.ReviewExternalReuse;
                result.Issues.Add(Issue(
                    "ExternalTaxonomyTermSetReviewRequired",
                    "The preferred GUID contains an exact external TermSet without migration provenance; explicit reuse approval is required."));
            }
            else
            {
                result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                result.Issues.Add(Issue(
                    "TaxonomyTermSetIdentityCollision",
                    "The preferred target TermSet GUID is occupied by a non-equivalent or differently owned asset; allocate a reviewed target identity."));
            }
            return result;
        }

        public static TaxonomyTermTargetProbe ProbeTerm(
            ClientContext context,
            TaxonomyTermMaterializationPlan plan,
            bool targetSetRequiresExternalApproval)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (plan == null || plan.Source == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            return ProbeTerm(
                context,
                GetStore(context, plan.TargetTermStoreId),
                plan,
                targetSetRequiresExternalApproval);
        }

        private static TaxonomyTermTargetProbe ProbeTerm(
            ClientContext context,
            TermStore store,
            TaxonomyTermMaterializationPlan plan,
            bool targetSetRequiresExternalApproval)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (plan == null || plan.Source == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var matches = store.GetTermsWithCustomProperty(Match(
                context,
                plan.OriginalIdentifierPropertyName,
                plan.OriginalIdentifier));
            context.Load(matches, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.PathOfTerm,
                value => value.CustomProperties,
                value => value.IsAvailableForTagging,
                value => value.Parent,
                value => value.TermSet));

            var preferred = store.GetTermInTermSet(plan.TargetTermSetId, plan.PreferredTargetTermId);
            context.Load(
                preferred,
                value => value.Id,
                value => value.Name,
                value => value.PathOfTerm,
                value => value.CustomProperties,
                value => value.IsAvailableForTagging,
                value => value.Parent,
                value => value.TermSet);

            var global = store.GetTerm(plan.PreferredTargetTermId);
            context.Load(
                global,
                value => value.Id,
                value => value.Name,
                value => value.PathOfTerm,
                value => value.CustomProperties,
                value => value.IsAvailableForTagging,
                value => value.Parent,
                value => value.TermSet);
            context.ExecuteQueryRetry();
            var preferredExists = !preferred.ServerObjectIsNull.GetValueOrDefault(true);
            var globalExists = !global.ServerObjectIsNull.GetValueOrDefault(true);
            LoadTermRelations(
                context,
                matches.AsEnumerable()
                    .Concat(preferredExists ? new[] { preferred } : Array.Empty<Term>())
                    .Concat(globalExists ? new[] { global } : Array.Empty<Term>()));

            var ownershipMatches = matches.AsEnumerable().ToList();
            if (preferredExists
                && string.Equals(
                    Property(preferred.CustomProperties, plan.OriginalIdentifierPropertyName),
                    plan.OriginalIdentifier,
                    StringComparison.Ordinal)
                && ownershipMatches.All(value => value.Id != preferred.Id))
            {
                ownershipMatches.Add(preferred);
            }
            var ids = ownershipMatches.Select(value => value.Id).Distinct().OrderBy(value => value).ToArray();
            var result = BaseTermProbe(plan, ids, preferredExists);
            if (ids.Length > 1)
            {
                result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                result.Issues.Add(Issue("DuplicateTaxonomyTermOwnership", "More than one target Term carries the same source identity."));
                return result;
            }
            if (ids.Length == 1)
            {
                var match = ownershipMatches.Single(value => value.Id == ids[0]);
                PopulateTerm(result, match);
                result.ResolvedTargetTermId = match.Id;
                if (match.TermSet.Id != plan.TargetTermSetId)
                {
                    result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                    result.Issues.Add(Issue("OwnedTaxonomyTermInWrongSet", "The provenance-matched Term exists outside the reviewed target TermSet."));
                }
                else if (ExactTermShape(match, plan)
                    && TaxonomyTermRelationshipFidelity.Matches(
                        plan,
                        result,
                        plan.TargetTermSetId,
                        out _))
                {
                    result.Disposition = TaxonomyAssetTargetDisposition.ReuseOwned;
                }
                else
                {
                    result.Disposition = TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift;
                    TaxonomyTermRelationshipFidelity.Matches(
                        plan,
                        result,
                        plan.TargetTermSetId,
                        out var relationshipDiagnostic);
                    result.Issues.Add(Issue(
                        "OwnedTaxonomyTermPlanDrift",
                        "The provenance-matched Term differs in name, parent topology, path, tagging availability, or captured relationship shape."
                        + (string.IsNullOrWhiteSpace(relationshipDiagnostic) ? string.Empty : " " + relationshipDiagnostic)));
                }
                return result;
            }
            if (preferredExists)
            {
                PopulateTerm(result, preferred);
                result.ResolvedTargetTermId = preferred.Id;
                var existingIdentity = Property(preferred.CustomProperties, plan.OriginalIdentifierPropertyName);
                var relationshipMatches = TaxonomyTermRelationshipFidelity.Matches(
                    plan,
                    result,
                    plan.TargetTermSetId,
                    out var relationshipDiagnostic);
                if (ExactTermShape(preferred, plan)
                    && relationshipMatches
                    && string.IsNullOrWhiteSpace(existingIdentity))
                {
                    result.Disposition = TaxonomyAssetTargetDisposition.ReviewExternalReuse;
                    result.Issues.Add(Issue("ExternalTaxonomyTermReviewRequired", "The exact Term exists without migration provenance; explicit external reuse approval is required."));
                }
                else
                {
                    result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                    result.Issues.Add(Issue(
                        relationshipMatches
                            ? "TaxonomyTermIdentityCollision"
                            : "ExternalTaxonomyTermRelationshipConflict",
                        "The preferred target Term GUID is occupied by a non-equivalent or differently owned Term."
                        + (string.IsNullOrWhiteSpace(relationshipDiagnostic) ? string.Empty : " " + relationshipDiagnostic)));
                }
                return result;
            }
            if (globalExists)
            {
                PopulateTerm(result, global);
                result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                result.Issues.Add(Issue("TaxonomyTermGlobalIdentityCollision", "The preferred Term GUID already exists in another target TermSet."));
                return result;
            }

            result.Disposition = targetSetRequiresExternalApproval
                ? TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval
                : TaxonomyAssetTargetDisposition.CreateMissing;
            return result;
        }

        private static TaxonomyTermTargetProbe ProbeTermForMissingSet(
            ClientContext context,
            TermStore store,
            TaxonomyTermMaterializationPlan plan)
        {
            var matches = store.GetTermsWithCustomProperty(Match(
                context,
                plan.OriginalIdentifierPropertyName,
                plan.OriginalIdentifier));
            context.Load(matches, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.PathOfTerm,
                value => value.CustomProperties,
                value => value.IsAvailableForTagging,
                value => value.Parent,
                value => value.TermSet));

            var global = store.GetTerm(plan.PreferredTargetTermId);
            context.Load(
                global,
                value => value.Id,
                value => value.Name,
                value => value.PathOfTerm,
                value => value.CustomProperties,
                value => value.IsAvailableForTagging,
                value => value.Parent,
                value => value.TermSet);
            context.ExecuteQueryRetry();

            var ownershipMatches = matches.AsEnumerable().ToArray();
            var globalExists = !global.ServerObjectIsNull.GetValueOrDefault(true);
            LoadTermRelations(
                context,
                ownershipMatches.Concat(globalExists ? new[] { global } : Array.Empty<Term>()));

            var ids = ownershipMatches.Select(value => value.Id).Distinct().OrderBy(value => value).ToArray();
            var result = BaseTermProbe(plan, ids, false);
            if (ids.Length > 1)
            {
                result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                result.Issues.Add(Issue(
                    "DuplicateTaxonomyTermOwnership",
                    "More than one target Term carries the same source identity while the planned target TermSet is absent."));
                return result;
            }
            if (ids.Length == 1)
            {
                var match = ownershipMatches.Single(value => value.Id == ids[0]);
                result.ResolvedTargetTermId = match.Id;
                PopulateTerm(result, match);
                result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                result.Issues.Add(Issue(
                    "OwnedTaxonomyTermInWrongSet",
                    "The provenance-matched Term already exists outside the missing planned target TermSet."));
                return result;
            }
            if (globalExists)
            {
                PopulateTerm(result, global);
                result.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                result.Issues.Add(Issue(
                    "TaxonomyTermGlobalIdentityCollision",
                    "The preferred Term GUID already exists while the planned target TermSet is absent."));
                return result;
            }

            result.Disposition = TaxonomyAssetTargetDisposition.CreateMissing;
            return result;
        }

        private static TaxonomyTermSetTargetProbe ProbeExternalTermSetCandidate(
            ClientContext context,
            TermStore store,
            TaxonomyTermSetMaterializationPlan setPlan,
            IReadOnlyCollection<TaxonomyTermMaterializationPlan> requiredTerms)
        {
            var reads = new List<Term>();
            foreach (var plan in requiredTerms.OrderBy(value => value.Source.TermId))
            {
                var term = store.GetTerm(plan.Source.TermId);
                context.Load(
                    term,
                    value => value.Id,
                    value => value.Name,
                    value => value.PathOfTerm,
                    value => value.CustomProperties,
                    value => value.IsAvailableForTagging,
                    value => value.TermSet);
                reads.Add(term);
            }
            context.ExecuteQueryRetry();
            var existing = reads.Where(value => !value.ServerObjectIsNull.GetValueOrDefault(true)).ToArray();
            if (existing.Length == 0)
            {
                return null;
            }
            LoadTermRelations(context, existing);
            foreach (var term in existing)
            {
                context.Load(
                    term.TermSet,
                    value => value.Id,
                    value => value.Name,
                    value => value.CustomProperties,
                    value => value.IsOpenForTermCreation,
                    value => value.IsAvailableForTagging);
            }
            context.ExecuteQueryRetry();

            var plansById = requiredTerms.ToDictionary(value => value.Source.TermId);
            var exact = new List<Term>();
            var relationshipConflicts = new List<string>();
            foreach (var term in existing)
            {
                var termPlan = plansById[term.Id];
                if (!ExactExternalTerm(term, termPlan))
                {
                    continue;
                }
                var termProbe = BaseTermProbe(termPlan, Array.Empty<Guid>(), true);
                PopulateTerm(termProbe, term);
                if (TaxonomyTermRelationshipFidelity.Matches(
                        termPlan,
                        termProbe,
                        term.TermSet.Id,
                        out var relationshipDiagnostic))
                {
                    exact.Add(term);
                }
                else
                {
                    relationshipConflicts.Add(term.Id.ToString("D") + ": " + relationshipDiagnostic);
                }
            }
            var setIds = exact.Select(value => value.TermSet.Id).Distinct().ToArray();
            if (exact.Count != existing.Length || setIds.Length != 1)
            {
                return new TaxonomyTermSetTargetProbe
                {
                    SourceTermStoreId = setPlan.Source.TermStoreId,
                    SourceTermSetId = setPlan.Source.TermSetId,
                    TargetTermStoreId = setPlan.TargetTermStoreId,
                    Disposition = TaxonomyAssetTargetDisposition.ResolveCollision,
                    ExternalCandidateInferredFromTerms = true,
                    SupportingTargetTermIds = exact.Select(value => value.Id).OrderBy(value => value).ToList(),
                    MissingTargetTermIds = requiredTerms.Select(value => value.Source.TermId).Except(existing.Select(value => value.Id)).OrderBy(value => value).ToList(),
                    Issues = new List<MigrationIssue>
                    {
                        Issue(
                            relationshipConflicts.Count == 0
                                ? "AmbiguousExternalTaxonomyTermSetCandidate"
                                : "ExternalTaxonomyTermRelationshipConflict",
                            "Existing exact Term identities do not resolve to one relationship-equivalent target TermSet."
                            + (relationshipConflicts.Count == 0
                                ? string.Empty
                                : " " + string.Join(" ", relationshipConflicts)))
                    }
                };
            }

            var candidate = exact[0].TermSet;
            var identity = Property(candidate.CustomProperties, setPlan.OriginalIdentifierPropertyName);
            var equivalent = string.Equals(candidate.Name, setPlan.SourceTermSetName, StringComparison.Ordinal)
                && candidate.IsOpenForTermCreation == setPlan.IsOpenForTermCreation
                && candidate.IsAvailableForTagging == setPlan.IsAvailableForTagging;
            var probe = BaseSetProbe(setPlan, Array.Empty<Guid>(), false);
            PopulateSet(probe, candidate);
            probe.ResolvedTargetTermSetId = candidate.Id;
            probe.ExternalCandidateInferredFromTerms = true;
            probe.SupportingTargetTermIds = exact.Select(value => value.Id).OrderBy(value => value).ToList();
            probe.MissingTargetTermIds = requiredTerms.Select(value => value.Source.TermId)
                .Except(existing.Select(value => value.Id))
                .OrderBy(value => value)
                .ToList();
            if (equivalent && string.IsNullOrWhiteSpace(identity))
            {
                probe.Disposition = TaxonomyAssetTargetDisposition.ReviewExternalReuse;
                probe.Issues.Add(Issue(
                    "ExternalTaxonomyTermSetReviewRequired",
                    "Existing exact Term identities infer one equivalent external TermSet; explicit mapping approval is required."));
            }
            else
            {
                probe.Disposition = TaxonomyAssetTargetDisposition.ResolveCollision;
                probe.Issues.Add(Issue(
                    "ExternalTaxonomyTermSetCandidateConflict",
                    "The inferred external TermSet has conflicting shape or ownership provenance."));
            }
            return probe;
        }

        private static TaxonomyTermSetTargetProbe BaseSetProbe(
            TaxonomyTermSetMaterializationPlan plan,
            IEnumerable<Guid> provenanceMatches,
            bool preferredExists)
        {
            return new TaxonomyTermSetTargetProbe
            {
                SourceTermStoreId = plan.Source.TermStoreId,
                SourceTermSetId = plan.Source.TermSetId,
                TargetTermStoreId = plan.TargetTermStoreId,
                ProvenanceMatches = provenanceMatches.OrderBy(value => value).ToList(),
                PreferredIdExists = preferredExists,
                Disposition = TaxonomyAssetTargetDisposition.TargetInspectionRequired
            };
        }

        private static TaxonomyTermTargetProbe BaseTermProbe(
            TaxonomyTermMaterializationPlan plan,
            IEnumerable<Guid> provenanceMatches,
            bool preferredExists)
        {
            return new TaxonomyTermTargetProbe
            {
                SourceTermStoreId = plan.Source.TermStoreId,
                SourceTermSetId = plan.Source.TermSetId,
                SourceTermId = plan.Source.TermId,
                TargetTermStoreId = plan.TargetTermStoreId,
                TargetTermSetId = plan.TargetTermSetId,
                ProvenanceMatches = provenanceMatches.OrderBy(value => value).ToList(),
                PreferredIdExists = preferredExists,
                Disposition = TaxonomyAssetTargetDisposition.TargetInspectionRequired
            };
        }

        private static TaxonomyTermTargetProbe CreateUnresolvedTermProbe(
            TaxonomyTermMaterializationPlan plan,
            TaxonomyTermSetTargetProbe setProbe)
        {
            var result = BaseTermProbe(plan, Array.Empty<Guid>(), false);
            result.Disposition = setProbe == null
                ? TaxonomyAssetTargetDisposition.RetryRequired
                : setProbe.Disposition;
            result.Issues.Add(Issue(
                "TargetTaxonomyTermSetMappingUnavailable",
                "The owning target TermSet has no usable mapping candidate yet."));
            return result;
        }

        private static void PopulateSet(TaxonomyTermSetTargetProbe target, TermSet source)
        {
            target.ExistingName = source.Name;
            target.ExistingOriginalIdentifier = Property(source.CustomProperties, TaxonomyAssetIdentity.OriginalIdentifierPropertyName);
            target.ExistingIsOpenForTermCreation = source.IsOpenForTermCreation;
            target.ExistingIsAvailableForTagging = source.IsAvailableForTagging;
        }

        private static void PopulateTerm(TaxonomyTermTargetProbe target, Term source)
        {
            target.ExistingName = source.Name;
            target.ExistingPath = source.PathOfTerm;
            target.ExistingOriginalIdentifier = Property(source.CustomProperties, TaxonomyAssetIdentity.OriginalIdentifierPropertyName);
            target.ExistingIsAvailableForTagging = source.IsAvailableForTagging;
            target.ExistingTermSetId = source.TermSet.Id;
            target.ExistingIsReused = source.IsReused;
            target.ExistingIsSourceTerm = source.IsSourceTerm;
            target.ExistingReuseSourceTermId = Id(source.SourceTerm);
            target.ExistingTermSetIds = ReadTermSetIds(source);
            target.ExistingPinSourceTermSetId = Id(source.PinSourceTermSet);
            target.ExistingParentTermId = source.Parent != null
                && !source.Parent.ServerObjectIsNull.GetValueOrDefault(true)
                && source.Parent.Id != Guid.Empty
                    ? source.Parent.Id
                    : (Guid?)null;
        }

        private static bool ExactSetShape(TermSet observed, TaxonomyTermSetMaterializationPlan expected)
        {
            return string.Equals(observed.Name, expected.TargetTermSetName, StringComparison.Ordinal)
                && observed.IsOpenForTermCreation == expected.IsOpenForTermCreation
                && observed.IsAvailableForTagging == expected.IsAvailableForTagging;
        }

        private static bool ExactExternalSetShape(TermSet observed, TaxonomyTermSetMaterializationPlan expected)
        {
            return string.Equals(observed.Name, expected.SourceTermSetName, StringComparison.Ordinal)
                && observed.IsOpenForTermCreation == expected.IsOpenForTermCreation
                && observed.IsAvailableForTagging == expected.IsAvailableForTagging;
        }

        private static bool ExactTermShape(Term observed, TaxonomyTermMaterializationPlan expected)
        {
            var parentId = observed.Parent != null
                && !observed.Parent.ServerObjectIsNull.GetValueOrDefault(true)
                && observed.Parent.Id != Guid.Empty
                    ? observed.Parent.Id
                    : (Guid?)null;
            return string.Equals(observed.Name, expected.Name, StringComparison.Ordinal)
                && (string.IsNullOrWhiteSpace(expected.SourcePath)
                    || string.Equals(observed.PathOfTerm, expected.SourcePath, StringComparison.Ordinal))
                && observed.IsAvailableForTagging == expected.IsAvailableForTagging
                && parentId == expected.TargetParentTermId;
        }

        private static bool ExactExternalTerm(Term observed, TaxonomyTermMaterializationPlan expected)
        {
            return string.Equals(observed.Name, expected.Name, StringComparison.Ordinal)
                && (string.IsNullOrWhiteSpace(expected.SourcePath)
                    || string.Equals(observed.PathOfTerm, expected.SourcePath, StringComparison.Ordinal))
                && observed.IsAvailableForTagging == expected.IsAvailableForTagging;
        }

        private static TermStore GetStore(ClientContext context, Guid id)
        {
            var store = TaxonomySession.GetTaxonomySession(context).TermStores.GetById(id);
            context.Load(store, value => value.Id);
            context.ExecuteQueryRetry();
            if (store.ServerObjectIsNull.GetValueOrDefault(true))
            {
                throw new InvalidOperationException("Target taxonomy TermStore is unavailable: " + id.ToString("D"));
            }
            return store;
        }

        private static IReadOnlyList<TermSet> ReadOwnedTermSets(
            ClientContext context,
            TermStore store,
            Guid groupId)
        {
            var group = store.GetGroup(groupId);
            context.Load(group, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            if (group.ServerObjectIsNull.GetValueOrDefault(true))
            {
                return Array.Empty<TermSet>();
            }
            var sets = group.TermSets;
            sets.RefreshLoad();
            context.Load(sets, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.CustomProperties,
                value => value.IsOpenForTermCreation,
                value => value.IsAvailableForTagging));
            context.ExecuteQueryRetry();
            return sets.AsEnumerable().ToArray();
        }

        private static void LoadTermRelations(ClientContext context, IEnumerable<Term> terms)
        {
            var materialized = terms.Where(value => value != null && !value.ServerObjectIsNull.GetValueOrDefault(true)).ToArray();
            foreach (var term in materialized)
            {
                context.Load(
                    term,
                    value => value.IsReused,
                    value => value.IsSourceTerm,
                    value => value.SourceTerm,
                    value => value.PinSourceTermSet,
                    value => value.Parent,
                    value => value.TermSet);
                context.Load(term.TermSet, value => value.Id);
                context.Load(term.TermSets, values => values.Include(value => value.Id));
            }
            if (materialized.Length > 0)
            {
                context.ExecuteQueryRetry();
            }
            foreach (var term in materialized)
            {
                if (Exists(term.Parent))
                {
                    context.Load(term.Parent, value => value.Id);
                }
                if (Exists(term.SourceTerm))
                {
                    context.Load(term.SourceTerm, value => value.Id);
                }
                if (Exists(term.PinSourceTermSet))
                {
                    context.Load(term.PinSourceTermSet, value => value.Id);
                }
            }
            if (materialized.Length > 0)
            {
                context.ExecuteQueryRetry();
            }
        }

        private static bool Exists(ClientObject value)
        {
            return value != null && !value.ServerObjectIsNull.GetValueOrDefault(true);
        }

        private static Guid? Id(Term value)
        {
            return Exists(value) && value.Id != Guid.Empty ? value.Id : (Guid?)null;
        }

        private static Guid? Id(TermSet value)
        {
            return Exists(value) && value.Id != Guid.Empty ? value.Id : (Guid?)null;
        }

        private static IList<Guid> ReadTermSetIds(Term term)
        {
            var ids = new List<Guid>();
            var owningSetId = Id(term?.TermSet);
            if (owningSetId.HasValue)
            {
                ids.Add(owningSetId.Value);
            }
            if (term?.TermSets != null)
            {
                ids.AddRange(term.TermSets
                    .Where(value => Exists(value) && value.Id != Guid.Empty)
                    .Select(value => value.Id));
            }
            return ids.Distinct().OrderBy(value => value).ToList();
        }

        private static CustomPropertyMatchInformation Match(
            ClientContext context,
            string propertyName,
            string propertyValue)
        {
            return new CustomPropertyMatchInformation(context)
            {
                CustomPropertyName = propertyName,
                CustomPropertyValue = propertyValue,
                StringMatchOption = StringMatchOption.ExactMatch,
                ResultCollectionSize = 10,
                TrimUnavailable = true
            };
        }

        private static IEnumerable<TaxonomyTermMaterializationPlan> OrderTerms(
            IEnumerable<TaxonomyTermMaterializationPlan> terms)
        {
            var remaining = (terms ?? Enumerable.Empty<TaxonomyTermMaterializationPlan>())
                .Where(value => value != null && value.Source != null)
                .GroupBy(value => TermKey(value.Source.TermStoreId, value.Source.TermSetId, value.Source.TermId), StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
            var emitted = new HashSet<string>(StringComparer.Ordinal);
            while (remaining.Count > 0)
            {
                var ready = remaining
                    .Where(value => !value.TargetParentTermId.HasValue
                        || emitted.Contains(TermKey(
                            value.Source.TermStoreId,
                            value.Source.TermSetId,
                            value.TargetParentTermId.Value)))
                    .OrderBy(value => value.Source.TermStoreId)
                    .ThenBy(value => value.Source.TermSetId)
                    .ThenBy(value => value.Source.TermId)
                    .ToArray();
                if (ready.Length == 0)
                {
                    foreach (var unresolved in remaining
                                 .OrderBy(value => value.Source.TermStoreId)
                                 .ThenBy(value => value.Source.TermSetId)
                                 .ThenBy(value => value.Source.TermId))
                    {
                        yield return unresolved;
                    }
                    yield break;
                }
                foreach (var item in ready)
                {
                    remaining.Remove(item);
                    emitted.Add(TermKey(item.Source.TermStoreId, item.Source.TermSetId, item.Source.TermId));
                    yield return item;
                }
            }
        }

        private static string Property(IDictionary<string, string> values, string name)
        {
            if (values != null && values.TryGetValue(name, out var value))
            {
                return value;
            }
            return null;
        }

        private static string SetKey(Guid storeId, Guid setId)
        {
            return storeId.ToString("D") + "/" + setId.ToString("D");
        }

        private static string TermKey(Guid storeId, Guid setId, Guid termId)
        {
            return SetKey(storeId, setId) + "/" + termId.ToString("D");
        }

        private static string ProbeDigest(TaxonomyTermSetTargetProbe probe)
        {
            return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(probe));
        }

        private static IList<string> VerificationAssertions(
            TaxonomyAssetTargetDisposition disposition,
            Guid targetSetId,
            IEnumerable<TaxonomyTermMaterializationPlan> termPlans = null,
            IEnumerable<TaxonomyTermTargetProbe> termProbes = null)
        {
            var assertions = new List<string>
            {
                "Fresh readback resolves exactly one target TermSet mapping for " + targetSetId.ToString("D") + "."
            };
            if (disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                || disposition == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift
                || disposition == TaxonomyAssetTargetDisposition.CreateMissing)
            {
                assertions.Add("Fresh readback proves the source identity URN is stored in pnp_reserved_term_original_identifier.");
            }
            else if (disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
            {
                assertions.Add("Fresh readback proves the explicitly approved external TermSet retains its existing ownership and exact shape.");
            }
            var probes = (termProbes ?? Enumerable.Empty<TaxonomyTermTargetProbe>())
                .Where(value => value != null)
                .ToDictionary(value => value.SourceTermId);
            foreach (var term in (termPlans ?? Enumerable.Empty<TaxonomyTermMaterializationPlan>())
                         .Where(value => value?.Source != null)
                         .OrderBy(value => value.Source.TermId))
            {
                probes.TryGetValue(term.Source.TermId, out var probe);
                assertions.AddRange(TaxonomyTermRelationshipFidelity.VerificationAssertions(
                    term,
                    probe,
                    targetSetId));
            }
            return assertions;
        }

        private static void AddTermSetIssue(
            ICollection<MigrationIssue> destination,
            TaxonomyTermSetTargetProbe probe)
        {
            foreach (var issue in probe.Issues)
            {
                issue.Subject = "termset:" + probe.SourceTermStoreId.ToString("D") + "/" + probe.SourceTermSetId.ToString("D");
                destination.Add(issue);
            }
        }

        private static void AddTermGroupIssue(
            ICollection<MigrationIssue> destination,
            TaxonomyTermGroupTargetProbe probe)
        {
            foreach (var issue in probe.Issues)
            {
                issue.Subject = "termgroup:" + probe.SourceTenantId.ToString("D") + "/" + probe.SourceTermStoreId.ToString("D");
                destination.Add(issue);
            }
        }

        private static void AddTermIssue(
            ICollection<MigrationIssue> destination,
            TaxonomyTermTargetProbe probe)
        {
            foreach (var issue in probe.Issues)
            {
                issue.Subject = "term:" + probe.SourceTermStoreId.ToString("D") + "/" + probe.SourceTermSetId.ToString("D") + "/" + probe.SourceTermId.ToString("D");
                destination.Add(issue);
            }
        }

        private static MigrationIssue Issue(string code, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Warning,
                Ingredient = "Taxonomy.Target",
                Message = message
            };
        }
    }
}
