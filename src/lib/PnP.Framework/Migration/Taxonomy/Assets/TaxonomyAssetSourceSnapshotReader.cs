using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    /// <summary>
    /// Captures the exact source TermSet metadata and the smallest required Term
    /// ancestor closure. It never enumerates unrelated Terms merely because a
    /// field is bound to a large shared TermSet.
    /// </summary>
    public static class TaxonomyAssetSourceSnapshotReader
    {
        public static TaxonomyAssetSourceSnapshot Read(
            ClientContext context,
            Guid sourceTenantId,
            IEnumerable<TaxonomyTermSetCaptureRequest> requests)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (sourceTenantId == Guid.Empty)
            {
                throw new ArgumentException("A source tenant identity is required.", nameof(sourceTenantId));
            }

            var normalized = Normalize(requests);
            var result = new TaxonomyAssetSourceSnapshot { SourceTenantId = sourceTenantId };
            if (normalized.Count == 0)
            {
                result.SnapshotDigest = ComputeDigest(result);
                return result;
            }

            var session = TaxonomySession.GetTaxonomySession(context);
            foreach (var storeGroup in normalized.GroupBy(value => value.SourceTermStoreId).OrderBy(value => value.Key))
            {
                var store = session.TermStores.GetById(storeGroup.Key);
                context.Load(store, value => value.Id);
                context.ExecuteQueryRetry();
                if (store.ServerObjectIsNull.GetValueOrDefault(true))
                {
                    foreach (var request in storeGroup)
                    {
                        result.TermSets.Add(UnavailableSet(sourceTenantId, request, "The source TermStore does not exist or is not readable."));
                    }
                    continue;
                }

                var sets = new List<SetRead>();
                foreach (var request in storeGroup.OrderBy(value => value.SourceTermSetId))
                {
                    var termSet = store.GetTermSet(request.SourceTermSetId);
                    context.Load(
                        termSet,
                        value => value.Id,
                        value => value.Name,
                        value => value.IsOpenForTermCreation,
                        value => value.IsAvailableForTagging);
                    sets.Add(new SetRead(request, termSet));
                }
                context.ExecuteQueryRetry();

                foreach (var item in sets)
                {
                    if (item.TermSet.ServerObjectIsNull.GetValueOrDefault(true))
                    {
                        result.TermSets.Add(UnavailableSet(sourceTenantId, item.Request, "The source TermSet does not exist or is not readable."));
                        continue;
                    }
                    var snapshot = new TaxonomyTermSetSourceSnapshot
                    {
                        SourceTenantId = sourceTenantId,
                        SourceTermStoreId = item.Request.SourceTermStoreId,
                        SourceTermSetId = item.Request.SourceTermSetId,
                        SourceWebUrl = context.Url,
                        Name = item.TermSet.Name,
                        IsOpenForTermCreation = item.TermSet.IsOpenForTermCreation,
                        IsAvailableForTagging = item.TermSet.IsAvailableForTagging,
                        Availability = EvidenceAvailability.Captured,
                        Consumers = item.Request.Consumers
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Distinct(StringComparer.Ordinal)
                            .OrderBy(value => value, StringComparer.Ordinal)
                            .ToList()
                    };
                    snapshot.EvidenceSha256 = ComputeDigest(snapshot);
                    result.TermSets.Add(snapshot);

                    ReadRequiredTerms(
                        context,
                        store,
                        sourceTenantId,
                        item.Request,
                        result.Terms,
                        result.Diagnostics);
                }
            }

            result.TermSets = result.TermSets
                .OrderBy(value => value.SourceTermStoreId)
                .ThenBy(value => value.SourceTermSetId)
                .ToList();
            result.Terms = result.Terms
                .GroupBy(value => TermKey(value.SourceTermStoreId, value.SourceTermSetId, value.SourceTermId), StringComparer.Ordinal)
                .Select(group => group.First())
                .OrderBy(value => value.SourceTermStoreId)
                .ThenBy(value => value.SourceTermSetId)
                .ThenBy(value => value.SourceTermId)
                .ToList();
            result.Diagnostics = result.Diagnostics
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList();
            result.SnapshotDigest = ComputeDigest(result);
            return result;
        }

        private static void ReadRequiredTerms(
            ClientContext context,
            TermStore store,
            Guid sourceTenantId,
            TaxonomyTermSetCaptureRequest request,
            ICollection<TaxonomyTermSourceSnapshot> destination,
            ICollection<string> diagnostics)
        {
            var pending = new HashSet<Guid>((request.RequiredTermIds ?? new List<Guid>()).Where(value => value != Guid.Empty));
            var observed = new HashSet<Guid>();
            while (pending.Count > 0)
            {
                var current = pending.Where(value => !observed.Contains(value)).OrderBy(value => value).ToArray();
                pending.Clear();
                if (current.Length == 0)
                {
                    break;
                }

                var reads = new List<TermRead>();
                foreach (var termId in current)
                {
                    observed.Add(termId);
                    var term = store.GetTermInTermSet(request.SourceTermSetId, termId);
                    context.Load(
                        term,
                        value => value.Id,
                        value => value.Name,
                        value => value.PathOfTerm,
                        value => value.IsAvailableForTagging,
                        value => value.IsReused,
                        value => value.IsSourceTerm,
                        value => value.SourceTerm,
                        value => value.PinSourceTermSet,
                        value => value.Parent,
                        value => value.TermSet);
                    reads.Add(new TermRead(termId, term));
                }
                context.ExecuteQueryRetry();

                var materialized = reads
                    .Where(value => !value.Term.ServerObjectIsNull.GetValueOrDefault(true))
                    .Select(value => value.Term)
                    .ToArray();
                foreach (var term in materialized)
                {
                    context.Load(term.TermSet, value => value.Id);
                    context.Load(term.TermSets, values => values.Include(value => value.Id));
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

                foreach (var read in reads)
                {
                    if (read.Term.ServerObjectIsNull.GetValueOrDefault(true))
                    {
                        var unavailable = new TaxonomyTermSourceSnapshot
                        {
                            SourceTenantId = sourceTenantId,
                            SourceTermStoreId = request.SourceTermStoreId,
                            SourceTermSetId = request.SourceTermSetId,
                            SourceTermId = read.RequestedTermId,
                            SourceWebUrl = context.Url,
                            Availability = EvidenceAvailability.Unavailable,
                            Diagnostics = new List<string> { "The required source Term is absent from the captured source TermSet." }
                        };
                        unavailable.EvidenceSha256 = ComputeDigest(unavailable);
                        destination.Add(unavailable);
                        diagnostics.Add("Required source taxonomy Term is unavailable: " + TermKey(request.SourceTermStoreId, request.SourceTermSetId, read.RequestedTermId) + ".");
                        continue;
                    }

                    Guid? parentId = null;
                    if (read.Term.Parent != null
                        && !read.Term.Parent.ServerObjectIsNull.GetValueOrDefault(true)
                        && read.Term.Parent.Id != Guid.Empty)
                    {
                        parentId = read.Term.Parent.Id;
                        if (!observed.Contains(parentId.Value))
                        {
                            pending.Add(parentId.Value);
                        }
                    }
                    var snapshot = new TaxonomyTermSourceSnapshot
                    {
                        SourceTenantId = sourceTenantId,
                        SourceTermStoreId = request.SourceTermStoreId,
                        SourceTermSetId = request.SourceTermSetId,
                        SourceTermId = read.Term.Id,
                        SourceWebUrl = context.Url,
                        SourceParentTermId = parentId,
                        Name = read.Term.Name,
                        Path = read.Term.PathOfTerm,
                        IsAvailableForTagging = read.Term.IsAvailableForTagging,
                        IsReused = read.Term.IsReused,
                        IsSourceTerm = read.Term.IsSourceTerm,
                        ReuseSourceTermId = Id(read.Term.SourceTerm),
                        TermSetIds = ReadTermSetIds(read.Term),
                        PinSourceTermSetId = Id(read.Term.PinSourceTermSet),
                        Availability = EvidenceAvailability.Captured
                    };
                    snapshot.EvidenceSha256 = ComputeDigest(snapshot);
                    destination.Add(snapshot);
                }
            }
        }

        private static IList<TaxonomyTermSetCaptureRequest> Normalize(
            IEnumerable<TaxonomyTermSetCaptureRequest> requests)
        {
            var result = new List<TaxonomyTermSetCaptureRequest>();
            foreach (var group in (requests ?? Enumerable.Empty<TaxonomyTermSetCaptureRequest>())
                         .Where(value => value != null
                             && value.SourceTermStoreId != Guid.Empty
                             && value.SourceTermSetId != Guid.Empty)
                         .GroupBy(value => SetKey(value.SourceTermStoreId, value.SourceTermSetId), StringComparer.Ordinal))
            {
                var first = group.First();
                result.Add(new TaxonomyTermSetCaptureRequest
                {
                    SourceTermStoreId = first.SourceTermStoreId,
                    SourceTermSetId = first.SourceTermSetId,
                    SourceWebUrls = group
                        .SelectMany(value => value.SourceWebUrls ?? new List<string>())
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                        .ToList(),
                    RequiredTermIds = group
                        .SelectMany(value => value.RequiredTermIds ?? new List<Guid>())
                        .Where(value => value != Guid.Empty)
                        .Distinct()
                        .OrderBy(value => value)
                        .ToList(),
                    Consumers = group
                        .SelectMany(value => value.Consumers ?? new List<string>())
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.Ordinal)
                        .OrderBy(value => value, StringComparer.Ordinal)
                        .ToList()
                });
            }
            return result
                .OrderBy(value => value.SourceTermStoreId)
                .ThenBy(value => value.SourceTermSetId)
                .ToList();
        }

        private static TaxonomyTermSetSourceSnapshot UnavailableSet(
            Guid sourceTenantId,
            TaxonomyTermSetCaptureRequest request,
            string diagnostic)
        {
            var snapshot = new TaxonomyTermSetSourceSnapshot
            {
                SourceTenantId = sourceTenantId,
                SourceTermStoreId = request.SourceTermStoreId,
                SourceTermSetId = request.SourceTermSetId,
                SourceWebUrl = request.SourceWebUrls.FirstOrDefault(),
                Availability = EvidenceAvailability.Unavailable,
                Consumers = request.Consumers.ToList(),
                Diagnostics = new List<string> { diagnostic }
            };
            snapshot.EvidenceSha256 = ComputeDigest(snapshot);
            return snapshot;
        }

        public static string ComputeDigest(TaxonomyTermSetSourceSnapshot snapshot)
        {
            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    snapshot,
                    nameof(TaxonomyTermSetSourceSnapshot.EvidenceSha256)));
        }

        public static string ComputeDigest(TaxonomyTermSourceSnapshot snapshot)
        {
            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    snapshot,
                    nameof(TaxonomyTermSourceSnapshot.EvidenceSha256)));
        }

        public static string ComputeDigest(TaxonomyAssetSourceSnapshot snapshot)
        {
            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    snapshot,
                    nameof(TaxonomyAssetSourceSnapshot.SnapshotDigest)));
        }

        private static string SetKey(Guid storeId, Guid setId)
        {
            return storeId.ToString("D") + "/" + setId.ToString("D");
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

        private static string TermKey(Guid storeId, Guid setId, Guid termId)
        {
            return SetKey(storeId, setId) + "/" + termId.ToString("D");
        }

        private sealed class SetRead
        {
            public SetRead(TaxonomyTermSetCaptureRequest request, TermSet termSet)
            {
                Request = request;
                TermSet = termSet;
            }

            public TaxonomyTermSetCaptureRequest Request { get; }

            public TermSet TermSet { get; }
        }

        private sealed class TermRead
        {
            public TermRead(Guid requestedTermId, Term term)
            {
                RequestedTermId = requestedTermId;
                Term = term;
            }

            public Guid RequestedTermId { get; }

            public Term Term { get; }
        }
    }
}
