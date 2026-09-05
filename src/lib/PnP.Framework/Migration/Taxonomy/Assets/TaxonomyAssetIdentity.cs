using PnP.Framework.Migration.Packaging;
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    /// <summary>
    /// Creates stable source identities and deterministic preferred target
    /// identities for taxonomy assets. Source WssIds are deliberately excluded.
    /// </summary>
    public static class TaxonomyAssetIdentity
    {
        public const string OriginalIdentifierPropertyName = "pnp_reserved_term_original_identifier";

        public const string TargetGroupName = "PnP Repro4 Migrated Taxonomy";

        private static readonly Guid GroupNamespace = new Guid("31ba6754-7f98-4df7-8383-b377c8d70c87");

        public static string TermGroup(TaxonomyTermGroupSourceIdentity source)
        {
            Validate(source, nameof(source));
            return string.Format(
                "urn:pnp:spo-termgroup:v1:{0}:{1}",
                source.TenantId.ToString("N"),
                source.TermStoreId.ToString("N"));
        }

        public static string TermSet(TaxonomyTermSetSourceIdentity source)
        {
            Validate(source, nameof(source));
            return string.Format(
                "urn:pnp:spo-termset:v1:{0}:{1}:{2}",
                source.TenantId.ToString("N"),
                source.TermStoreId.ToString("N"),
                source.TermSetId.ToString("N"));
        }

        public static string Term(TaxonomyTermSourceIdentity source)
        {
            Validate(source, nameof(source));
            return string.Format(
                "urn:pnp:spo-term:v1:{0}:{1}:{2}:{3}",
                source.TenantId.ToString("N"),
                source.TermStoreId.ToString("N"),
                source.TermSetId.ToString("N"),
                source.TermId.ToString("N"));
        }

        public static Guid TargetGroupId(Guid sourceTenantId, Guid sourceTermStoreId)
        {
            Validate(sourceTenantId, nameof(sourceTenantId));
            Validate(sourceTermStoreId, nameof(sourceTermStoreId));
            return DeterministicGuid(string.Format(
                "{0:N}:{1:N}:{2:N}",
                GroupNamespace,
                sourceTenantId,
                sourceTermStoreId));
        }

        public static TaxonomyTermGroupMaterializationPlan CreateTermGroupPlan(
            Guid sourceTenantId,
            Guid sourceTermStoreId,
            Guid targetTermStoreId)
        {
            Validate(sourceTenantId, nameof(sourceTenantId));
            Validate(sourceTermStoreId, nameof(sourceTermStoreId));
            Validate(targetTermStoreId, nameof(targetTermStoreId));
            var plan = new TaxonomyTermGroupMaterializationPlan
            {
                Source = new TaxonomyTermGroupSourceIdentity
                {
                    TenantId = sourceTenantId,
                    TermStoreId = sourceTermStoreId
                },
                TargetTermStoreId = targetTermStoreId,
                PreferredTargetGroupId = TargetGroupId(sourceTenantId, sourceTermStoreId),
                TargetGroupName = TargetGroupName
            };
            plan.PlanDigest = ComputePlanDigest(plan);
            return plan;
        }

        public static TaxonomyTermSetMaterializationPlan CreateTermSetPlan(
            TaxonomyTermSetSourceSnapshot source,
            Guid targetTermStoreId)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            Validate(source.SourceTenantId, nameof(source.SourceTenantId));
            Validate(source.SourceTermStoreId, nameof(source.SourceTermStoreId));
            Validate(source.SourceTermSetId, nameof(source.SourceTermSetId));
            Validate(targetTermStoreId, nameof(targetTermStoreId));
            if (string.IsNullOrWhiteSpace(source.Name))
            {
                throw new ArgumentException("A captured source TermSet name is required.", nameof(source));
            }

            var identity = new TaxonomyTermSetSourceIdentity
            {
                TenantId = source.SourceTenantId,
                TermStoreId = source.SourceTermStoreId,
                TermSetId = source.SourceTermSetId
            };
            var compactName = source.Name.Trim();
            if (compactName.Length > 180)
            {
                compactName = compactName.Substring(0, 180).TrimEnd();
            }
            var plan = new TaxonomyTermSetMaterializationPlan
            {
                Source = identity,
                TargetTermStoreId = targetTermStoreId,
                TargetGroupId = TargetGroupId(source.SourceTenantId, source.SourceTermStoreId),
                TargetGroupName = TargetGroupName,
                PreferredTargetTermSetId = source.SourceTermSetId,
                SourceTermSetName = source.Name,
                TargetTermSetName = compactName + " [" + source.SourceTermSetId.ToString("N") + "]",
                Language = source.Language <= 0 ? 1033 : source.Language,
                IsOpenForTermCreation = source.IsOpenForTermCreation,
                IsAvailableForTagging = source.IsAvailableForTagging,
                OriginalIdentifierPropertyName = OriginalIdentifierPropertyName,
                OriginalIdentifier = TermSet(identity),
                SourceEvidenceSha256 = source.EvidenceSha256
            };
            plan.PlanDigest = ComputePlanDigest(plan);
            return plan;
        }

        public static TaxonomyTermMaterializationPlan CreateTermPlan(
            TaxonomyTermSourceSnapshot source,
            Guid targetTermStoreId,
            Guid targetTermSetId,
            Guid? targetParentTermId)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            Validate(source.SourceTenantId, nameof(source.SourceTenantId));
            Validate(source.SourceTermStoreId, nameof(source.SourceTermStoreId));
            Validate(source.SourceTermSetId, nameof(source.SourceTermSetId));
            Validate(source.SourceTermId, nameof(source.SourceTermId));
            Validate(targetTermStoreId, nameof(targetTermStoreId));
            Validate(targetTermSetId, nameof(targetTermSetId));
            if (string.IsNullOrWhiteSpace(source.Name))
            {
                throw new ArgumentException("A captured source Term name is required.", nameof(source));
            }

            var identity = new TaxonomyTermSourceIdentity
            {
                TenantId = source.SourceTenantId,
                TermStoreId = source.SourceTermStoreId,
                TermSetId = source.SourceTermSetId,
                TermId = source.SourceTermId
            };
            var plan = new TaxonomyTermMaterializationPlan
            {
                Source = identity,
                TargetTermStoreId = targetTermStoreId,
                TargetTermSetId = targetTermSetId,
                TargetParentTermId = targetParentTermId,
                PreferredTargetTermId = source.SourceTermId,
                Name = source.Name.Trim(),
                SourcePath = source.Path,
                Language = source.Language <= 0 ? 1033 : source.Language,
                IsAvailableForTagging = source.IsAvailableForTagging,
                SourceIsReused = source.IsReused,
                SourceIsSourceTerm = source.IsSourceTerm,
                SourceReuseSourceTermId = source.ReuseSourceTermId,
                SourceTermSetIds = (source.TermSetIds ?? new System.Collections.Generic.List<Guid>())
                    .Where(value => value != Guid.Empty)
                    .Distinct()
                    .OrderBy(value => value)
                    .ToList(),
                SourcePinSourceTermSetId = source.PinSourceTermSetId,
                OriginalIdentifierPropertyName = OriginalIdentifierPropertyName,
                OriginalIdentifier = Term(identity),
                SourceEvidenceSha256 = source.EvidenceSha256
            };
            plan.PlanDigest = ComputePlanDigest(plan);
            return plan;
        }

        public static string ComputePlanDigest(TaxonomyTermSetMaterializationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    plan,
                    nameof(TaxonomyTermSetMaterializationPlan.PlanDigest)));
        }

        public static string ComputePlanDigest(TaxonomyTermGroupMaterializationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    plan,
                    nameof(TaxonomyTermGroupMaterializationPlan.PlanDigest)));
        }

        public static string ComputePlanDigest(TaxonomyTermMaterializationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    plan,
                    nameof(TaxonomyTermMaterializationPlan.PlanDigest)));
        }

        private static Guid DeterministicGuid(string value)
        {
            byte[] hash;
            using (var algorithm = SHA256.Create())
            {
                hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value));
            }
            var bytes = new byte[16];
            Array.Copy(hash, bytes, bytes.Length);
            bytes[7] = (byte)((bytes[7] & 0x0f) | 0x50);
            bytes[8] = (byte)((bytes[8] & 0x3f) | 0x80);
            return new Guid(bytes);
        }

        private static void Validate(TaxonomyTermSetSourceIdentity source, string name)
        {
            if (source == null)
            {
                throw new ArgumentNullException(name);
            }
            Validate(source.TenantId, nameof(source.TenantId));
            Validate(source.TermStoreId, nameof(source.TermStoreId));
            Validate(source.TermSetId, nameof(source.TermSetId));
        }

        private static void Validate(TaxonomyTermGroupSourceIdentity source, string name)
        {
            if (source == null)
            {
                throw new ArgumentNullException(name);
            }
            Validate(source.TenantId, nameof(source.TenantId));
            Validate(source.TermStoreId, nameof(source.TermStoreId));
        }

        private static void Validate(TaxonomyTermSourceIdentity source, string name)
        {
            if (source == null)
            {
                throw new ArgumentNullException(name);
            }
            Validate(source.TenantId, nameof(source.TenantId));
            Validate(source.TermStoreId, nameof(source.TermStoreId));
            Validate(source.TermSetId, nameof(source.TermSetId));
            Validate(source.TermId, nameof(source.TermId));
        }

        private static void Validate(Guid value, string name)
        {
            if (value == Guid.Empty)
            {
                throw new ArgumentException("A non-empty GUID is required.", name);
            }
        }
    }
}
