using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Verification;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Pages.Ingredients;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public sealed class PublishingPageImportReceipt
    {
        public string SchemaVersion { get; set; } = PublishingPagePackageContract.ReceiptSchemaVersion;

        public DateTimeOffset StartedAtUtc { get; set; }

        public DateTimeOffset CompletedAtUtc { get; set; }

        public Guid OperationId { get; set; }

        public MigrationExecutionStatus ExecutionStatus { get; set; }

        public bool PartialExecution { get; set; }

        public PageIngredientExecutionFrontier ExecutionFrontier { get; set; }

        public IList<string> CompletedIngredientIds { get; set; } = new List<string>();

        public IList<string> VerifiedIngredientIds { get; set; } = new List<string>();

        public IList<string> PendingVerificationIngredientIds { get; set; } = new List<string>();

        public IList<string> FailedVerificationIngredientIds { get; set; } = new List<string>();

        public int DeferredIngredientCount { get; set; }

        public int AuthorizationBlockedIngredientCount { get; set; }

        public ExecutionAdmissionFailure AdmissionFailure { get; set; }

        public bool MutationStarted { get; set; }

        public IList<MigrationMutationReceipt> Steps { get; set; } = new List<MigrationMutationReceipt>();

        public string ApprovedPlanDigest { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public Guid TargetFileUniqueId { get; set; }

        public int TargetListItemId { get; set; }

        public string TargetContentTypeId { get; set; }

        public string TargetVersionLabel { get; set; }

        public PublishingPageTargetLifecycle ExpectedLifecycle { get; set; }

        public PublishingPageTargetLifecycle ApprovedLifecycle { get; set; }

        public string ActualFileLevel { get; set; }

        public string ActualCheckOutType { get; set; }

        public int? ActualModerationStatus { get; set; }

        public bool LifecycleMatched { get; set; }

        public bool SecurityMatched { get; set; }

        public bool OwnershipMatched { get; set; }

        public bool PageArtifactMatched { get; set; }

        public bool LayoutMatched { get; set; }

        public bool ContentTypeMatched { get; set; }

        public bool PageFieldsMatched { get; set; }

        public bool DependenciesMatched { get; set; }

        public string ApprovedPublishingPageContentSha256 { get; set; }

        public string ExpectedPublishingPageContentSha256 { get; set; }

        public string PersistedPublishingPageContentSha256 { get; set; }

        public bool StorageContentEqual { get; set; }

        public int ImportedWebPartCount { get; set; }

        public bool WebPartsMatched { get; set; }

        public IList<PublishingPageWebPartVerificationResult> WebPartResults { get; set; } = new List<PublishingPageWebPartVerificationResult>();

        public int MaterializedDependencyCount { get; set; }

        public TopologyMaterializationReceipt TopologyMaterialization { get; set; }

        public bool TopologyMatched { get; set; }

        public IList<ListMaterializationReceipt> ListMaterializations { get; set; } = new List<ListMaterializationReceipt>();

        public bool ListsMatched { get; set; }

        public IList<PageFieldImportResult> FieldResults { get; set; } = new List<PageFieldImportResult>();

        public bool TaxonomyRelationshipsMatched { get; set; }

        public IList<TaxonomyRelationshipVerificationResult> TaxonomyRelationshipResults { get; set; } = new List<TaxonomyRelationshipVerificationResult>();

        public bool FreshReadbackPassed { get; set; }

        public StorageVerificationStatus StorageVerificationStatus { get; set; }

        public RuntimeVerificationStatus RuntimeVerificationStatus { get; set; }

        public MigrationAcceptanceStatus AcceptanceStatus { get; set; }

        public IList<string> Warnings { get; set; } = new List<string>();

    }
}
