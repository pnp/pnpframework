using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Reporting;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Verification;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationPlanner
    {
        private static readonly RuntimeVerificationRequirement[] RuntimeVerificationRequirements =
        {
            Requirement("page-reachability", RuntimeVerificationRequirementKind.PageReachability, "Fresh navigation reaches the target classic page."),
            Requirement("error-shell-absence", RuntimeVerificationRequirementKind.ErrorShellAbsence, "The target is not a login, access denied, not found, or SharePoint error shell."),
            Requirement("authored-dom-equality", RuntimeVerificationRequirementKind.AuthoredDomEquality, "Normalized authored DOM is equal."),
            Requirement("resource-inventory-equality", RuntimeVerificationRequirementKind.ResourceInventoryEquality, "Authored resource inventory is equal."),
            Requirement("script-inventory-equality", RuntimeVerificationRequirementKind.ScriptInventoryEquality, "Authored script inventory is equal."),
            Requirement("inline-event-inventory-equality", RuntimeVerificationRequirementKind.InlineEventInventoryEquality, "Inline event inventory is equal."),
            Requirement("screenshot-capture", RuntimeVerificationRequirementKind.ScreenshotCapture, "Full-page and authored-canvas screenshots are captured.")
        };

        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PagePlanningOptions options)
        {
            return Plan(targetContext, exportPackage, options, null);
        }

        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PagePlanningOptions options,
            IMigrationArtifactStore artifactStore)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            PublishingPagePackageValidator.ValidateExport(exportPackage, artifactStore);
            ValidateOptions(options);
            if (!string.Equals(exportPackage.Snapshot.SourceProfile, EnterpriseWikiMigrationProfile.SourceProfile, StringComparison.Ordinal))
            {
                throw new InvalidOperationException($"The '{exportPackage.Snapshot.SourceProfile}' source profile cannot be planned by the Enterprise Wiki planner.");
            }

            var targetWeb = targetContext.Web;
            var targetSite = targetContext.Site;
            var targetRootWeb = targetContext.Site.RootWeb;
            targetContext.Load(targetWeb, web => web.Url, web => web.ServerRelativeUrl);
            targetContext.Load(targetSite, site => site.Id);
            targetContext.Load(targetRootWeb, web => web.Url, web => web.ServerRelativeUrl, web => web.Title, web => web.WebTemplate, web => web.Configuration);
            targetContext.ExecuteQueryRetry();

            var snapshot = exportPackage.Snapshot;
            var targetPagePath = PagePath.Normalize(targetWeb.ServerRelativeUrl, options.TargetPageServerRelativeUrl, "Pages");
            var blockers = snapshot.Blockers.ToList();
            var warnings = snapshot.Warnings.ToList();
            AddSecurityPolicyDecision(snapshot, options, blockers);
            AddManagedMetadataDecision(snapshot, options, blockers);

            var layoutMaterialization = PublishingPageLayoutPlanFactory.Create(
                snapshot.Layout,
                new Uri(snapshot.Source.WebUrl),
                new Uri(targetWeb.Url),
                new Uri(targetRootWeb.Url),
                EnterpriseWikiMigrationProfile.PageLayoutFileName,
                options.TaxonomySchemaMappings,
                artifactStore);
            var layoutTargetProbe = layoutMaterialization.Disposition == PublishingPageLayoutMaterializationDisposition.Block
                ? null
                : PublishingPageLayoutTargetInspector.Inspect(targetContext, layoutMaterialization);
            var layoutAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(layoutMaterialization, layoutTargetProbe);
            foreach (var issue in layoutAdmission.Issues)
            {
                blockers.Add($"{issue.Code}: {issue.Message}");
            }
            foreach (var warning in layoutAdmission.Warnings)
            {
                warnings.Add(warning);
            }

            var targetLifecycle = PublishingPageLifecyclePolicy.DeriveTargetLifecycle(snapshot.Lifecycle);
            var lifecycleReason = DescribeLifecycleDecision(snapshot.Lifecycle, targetLifecycle, warnings);
            var replacements = PageReferencePlanner.BuildTextReplacements(snapshot.Source, targetWeb.Url, targetWeb.ServerRelativeUrl);
            var dependencyActions = PageReferencePlanner.BuildActions(
                snapshot.Source,
                snapshot.Dependencies,
                targetWeb.Url,
                targetWeb.ServerRelativeUrl,
                options,
                blockers);
            var targetProbe = EnterpriseWikiTargetInspector.Inspect(
                targetContext,
                targetPagePath,
                dependencyActions,
                targetLifecycle,
                layoutTargetProbe,
                blockers);
            TopologyPlan topology = null;
            ListMigrationPlanSet listMigration = null;
            var webPartActions = new List<ClassicWebPartAction>();
            if (snapshot.SourceTopology != null)
            {
                var topologyResult = new TopologyPlanner().Build(
                    new[] { snapshot.SourceTopology },
                    new[]
                    {
                        new TargetSiteCollectionSpec
                        {
                            SourceSiteId = snapshot.SourceTopology.SiteId,
                            Mode = TargetSiteMode.ExistingTargetSite,
                            TargetSiteUrl = targetRootWeb.Url,
                            ExpectedTargetSiteId = targetSite.Id,
                            Title = targetRootWeb.Title,
                            Template = targetRootWeb.WebTemplate
                        }
                    },
                    options.TopologyPolicy);
                foreach (var issue in topologyResult.Issues)
                {
                    blockers.Add(issue.Code + ": " + issue.Message);
                }
                topology = topologyResult.Plan;
                if (topology != null)
                {
                    var pageWebMapping = topology.SiteCollections.SelectMany(value => value.Webs)
                        .SingleOrDefault(value => value.SourceWebId == snapshot.Source.WebId);
                    if (pageWebMapping == null || !string.Equals(pageWebMapping.TargetWebUrl.TrimEnd('/'), targetWeb.Url.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
                    {
                        blockers.Add("TargetPageWebTopologyMismatch: the target connection Web must be the mapped target for the source page Web. Supply a topology override or connect to the mapped child Web.");
                    }

                    listMigration = ListMigrationPlanFactory.Create(
                        snapshot.ListDependencies,
                        snapshot.ListLookupDependencies,
                        topology,
                        options.TaxonomySchemaMappings,
                        options.ListTargetOverrides);
                    var sourceLists = snapshot.ListDependencies.ToDictionary(value => value.SourceListId);
                    foreach (var listPlan in listMigration.Lists)
                    {
                        if (listPlan.IsExecutable)
                        {
                            listPlan.TargetProbe = ListTargetInspector.Inspect(targetContext, sourceLists[listPlan.SourceListId], listPlan);
                            foreach (var issue in listPlan.TargetProbe.Issues)
                            {
                                blockers.Add(issue.Code + ": " + issue.Message);
                            }
                        }
                        foreach (var issue in listPlan.Issues)
                        {
                            blockers.Add(issue.Code + ": " + issue.Message);
                        }
                    }
                    foreach (var issue in listMigration.Issues)
                    {
                        blockers.Add(issue.Code + ": " + issue.Message);
                    }
                    ListMigrationPlanFactory.SealTargetAnalysis(listMigration);
                }
            }
            else if (snapshot.ListWebPartBindings.Count > 0 || snapshot.ListDependencies.Count > 0)
            {
                blockers.Add("SourceTopologyUnavailable: list-bound Web Parts require the exact source Web ownership closure.");
            }

            var bindingByWebPart = snapshot.ListWebPartBindings.ToDictionary(value => value.SourceWebPartId);
            var listPlans = listMigration == null
                ? new Dictionary<Guid, ListMaterializationPlan>()
                : listMigration.Lists.ToDictionary(value => value.SourceListId);
            foreach (var webPart in snapshot.WebParts)
            {
                ClassicListWebPartBindingSnapshot binding;
                if (!bindingByWebPart.TryGetValue(webPart.Id, out binding))
                {
                    webPartActions.Add(new ClassicWebPartAction
                    {
                        SourceWebPartId = webPart.Id,
                        Disposition = ClassicWebPartDisposition.CopyCaptured,
                        Reason = "Copy the portable shared Web Part export after approved text rewrites."
                    });
                    continue;
                }
                ListMaterializationPlan listPlan;
                if (!listPlans.TryGetValue(binding.SourceListId, out listPlan) || !listPlan.IsExecutable)
                {
                    webPartActions.Add(new ClassicWebPartAction
                    {
                        SourceWebPartId = webPart.Id,
                        Disposition = ClassicWebPartDisposition.Block,
                        SourceListWebId = binding.SourceListWebId,
                        SourceListId = binding.SourceListId,
                        SourceViewId = binding.SourceViewId,
                        Reason = "The bound source List has no executable target materialization plan."
                    });
                    blockers.Add("ListWebPartBindingBlocked: Web Part '" + (webPart.Title ?? webPart.Id.ToString("D")) + "' has no executable target List plan.");
                    continue;
                }
                if (!binding.SourceViewId.HasValue
                    || !listPlan.Views.Any(value => value.SourceViewId == binding.SourceViewId.Value
                        && (value.Disposition == ListViewMaterializationDisposition.CreateOrReuseWebPartView
                            || value.Disposition == ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView)))
                {
                    webPartActions.Add(new ClassicWebPartAction
                    {
                        SourceWebPartId = webPart.Id,
                        Disposition = ClassicWebPartDisposition.Block,
                        SourceListWebId = binding.SourceListWebId,
                        SourceListId = binding.SourceListId,
                        SourceViewId = binding.SourceViewId,
                        Reason = "The captured List Web Part View has no executable target View plan."
                    });
                    blockers.Add("ListWebPartViewMappingBlocked: Web Part '" + (webPart.Title ?? webPart.Id.ToString("D")) + "' has no captured executable source View identity.");
                    continue;
                }
                webPartActions.Add(new ClassicWebPartAction
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
            var fieldActions = PageFieldPlanner.BuildActions(
                targetContext,
                snapshot.Fields,
                EnterpriseWikiMigrationProfile.HandledFieldNames,
                EnterpriseWikiMigrationProfile.AdditionalFieldNames,
                options,
                warnings);
            var expectedContent = PageTextTransformer.Rewrite(snapshot.PublishingPageContent, replacements);
            var expectedContentDigest = PublishingPageDigest.ComputeSha256(expectedContent);
            var plan = new PublishingPageMigrationPlan
            {
                SourceSnapshotDigest = exportPackage.SnapshotDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                TargetWebUrl = targetWeb.Url.TrimEnd('/'),
                TargetWebServerRelativeUrl = targetWeb.ServerRelativeUrl,
                TargetPageServerRelativeUrl = targetPagePath,
                PageLayoutName = layoutMaterialization.TargetPageLayoutName,
                Operation = PageMigrationOperation.CreatePage,
                TargetLifecycle = targetLifecycle,
                LifecycleReason = lifecycleReason,
                CreateOnly = options.CreateOnly,
                PlanningPolicy = CopyOptions(options, targetPagePath),
                TargetProbe = targetProbe,
                LayoutMaterialization = layoutMaterialization,
                LayoutTargetProbe = layoutTargetProbe,
                LayoutAdmission = layoutAdmission,
                FieldActions = fieldActions,
                DependencyActions = dependencyActions,
                Topology = topology,
                ListMigration = listMigration,
                WebPartActions = webPartActions.OrderBy(value => value.SourceWebPartId).ToList(),
                Replacements = replacements,
                ExpectedPublishingPageContentSha256 = expectedContentDigest,
                StorageAssertions = PageStorageAssertionBuilder.Build(
                    snapshot,
                    targetPagePath,
                    dependencyActions,
                    expectedContentDigest,
                    targetLifecycle),
                RuntimeVerification = new RuntimeVerificationManifest
                {
                    Requirements = RuntimeVerificationRequirements.Select(CopyRequirement).ToList()
                },
                Blockers = blockers.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()
            };
            var package = new PublishingPageMigrationPackage
            {
                PlannedAtUtc = DateTimeOffset.UtcNow,
                ExportedAtUtc = exportPackage.ExportedAtUtc,
                State = plan.IsExecutable ? PublishingPagePackageState.ApprovalReady : PublishingPagePackageState.Blocked,
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = exportPackage.SnapshotDigest,
                PlanDigest = PublishingPageDigest.ComputePlanDigest(plan),
                Report = BuildReportSummary(snapshot, plan)
            };
            PublishingPagePackageValidator.ValidateMigration(package);
            return package;
        }

        private static RuntimeVerificationRequirement Requirement(
            string id,
            RuntimeVerificationRequirementKind kind,
            string description)
        {
            return new RuntimeVerificationRequirement
            {
                Id = id,
                Kind = kind,
                Description = description,
                Required = true
            };
        }

        private static RuntimeVerificationRequirement CopyRequirement(RuntimeVerificationRequirement requirement)
        {
            return new RuntimeVerificationRequirement
            {
                Id = requirement.Id,
                Kind = requirement.Kind,
                Required = requirement.Required,
                Description = requirement.Description
            };
        }

        private static void ValidateOptions(PagePlanningOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.TargetPageServerRelativeUrl))
            {
                throw new ArgumentException("A target page path is required.", nameof(options));
            }

            if (!options.CreateOnly)
            {
                throw new NotSupportedException("Only create-page plans are supported. Deferred-field recovery remains represented by the package schema but is not executable yet.");
            }
        }

        private static void AddSecurityPolicyDecision(
            PublishingPageCaptureBundle snapshot,
            PagePlanningOptions options,
            ICollection<string> blockers)
        {
            if (snapshot.Security.HasUniqueRoleAssignments && options.RequireInheritedPermissions)
            {
                blockers.Add("The source page has unique role assignments. The Enterprise Wiki profile requires inherited permissions.");
            }
        }

        private static void AddManagedMetadataDecision(
            PublishingPageCaptureBundle snapshot,
            PagePlanningOptions options,
            ICollection<string> blockers)
        {
            var managedMetadata = snapshot.Fields
                .Where(field => field.HasValue)
                .Where(field => EnterpriseWikiMigrationProfile.AdditionalFieldNames.Contains(field.InternalName))
                .Where(field => field.Kind == PageFieldValueKind.Taxonomy || field.Kind == PageFieldValueKind.TaxonomyCollection)
                .ToArray();
            if (managedMetadata.Length > 0 && options.BlockOnManagedMetadata)
            {
                blockers.Add($"The source snapshot contains {managedMetadata.Length} non-empty managed metadata field value(s), but no reviewed target term mapping was supplied.");
            }
        }

        private static string DescribeLifecycleDecision(
            PageLifecycleSnapshot lifecycle,
            PublishingPageTargetLifecycle targetLifecycle,
            ICollection<string> warnings)
        {
            if (lifecycle == null || string.IsNullOrWhiteSpace(lifecycle.Level))
            {
                warnings.Add("Source lifecycle evidence is incomplete. The conservative target lifecycle is Draft.");
                return "Source lifecycle evidence is incomplete, so the target will remain Draft.";
            }

            if (string.Equals(lifecycle.Level, "Published", StringComparison.OrdinalIgnoreCase)
                && targetLifecycle == PublishingPageTargetLifecycle.Draft)
            {
                warnings.Add("Source lifecycle evidence is contradictory. The conservative target lifecycle is Draft.");
                return $"The source reports Published but has conflicting checkout '{lifecycle.CheckOutType ?? "unknown"}' or moderation '{lifecycle.ModerationStatus?.ToString(CultureInfo.InvariantCulture) ?? "unknown"}' evidence, so the target will remain Draft.";
            }

            return targetLifecycle == PublishingPageTargetLifecycle.Published
                ? "The source file level is Published with no conflicting checkout or moderation evidence, so the target will be published."
                : $"The source file level is '{lifecycle.Level}', so the target will remain Draft.";
        }

        private static PagePlanningOptions CopyOptions(
            PagePlanningOptions options,
            string targetPagePath)
        {
            return new PagePlanningOptions
            {
                TargetPageServerRelativeUrl = targetPagePath,
                RequireInheritedPermissions = options.RequireInheritedPermissions,
                BlockOnManagedMetadata = options.BlockOnManagedMetadata,
                AllowExternalResourceReferences = options.AllowExternalResourceReferences,
                CreateOnly = options.CreateOnly,
                TaxonomySchemaMappings = (options.TaxonomySchemaMappings ?? new List<TaxonomyTargetMapping>())
                    .Select(value => new TaxonomyTargetMapping
                    {
                        SourceTermStoreId = value.SourceTermStoreId,
                        SourceTermSetId = value.SourceTermSetId,
                        TargetTermStoreId = value.TargetTermStoreId,
                        TargetTermSetId = value.TargetTermSetId
                    })
                    .ToList(),
                TopologyPolicy = new TopologyPlanningPolicy
                {
                    DefaultChildWebTemplate = options.TopologyPolicy?.DefaultChildWebTemplate,
                    DefaultChildWebConfiguration = options.TopologyPolicy == null ? 0 : options.TopologyPolicy.DefaultChildWebConfiguration,
                    WebOverrides = (options.TopologyPolicy?.WebOverrides ?? new List<TargetWebOverride>()).Select(value => new TargetWebOverride
                    {
                        SourceWebId = value.SourceWebId,
                        TargetUrlSegment = value.TargetUrlSegment,
                        TargetTitle = value.TargetTitle,
                        TargetTemplate = value.TargetTemplate,
                        TargetConfiguration = value.TargetConfiguration
                    }).ToList()
                },
                ListTargetOverrides = (options.ListTargetOverrides ?? new List<ListTargetOverride>()).Select(value => new ListTargetOverride
                {
                    SourceListId = value.SourceListId,
                    TargetTitle = value.TargetTitle,
                    TargetRootFolderServerRelativeUrl = value.TargetRootFolderServerRelativeUrl
                }).ToList()
            };
        }

        private static PublishingPageMigrationReport BuildReportSummary(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            return new PublishingPageMigrationReport
            {
                Summary = plan.IsExecutable
                    ? "Source export and target analysis completed. Import requires explicit approval of the sealed plan digest."
                    : "The package is sealed for review but cannot be imported until every blocker is resolved and a new plan is generated.",
                CapturedIngredients = new List<string>
                {
                    "Page/file/list item identity and source stability fence",
                    $"{snapshot.SourceProfile} content type and publishing layout evidence",
                    $"All {snapshot.Fields.Count} source Pages-library field definitions and returned values",
                    $"{snapshot.WebParts.Count} shared Web Part export(s) with zone placement",
                    $"{snapshot.Dependencies.Count} authored dependency/link snapshot(s)",
                    "Page security inheritance and source lifecycle evidence",
                    "Target publishing library, versioning, lifecycle, field, layout, and create-only probes"
                },
                Blockers = plan.Blockers.ToList(),
                Warnings = plan.Warnings.ToList()
            };
        }
    }
}
