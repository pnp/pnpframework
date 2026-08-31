using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.WebParts;
using PnP.Framework.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using CsomFile = Microsoft.SharePoint.Client.File;

namespace PnP.Framework.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationService
    {
        private const int PublishingPagesListTemplate = 850;

        private static readonly HashSet<string> AdditionalFieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ArticleByLine",
            "PublishingContact",
            "PublishingPageDescription",
            "PublishingPageImage",
            "PublishingRollupImage",
            "SeoBrowserTitle",
            "SeoKeywords",
            "SeoMetaDescription",
            "Wiki_x0020_Page_x0020_Categories"
        };

        private static readonly HashSet<string> HandledFieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ContentTypeId",
            "PublishingPageContent",
            "PublishingPageLayout",
            "Title"
        };

        private static readonly string[] BrowserAssertions =
        {
            "fresh-navigation-reaches-target-classic-page",
            "no-login-access-denied-not-found-or-sharepoint-error-shell",
            "normalized-authored-dom-equal",
            "resource-script-and-inline-event-inventory-equal",
            "full-page-and-authored-canvas-screenshots-captured"
        };

        private static readonly Regex CssUrlPattern = new Regex(
            @"url\(\s*(?:['""](?<url>.*?)['""]|(?<url>[^)]*?))\s*\)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public IReadOnlyList<string> Discover(ClientContext sourceContext)
        {
            if (sourceContext == null)
            {
                throw new ArgumentNullException(nameof(sourceContext));
            }

            var pages = sourceContext.Web.GetPagesLibrary();
            if (pages == null)
            {
                return Array.Empty<string>();
            }

            var query = new CamlQuery
            {
                ViewXml = $@"<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <BeginsWith>
        <FieldRef Name='ContentTypeId' />
        <Value Type='ContentTypeId'>{BuiltInContentTypeId.EnterpriseWikiPage}</Value>
      </BeginsWith>
    </Where>
    <OrderBy><FieldRef Name='FileRef' Ascending='TRUE' /></OrderBy>
  </Query>
  <ViewFields>
    <FieldRef Name='FileRef' />
    <FieldRef Name='ContentTypeId' />
  </ViewFields>
</View>"
            };
            var items = pages.GetItems(query);
            sourceContext.Load(items);
            sourceContext.ExecuteQueryRetry();

            return items
                .Select(item => new
                {
                    ContentTypeId = Convert.ToString(item["ContentTypeId"], CultureInfo.InvariantCulture),
                    FileRef = Convert.ToString(item["FileRef"], CultureInfo.InvariantCulture)
                })
                .Where(item => IsEnterpriseWikiContentType(item.ContentTypeId))
                .Select(item => item.FileRef)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        public EnterpriseWikiExportPackage Export(
            ClientContext sourceContext,
            EnterpriseWikiExportOptions options)
        {
            if (sourceContext == null)
            {
                throw new ArgumentNullException(nameof(sourceContext));
            }

            ValidateExportOptions(options);
            var sourceWeb = sourceContext.Web;
            sourceContext.Load(sourceWeb, web => web.Url, web => web.ServerRelativeUrl, web => web.Title);
            sourceContext.ExecuteQueryRetry();

            var sourcePagePath = NormalizePagePath(sourceWeb.ServerRelativeUrl, options.SourcePageServerRelativeUrl, "Pages");
            var blockers = new List<string>();
            var warnings = new List<string>();
            var sourceCapture = CaptureSourcePage(sourceContext, sourcePagePath, options, blockers, warnings);
            var dependencies = CaptureDependencies(
                sourceContext,
                sourceCapture.Identity,
                sourceCapture.PublishingPageContent,
                sourceCapture.WebParts,
                options,
                warnings);

            if (!sourceCapture.Identity.PageLayoutUrl.EndsWith("/EnterpriseWiki.aspx", StringComparison.OrdinalIgnoreCase))
            {
                blockers.Add($"The source page uses layout '{sourceCapture.Identity.PageLayoutUrl}'. The exact profile requires EnterpriseWiki.aspx.");
            }

            var afterFence = ReadFence(sourceContext, sourcePagePath);
            if (!FenceEquals(sourceCapture.SourceFence, afterFence))
            {
                blockers.Add("The source page changed while it was being exported. Discard this snapshot and export again.");
            }

            var snapshot = new EnterpriseWikiSnapshot
            {
                CapturePolicy = new EnterpriseWikiExportOptions
                {
                    SourcePageServerRelativeUrl = sourcePagePath,
                    IncludeWebParts = options.IncludeWebParts,
                    MaximumDependencyBytes = options.MaximumDependencyBytes
                },
                Source = sourceCapture.Identity,
                PublishingPageContent = sourceCapture.PublishingPageContent,
                PublishingPageContentSha256 = EnterpriseWikiPackageSerializer.ComputeSha256(sourceCapture.PublishingPageContent ?? string.Empty),
                Fields = sourceCapture.Fields.OrderBy(field => field.InternalName, StringComparer.Ordinal).ToList(),
                WebParts = sourceCapture.WebParts
                    .OrderBy(webPart => webPart.ZoneId, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(webPart => webPart.ZoneIndex)
                    .ThenBy(webPart => webPart.Id)
                    .ToList(),
                Dependencies = dependencies
                    .OrderBy(dependency => dependency.SourceAbsoluteUrl, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(dependency => dependency.Consumer, StringComparer.Ordinal)
                    .ToList(),
                Security = sourceCapture.Security,
                Lifecycle = sourceCapture.Lifecycle,
                SourceFence = sourceCapture.SourceFence,
                Blockers = blockers.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()
            };

            return new EnterpriseWikiExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Snapshot = snapshot,
                SnapshotDigest = EnterpriseWikiPackageSerializer.ComputeSnapshotDigest(snapshot)
            };
        }

        public EnterpriseWikiMigrationPackage Plan(
            ClientContext targetContext,
            EnterpriseWikiExportPackage exportPackage,
            EnterpriseWikiPlanningOptions options)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            EnterpriseWikiPackageSerializer.ValidateExport(exportPackage);
            ValidatePlanningOptions(options);
            var snapshot = exportPackage.Snapshot;
            var targetWeb = targetContext.Web;
            targetContext.Load(targetWeb, web => web.Url, web => web.ServerRelativeUrl, web => web.Title, web => web.WebTemplate, web => web.Configuration);
            targetContext.ExecuteQueryRetry();

            var targetPagePath = NormalizePagePath(targetWeb.ServerRelativeUrl, options.TargetPageServerRelativeUrl, "Pages");
            var blockers = snapshot.Blockers.ToList();
            var warnings = snapshot.Warnings.ToList();
            if (snapshot.Security.HasUniqueRoleAssignments && options.RequireInheritedPermissions)
            {
                blockers.Add("The source page has unique role assignments. The exact profile requires inherited permissions.");
            }

            var managedMetadata = snapshot.Fields
                .Where(field => field.HasValue)
                .Where(field => AdditionalFieldNames.Contains(field.InternalName))
                .Where(field => field.Kind == EnterpriseWikiFieldValueKind.Taxonomy || field.Kind == EnterpriseWikiFieldValueKind.TaxonomyCollection)
                .ToArray();
            if (managedMetadata.Length > 0 && options.BlockOnManagedMetadata)
            {
                blockers.Add($"The source snapshot contains {managedMetadata.Length} non-empty managed metadata field value(s), but no reviewed target term mapping was supplied.");
            }

            var targetLifecycle = DeriveTargetLifecycle(snapshot.Lifecycle);
            var lifecycleReason = targetLifecycle == EnterpriseWikiTargetLifecycle.Published
                ? "The source file level is Published with no conflicting checkout or moderation evidence, so the target will be published."
                : $"The source file level is '{snapshot.Lifecycle?.Level ?? "unknown"}', so the target will remain Draft.";
            if (snapshot.Lifecycle == null || string.IsNullOrWhiteSpace(snapshot.Lifecycle.Level))
            {
                warnings.Add("Source lifecycle evidence is incomplete. The conservative target lifecycle is Draft.");
            }
            else if (string.Equals(snapshot.Lifecycle.Level, "Published", StringComparison.OrdinalIgnoreCase)
                && targetLifecycle == EnterpriseWikiTargetLifecycle.Draft)
            {
                lifecycleReason = $"The source reports Published but has conflicting checkout '{snapshot.Lifecycle.CheckOutType ?? "unknown"}' or moderation '{snapshot.Lifecycle.ModerationStatus?.ToString(CultureInfo.InvariantCulture) ?? "unknown"}' evidence, so the target will remain Draft.";
                warnings.Add("Source lifecycle evidence is contradictory. The conservative target lifecycle is Draft.");
            }

            var replacements = BuildReplacements(snapshot.Source, targetWeb.Url, targetWeb.ServerRelativeUrl);
            var dependencyActions = BuildDependencyActions(snapshot, targetWeb.Url, targetWeb.ServerRelativeUrl, options, blockers, warnings);
            var targetProbe = ProbeTarget(targetContext, targetPagePath, dependencyActions, targetLifecycle, blockers, warnings);
            var fieldActions = BuildFieldActions(targetContext, snapshot.Fields, options, blockers, warnings);
            var expectedContent = RewriteContent(snapshot.PublishingPageContent, replacements);
            var expectedContentDigest = EnterpriseWikiPackageSerializer.ComputeSha256(expectedContent);
            var plan = new EnterpriseWikiMigrationPlan
            {
                SourceSnapshotDigest = exportPackage.SnapshotDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                TargetWebUrl = targetWeb.Url.TrimEnd('/'),
                TargetWebServerRelativeUrl = targetWeb.ServerRelativeUrl,
                TargetPageServerRelativeUrl = targetPagePath,
                PageLayoutName = "EnterpriseWiki",
                Operation = EnterpriseWikiMigrationOperation.CreatePage,
                TargetLifecycle = targetLifecycle,
                LifecycleReason = lifecycleReason,
                CreateOnly = options.CreateOnly,
                PlanningPolicy = new EnterpriseWikiPlanningOptions
                {
                    TargetPageServerRelativeUrl = targetPagePath,
                    RequireInheritedPermissions = options.RequireInheritedPermissions,
                    BlockOnManagedMetadata = options.BlockOnManagedMetadata,
                    AllowExternalResourceReferences = options.AllowExternalResourceReferences,
                    CreateOnly = options.CreateOnly
                },
                TargetProbe = targetProbe,
                FieldActions = fieldActions,
                DependencyActions = dependencyActions,
                Replacements = replacements,
                ExpectedPublishingPageContentSha256 = expectedContentDigest,
                StorageAssertions = BuildStorageAssertions(snapshot, targetPagePath, dependencyActions, expectedContentDigest, targetLifecycle),
                BrowserAssertions = BrowserAssertions.ToList(),
                Blockers = blockers.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()
            };
            var planDigest = EnterpriseWikiPackageSerializer.ComputePlanDigest(plan);
            var state = plan.IsExecutable ? EnterpriseWikiPackageState.ApprovalReady : EnterpriseWikiPackageState.Blocked;
            var report = new EnterpriseWikiCustomerReport
            {
                Summary = plan.IsExecutable
                    ? "Source export and target analysis completed. Import requires explicit approval of the sealed plan digest."
                    : "The package is sealed for review but cannot be imported until every blocker is resolved and a new plan is generated.",
                CapturedIngredients = new List<string>
                {
                    "Page/file/list item identity and source stability fence",
                    "Enterprise Wiki content type and EnterpriseWiki.aspx layout",
                    $"All {snapshot.Fields.Count} source Pages-library field definitions and returned values",
                    $"{snapshot.WebParts.Count} shared Web Part export(s) with zone placement",
                    $"{snapshot.Dependencies.Count} authored dependency/link snapshot(s)",
                    "Page security inheritance and source lifecycle evidence",
                    "Target publishing library, versioning, lifecycle, field, layout, and create-only probes"
                },
                Blockers = plan.Blockers.ToList(),
                Warnings = plan.Warnings.ToList()
            };

            return new EnterpriseWikiMigrationPackage
            {
                PlannedAtUtc = DateTimeOffset.UtcNow,
                ExportedAtUtc = exportPackage.ExportedAtUtc,
                State = state,
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = exportPackage.SnapshotDigest,
                PlanDigest = planDigest,
                Report = report
            };
        }

        public EnterpriseWikiImportReceipt Import(
            ClientContext targetContext,
            EnterpriseWikiMigrationPackage package,
            string approvedPlanDigest)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            EnterpriseWikiPackageSerializer.ValidateMigration(package);
            ValidateImportPlan(package);
            if (package.State != EnterpriseWikiPackageState.ApprovalReady || !package.Plan.IsExecutable)
            {
                throw new InvalidOperationException("The Enterprise Wiki package is not approval-ready.");
            }

            if (string.IsNullOrWhiteSpace(approvedPlanDigest)
                || !string.Equals(approvedPlanDigest, package.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The approved plan digest does not match the sealed Enterprise Wiki package.");
            }

            var startedAt = DateTimeOffset.UtcNow;
            var targetWeb = targetContext.Web;
            targetContext.Load(targetWeb, web => web.Url, web => web.ServerRelativeUrl);
            targetContext.ExecuteQueryRetry();
            if (!UriEquals(targetWeb.Url, package.Plan.TargetWebUrl))
            {
                throw new InvalidOperationException($"The target connection points to '{targetWeb.Url}', but the approved plan targets '{package.Plan.TargetWebUrl}'.");
            }

            var preflightBlockers = new List<string>();
            var preflightWarnings = new List<string>();
            var freshProbe = ProbeTarget(
                targetContext,
                package.Plan.TargetPageServerRelativeUrl,
                package.Plan.DependencyActions,
                package.Plan.TargetLifecycle,
                preflightBlockers,
                preflightWarnings);
            if (preflightBlockers.Count > 0)
            {
                throw new InvalidOperationException("Fresh target preflight failed: " + string.Join(" ", preflightBlockers));
            }

            if (!string.Equals(freshProbe.EnterpriseWikiContentTypeId, package.Plan.TargetProbe.EnterpriseWikiContentTypeId, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(freshProbe.EnterpriseWikiLayoutUrl, package.Plan.TargetProbe.EnterpriseWikiLayoutUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The target Enterprise Wiki content type or layout changed after approval.");
            }

            var materialized = MaterializeDependencies(targetContext, package.Snapshot.Dependencies, package.Plan.DependencyActions);
            var rewrittenContent = RewriteContent(package.Snapshot.PublishingPageContent, package.Plan.Replacements);
            var pages = targetWeb.GetPagesLibrary();
            targetContext.Load(pages, list => list.EnableModeration, list => list.ForceCheckout);
            targetContext.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            targetContext.ExecuteQueryRetry();
            var targetDirectory = GetDirectoryName(package.Plan.TargetPageServerRelativeUrl);
            if (!string.Equals(targetDirectory, pages.RootFolder.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new NotSupportedException("The Enterprise Wiki import profile supports pages in the root of the target Pages library only.");
            }

            var targetFileName = GetFileName(package.Plan.TargetPageServerRelativeUrl);
            targetWeb.AddPublishingPage(targetFileName, package.Plan.PageLayoutName, package.Snapshot.Source.Title, false);
            var targetFile = targetWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
            var targetItem = targetFile.ListItemAllFields;
            targetContext.Load(targetFile, file => file.Exists, file => file.CheckOutType);
            targetContext.Load(targetItem, item => item.Id);
            targetContext.ExecuteQueryRetry();
            if (pages.ForceCheckout && targetFile.CheckOutType == CheckOutType.None)
            {
                targetFile.CheckOut();
                targetContext.ExecuteQueryRetry();
            }

            targetItem["Title"] = package.Snapshot.Source.Title;
            targetItem["PublishingPageContent"] = rewrittenContent;
            targetItem.Update();
            targetContext.ExecuteQueryRetry();
            var fieldResults = ApplyPlannedFields(
                targetContext,
                targetItem,
                package.Snapshot.Fields,
                package.Plan.FieldActions,
                package.Plan.Replacements,
                preflightWarnings);

            foreach (var webPart in package.Snapshot.WebParts)
            {
                var entity = new WebPartEntity
                {
                    WebPartIndex = webPart.ZoneIndex,
                    WebPartTitle = webPart.Title,
                    WebPartZone = webPart.ZoneId,
                    WebPartXml = RewriteContent(webPart.ExportXml, package.Plan.Replacements)
                };
                targetWeb.AddWebPartToWebPartPage(package.Plan.TargetPageServerRelativeUrl, entity);
            }

            targetContext.Load(targetFile, file => file.CheckOutType);
            targetContext.ExecuteQueryRetry();
            var plannedFieldFailure = fieldResults.Any(result => result.Attempted && !result.Succeeded);
            if (targetFile.CheckOutType != CheckOutType.None)
            {
                var checkinType = package.Plan.TargetLifecycle == EnterpriseWikiTargetLifecycle.Published && !plannedFieldFailure
                    ? CheckinType.MajorCheckIn
                    : CheckinType.MinorCheckIn;
                targetFile.CheckIn("PnP Enterprise Wiki import", checkinType);
                targetContext.ExecuteQueryRetry();
            }

            if (package.Plan.TargetLifecycle == EnterpriseWikiTargetLifecycle.Published && !plannedFieldFailure)
            {
                targetFile.Publish("PnP Enterprise Wiki import");
                targetContext.ExecuteQueryRetry();
                if (pages.EnableModeration)
                {
                    targetFile.Approve("PnP Enterprise Wiki import");
                    targetContext.ExecuteQueryRetry();
                }
            }
            else if (plannedFieldFailure)
            {
                preflightWarnings.Add("One or more planned field updates failed. The page was not published.");
            }

            return ReadImportReceipt(
                targetContext,
                package,
                approvedPlanDigest,
                startedAt,
                materialized,
                fieldResults,
                preflightWarnings);
        }

        public static bool IsEnterpriseWikiContentType(string contentTypeId)
        {
            if (string.IsNullOrWhiteSpace(contentTypeId))
            {
                return false;
            }

            return contentTypeId.StartsWith(BuiltInContentTypeId.EnterpriseWikiPage, StringComparison.OrdinalIgnoreCase)
                && !contentTypeId.StartsWith(BuiltInContentTypeId.ProjectPage, StringComparison.OrdinalIgnoreCase);
        }

        private static void ValidateImportPlan(EnterpriseWikiMigrationPackage package)
        {
            if (package.Plan.Operation != EnterpriseWikiMigrationOperation.CreatePage || !package.Plan.CreateOnly)
            {
                throw new NotSupportedException($"Migration operation '{package.Plan.Operation}' is not executable by this importer.");
            }

            var derivedLifecycle = DeriveTargetLifecycle(package.Snapshot.Lifecycle);
            if (package.Plan.TargetLifecycle != derivedLifecycle)
            {
                throw new InvalidDataException($"Planned lifecycle '{package.Plan.TargetLifecycle}' does not match the source-derived lifecycle '{derivedLifecycle}'.");
            }

            var fieldByName = package.Snapshot.Fields.ToDictionary(item => item.InternalName, StringComparer.OrdinalIgnoreCase);
            foreach (var action in package.Plan.FieldActions.Where(item => item.Disposition == EnterpriseWikiFieldDisposition.Apply))
            {
                if (!AdditionalFieldNames.Contains(action.SourceInternalName)
                    || !string.Equals(action.SourceInternalName, action.TargetInternalName, StringComparison.OrdinalIgnoreCase)
                    || !fieldByName.TryGetValue(action.SourceInternalName, out var field)
                    || field.ReadOnly
                    || !field.HasValue
                    || field.CaptureStatus != EnterpriseWikiCaptureStatus.Captured
                    || !IsImportableFieldKind(field.Kind))
                {
                    throw new InvalidDataException($"Field action '{action.SourceInternalName}' is marked Apply but is not supported by the current importer.");
                }
            }
        }

        public static string RewriteContent(string value, IEnumerable<EnterpriseWikiTextReplacement> replacements)
        {
            var result = value ?? string.Empty;
            foreach (var replacement in (replacements ?? Array.Empty<EnterpriseWikiTextReplacement>())
                         .Where(item => !string.IsNullOrEmpty(item.Source))
                         .OrderByDescending(item => item.Source.Length))
            {
                result = Regex.Replace(
                    result,
                    Regex.Escape(replacement.Source),
                    _ => replacement.Target ?? string.Empty,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            return result;
        }

        internal static string GetWebPartMigrationBlocker(string exportXml)
        {
            if (string.IsNullOrWhiteSpace(exportXml))
            {
                return "the shared Web Part export is empty";
            }

            XDocument document;
            try
            {
                document = XDocument.Parse(exportXml, LoadOptions.None);
            }
            catch (Exception exception) when (exception is System.Xml.XmlException || exception is ArgumentException)
            {
                return $"the shared Web Part export is not valid XML ({exception.Message})";
            }

            var typeName = document
                .Descendants()
                .Where(element => string.Equals(element.Name.LocalName, "type", StringComparison.OrdinalIgnoreCase))
                .Select(element => (string)element.Attribute("name"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (string.IsNullOrWhiteSpace(typeName))
            {
                return "the shared Web Part export does not declare a v3 type";
            }

            if (typeName.EndsWith(".ErrorWebPart", StringComparison.OrdinalIgnoreCase))
            {
                return $"type '{typeName}' represents a Web Part that is already unavailable on the source page";
            }

            if (typeName.EndsWith(".RSSAggregatorWebPart", StringComparison.OrdinalIgnoreCase))
            {
                return $"type '{typeName}' is not supported by the current deterministic import profile and must be replaced or explicitly mapped";
            }

            var properties = document
                .Descendants()
                .Where(element => string.Equals(element.Name.LocalName, "property", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            var hasSourceListBinding = properties.Any(property =>
            {
                var name = (string)property.Attribute("name");
                if (!string.Equals(name, "ListId", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(name, "ListName", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var value = property.Value?.Trim().Trim('{', '}');
                return !string.IsNullOrWhiteSpace(value)
                    && !string.Equals(value, Guid.Empty.ToString(), StringComparison.OrdinalIgnoreCase);
            });
            if (hasSourceListBinding
                || typeName.EndsWith(".XsltListViewWebPart", StringComparison.OrdinalIgnoreCase)
                || typeName.EndsWith(".ListViewWebPart", StringComparison.OrdinalIgnoreCase))
            {
                return $"type '{typeName}' is bound to a source list; the exact profile requires a reviewed target-list and view mapping before import";
            }

            return null;
        }

        private static void ValidateExportOptions(EnterpriseWikiExportOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.SourcePageServerRelativeUrl))
            {
                throw new ArgumentException("A source page path is required.", nameof(options));
            }

            if (options.MaximumDependencyBytes <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(options), "MaximumDependencyBytes must be greater than zero.");
            }
        }

        private static void ValidatePlanningOptions(EnterpriseWikiPlanningOptions options)
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
                throw new NotSupportedException("Only CreatePage plans are supported. Deferred-field recovery remains represented by the package schema but is not executable yet.");
            }
        }

        private static SourceCapture CaptureSourcePage(
            ClientContext context,
            string pagePath,
            EnterpriseWikiExportOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var web = context.Web;
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            var item = file.ListItemAllFields;
            var contentType = item.ContentType;
            context.Load(web, value => value.Url, value => value.ServerRelativeUrl);
            context.Load(file,
                value => value.Exists,
                value => value.UniqueId,
                value => value.ServerRelativeUrl,
                value => value.UIVersionLabel,
                value => value.Length,
                value => value.TimeLastModified,
                value => value.CheckOutType,
                value => value.Level,
                value => value.TimeCreated);
            context.Load(item);
            context.Load(contentType, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            context.Load(item, value => value.Id, value => value.HasUniqueRoleAssignments);
            context.ExecuteQueryRetry();
            if (!file.Exists)
            {
                throw new FileNotFoundException("The source Enterprise Wiki page was not found.", pagePath);
            }

            var contentTypeId = contentType.Id.StringValue;
            if (!IsEnterpriseWikiContentType(contentTypeId))
            {
                blockers.Add($"Source ContentTypeId '{contentTypeId}' is not an Enterprise Wiki Page content type (Project Page is intentionally excluded).");
            }

            var content = GetFieldString(item, "PublishingPageContent") ?? string.Empty;
            var layout = item.FieldValues.TryGetValue("PublishingPageLayout", out var layoutValue)
                ? layoutValue as FieldUrlValue
                : null;
            if (string.IsNullOrWhiteSpace(content))
            {
                warnings.Add("PublishingPageContent is empty.");
            }

            if (layout == null || string.IsNullOrWhiteSpace(layout.Url))
            {
                blockers.Add("PublishingPageLayout is unavailable on the source page.");
            }

            var identity = new EnterpriseWikiPageIdentity
            {
                WebUrl = web.Url.TrimEnd('/'),
                WebServerRelativeUrl = web.ServerRelativeUrl,
                PageServerRelativeUrl = file.ServerRelativeUrl,
                ListItemId = item.Id,
                FileUniqueId = file.UniqueId,
                ContentTypeId = contentTypeId,
                ContentTypeName = contentType.Name,
                VersionLabel = file.UIVersionLabel,
                Length = file.Length,
                ModifiedUtc = file.TimeLastModified.ToUniversalTime(),
                Title = GetFieldString(item, "Title") ?? GetFileName(pagePath),
                PageLayoutUrl = layout?.Url ?? string.Empty,
                PageLayoutDescription = layout?.Description
            };
            var fields = CaptureFields(context, item, blockers, warnings);
            var webParts = options.IncludeWebParts
                ? CaptureWebParts(web, pagePath, blockers)
                : new List<EnterpriseWikiWebPartSnapshot>();
            var security = CaptureSecurity(context, item, warnings);
            var moderationStatus = TryGetInt32(item, "_ModerationStatus");
            var lifecycle = new EnterpriseWikiLifecycleSnapshot
            {
                CheckOutType = file.CheckOutType.ToString(),
                Level = file.Level.ToString(),
                ModerationStatus = moderationStatus,
                CreatedUtc = file.TimeCreated.ToUniversalTime(),
                ModifiedUtc = file.TimeLastModified.ToUniversalTime()
            };
            return new SourceCapture
            {
                Identity = identity,
                PublishingPageContent = content,
                Fields = fields,
                WebParts = webParts,
                Security = security,
                Lifecycle = lifecycle,
                SourceFence = ToFence(file)
            };
        }

        private static List<EnterpriseWikiFieldValueSnapshot> CaptureFields(
            ClientContext context,
            ListItem item,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var parentList = item.ParentList;
            context.Load(parentList.Fields, fields => fields.Include(
                field => field.Id,
                field => field.InternalName,
                field => field.Title,
                field => field.TypeAsString,
                field => field.SchemaXml,
                field => field.ReadOnlyField,
                field => field.Hidden,
                field => field.Required));
            context.ExecuteQueryRetry();
            var result = new List<EnterpriseWikiFieldValueSnapshot>();
            foreach (var field in parentList.Fields.OrderBy(value => value.InternalName, StringComparer.Ordinal))
            {
                if (!item.FieldValues.TryGetValue(field.InternalName, out var value))
                {
                    result.Add(CreateFieldSnapshot(field, EnterpriseWikiCaptureStatus.NotReturned, EnterpriseWikiFieldValueKind.Null));
                    continue;
                }

                try
                {
                    var snapshot = SerializeFieldValue(field, value);
                    result.Add(snapshot);
                    if (snapshot.CaptureStatus == EnterpriseWikiCaptureStatus.CapturedWithLimitations)
                    {
                        warnings.Add($"Field '{field.InternalName}' has value type '{snapshot.RawType}' that was captured as recovery evidence only.");
                    }
                }
                catch (Exception exception)
                {
                    var snapshot = CreateFieldSnapshot(field, EnterpriseWikiCaptureStatus.Failed, EnterpriseWikiFieldValueKind.Unsupported);
                    snapshot.HasValue = value != null;
                    snapshot.RawType = value?.GetType().FullName;
                    snapshot.RawValue = SafeConvertToString(value);
                    snapshot.Diagnostics.Add(exception.Message);
                    result.Add(snapshot);
                    warnings.Add($"Field '{field.InternalName}' could not be fully serialized and remains in the snapshot with diagnostics: {exception.Message}");
                }
            }

            return result;
        }

        private static EnterpriseWikiFieldValueSnapshot SerializeFieldValue(Field field, object value)
        {
            var snapshot = CreateFieldSnapshot(field, EnterpriseWikiCaptureStatus.Captured, EnterpriseWikiFieldValueKind.Null);
            snapshot.HasValue = value != null;
            snapshot.RawType = value?.GetType().FullName;
            snapshot.RawValue = SafeConvertToString(value);
            snapshot.RawValueJson = TrySerializeRawValue(value, snapshot.Diagnostics);
            if (value == null)
            {
                return snapshot;
            }

            if (value is FieldUrlValue url)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.Url;
                snapshot.UrlValue = new EnterpriseWikiUrlValueSnapshot
                {
                    Url = url.Url,
                    Description = url.Description
                };
            }
            else if (value is TaxonomyFieldValueCollection taxonomyCollection)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.TaxonomyCollection;
                snapshot.TaxonomyValues = taxonomyCollection.Select(ToTaxonomyValue).ToList();
            }
            else if (value is TaxonomyFieldValue taxonomy)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.Taxonomy;
                snapshot.TaxonomyValues.Add(ToTaxonomyValue(taxonomy));
            }
            else if (value is FieldUserValue[] users)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.UserCollection;
                snapshot.LookupValues = users.Select(ToLookupValue).ToList();
            }
            else if (value is FieldUserValue user)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.User;
                snapshot.LookupValues.Add(ToLookupValue(user));
            }
            else if (value is FieldLookupValue[] lookups)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.LookupCollection;
                snapshot.LookupValues = lookups.Select(ToLookupValue).ToList();
            }
            else if (value is FieldLookupValue lookup)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.Lookup;
                snapshot.LookupValues.Add(ToLookupValue(lookup));
            }
            else if (value is DateTime dateTime)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.DateTime;
                snapshot.Value = dateTime.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }
            else if (value is bool boolean)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.Boolean;
                snapshot.Value = boolean ? "true" : "false";
            }
            else if (value is byte || value is short || value is int || value is long || value is float || value is double || value is decimal)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.Number;
                snapshot.Value = Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            else if (value is Guid guid)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.Guid;
                snapshot.Value = guid.ToString("D");
            }
            else if (value is byte[] bytes)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.ByteArray;
                snapshot.BinaryBase64 = Convert.ToBase64String(bytes);
            }
            else if (value is string[] strings)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.StringCollection;
                snapshot.StringValues = strings.ToList();
            }
            else if (value is string text)
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.String;
                snapshot.Value = text;
            }
            else
            {
                snapshot.Kind = EnterpriseWikiFieldValueKind.Unsupported;
                snapshot.CaptureStatus = EnterpriseWikiCaptureStatus.CapturedWithLimitations;
                snapshot.Diagnostics.Add("No typed importer is registered for this runtime value. Raw type, text, and best-effort JSON are retained for future recovery.");
            }

            return snapshot;
        }

        private static EnterpriseWikiFieldValueSnapshot CreateFieldSnapshot(
            Field field,
            EnterpriseWikiCaptureStatus captureStatus,
            EnterpriseWikiFieldValueKind kind)
        {
            return new EnterpriseWikiFieldValueSnapshot
            {
                Id = field.Id,
                InternalName = field.InternalName,
                Title = field.Title,
                TypeAsString = field.TypeAsString,
                SchemaXml = field.SchemaXml,
                ReadOnly = field.ReadOnlyField,
                Hidden = field.Hidden,
                Required = field.Required,
                Kind = kind,
                CaptureStatus = captureStatus
            };
        }

        private static EnterpriseWikiLookupValueSnapshot ToLookupValue(FieldLookupValue value)
        {
            return new EnterpriseWikiLookupValueSnapshot
            {
                LookupId = value.LookupId,
                LookupValue = value.LookupValue
            };
        }

        private static EnterpriseWikiTaxonomyValueSnapshot ToTaxonomyValue(TaxonomyFieldValue value)
        {
            return new EnterpriseWikiTaxonomyValueSnapshot
            {
                Label = value.Label,
                TermGuid = value.TermGuid,
                WssId = value.WssId
            };
        }

        private static string SafeConvertToString(object value)
        {
            try
            {
                return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }

        private static string TrySerializeRawValue(object value, ICollection<string> diagnostics)
        {
            if (value == null)
            {
                return null;
            }

            try
            {
                return JsonSerializer.Serialize(value, value.GetType());
            }
            catch (Exception exception) when (exception is NotSupportedException || exception is JsonException)
            {
                diagnostics.Add($"Best-effort raw JSON serialization was unavailable: {exception.Message}");
                return null;
            }
        }

        private static List<EnterpriseWikiWebPartSnapshot> CaptureWebParts(Web web, string pagePath, ICollection<string> blockers)
        {
            var result = new List<EnterpriseWikiWebPartSnapshot>();
            var context = (ClientContext)web.Context;
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            var manager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
            IEnumerable<WebPartDefinition> webParts;
            try
            {
                webParts = web.GetWebParts(pagePath).ToArray();
            }
            catch (ServerException exception)
            {
                blockers.Add($"Shared Web Part inventory could not be captured: {exception.Message}");
                return result;
            }

            foreach (var webPart in webParts)
            {
                string xml;
                try
                {
                    var export = manager.ExportWebPart(webPart.Id);
                    context.ExecuteQueryRetry();
                    xml = export.Value;
                }
                catch (ServerException exception)
                {
                    blockers.Add($"Web Part '{webPart.Id}' could not be exported: {exception.Message}");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(xml))
                {
                    blockers.Add($"Web Part '{webPart.Id}' returned an empty export.");
                    continue;
                }

                var snapshot = new EnterpriseWikiWebPartSnapshot
                {
                    Id = webPart.Id,
                    Title = webPart.WebPart.Title,
                    ZoneId = webPart.ZoneId,
                    ZoneIndex = webPart.WebPart.ZoneIndex,
                    Hidden = webPart.WebPart.Hidden,
                    ExportXml = xml,
                    ExportSha256 = EnterpriseWikiPackageSerializer.ComputeSha256(xml)
                };
                result.Add(snapshot);
                var migrationBlocker = GetWebPartMigrationBlocker(xml);
                if (!string.IsNullOrWhiteSpace(migrationBlocker))
                {
                    var title = string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Id.ToString() : snapshot.Title;
                    blockers.Add($"Web Part '{title}' ({snapshot.Id}) cannot be copied: {migrationBlocker}.");
                }
            }

            return result;
        }

        private static EnterpriseWikiSecuritySnapshot CaptureSecurity(
            ClientContext context,
            ListItem item,
            ICollection<string> warnings)
        {
            var result = new EnterpriseWikiSecuritySnapshot
            {
                HasUniqueRoleAssignments = item.HasUniqueRoleAssignments
            };
            if (!item.HasUniqueRoleAssignments)
            {
                return result;
            }

            try
            {
                var assignments = item.RoleAssignments;
                context.Load(assignments);
                context.ExecuteQueryRetry();
                foreach (var assignment in assignments)
                {
                    context.Load(assignment.Member, member => member.LoginName, member => member.Title);
                    context.Load(assignment.RoleDefinitionBindings, definitions => definitions.Include(definition => definition.Name));
                }

                context.ExecuteQueryRetry();
                foreach (var assignment in assignments)
                {
                    result.RoleAssignments.Add(new EnterpriseWikiRoleAssignmentSnapshot
                    {
                        PrincipalLoginName = assignment.Member.LoginName,
                        PrincipalTitle = assignment.Member.Title,
                        RoleDefinitionNames = assignment.RoleDefinitionBindings
                            .Select(definition => definition.Name)
                            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                            .ToList()
                    });
                }
            }
            catch (Exception exception) when (IsAccessDenied(exception))
            {
                const string diagnostic = "The source page has unique permissions, but the current principal cannot enumerate its role assignments. Permission replay is not supported by this migration profile, so page capture continued without ACL details.";
                warnings.Add(diagnostic);
            }

            return result;
        }

        private static bool IsAccessDenied(Exception exception)
        {
            for (var current = exception; current != null; current = current.InnerException)
            {
                if (current is UnauthorizedAccessException || current is ServerUnauthorizedAccessException)
                {
                    return true;
                }

                if (current is ServerException serverException && serverException.ServerErrorCode == -2147024891)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<EnterpriseWikiDependencySnapshot> CaptureDependencies(
            ClientContext sourceContext,
            EnterpriseWikiPageIdentity source,
            string publishingPageContent,
            IEnumerable<EnterpriseWikiWebPartSnapshot> webParts,
            EnterpriseWikiExportOptions options,
            ICollection<string> warnings)
        {
            var candidates = ExtractDependencyCandidates(publishingPageContent);
            foreach (var webPart in webParts)
            {
                candidates.AddRange(ExtractUrlCandidates(webPart.ExportXml, $"webpart:{webPart.Id}"));
            }

            var sourceWebUri = new Uri(EnsureTrailingSlash(source.WebUrl));
            var sourcePageUri = new Uri(sourceWebUri.GetLeftPart(UriPartial.Authority) + EncodePath(source.PageServerRelativeUrl));
            var result = new List<EnterpriseWikiDependencySnapshot>();
            foreach (var candidate in candidates
                         .GroupBy(item => $"{item.Consumer}\n{item.Value}", StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                if (!TryResolveDependencyUri(sourcePageUri, candidate.Value, out var absoluteUri))
                {
                    continue;
                }

                var dependency = CaptureDependency(sourceContext, sourceWebUri, candidate, absoluteUri, options, warnings);
                result.Add(dependency);
            }

            return result;
        }

        private static EnterpriseWikiDependencySnapshot CaptureDependency(
            ClientContext sourceContext,
            Uri sourceWebUri,
            DependencyCandidate candidate,
            Uri absoluteUri,
            EnterpriseWikiExportOptions options,
            ICollection<string> warnings)
        {
            var dependency = new EnterpriseWikiDependencySnapshot
            {
                Id = EnterpriseWikiPackageSerializer.ComputeSha256($"{candidate.Consumer}\n{absoluteUri.AbsoluteUri}"),
                OriginalValue = candidate.Value,
                SourceAbsoluteUrl = absoluteUri.AbsoluteUri,
                Consumer = candidate.Consumer,
                Kind = candidate.Kind,
                IsRenderableResource = candidate.IsRenderableResource,
                CaptureStatus = EnterpriseWikiCaptureStatus.Captured
            };
            if (!string.Equals(sourceWebUri.Host, absoluteUri.Host, StringComparison.OrdinalIgnoreCase))
            {
                return dependency;
            }

            var sourcePath = Uri.UnescapeDataString(absoluteUri.AbsolutePath);
            var sourceWebPath = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/');
            dependency.SourceServerRelativeUrl = sourcePath;
            if (!candidate.IsRenderableResource || IsSharePointRuntimePath(sourcePath))
            {
                return dependency;
            }

            if (candidate.Kind == EnterpriseWikiDependencyKind.IFrame)
            {
                dependency.Diagnostics.Add("Same-tenant iframe dependencies require a separately reviewed page/application profile during planning.");
                return dependency;
            }

            if (!IsPathWithin(sourcePath, sourceWebPath))
            {
                dependency.Diagnostics.Add("The resource is outside the captured source web boundary.");
                return dependency;
            }

            try
            {
                var payload = CaptureFilePayload(sourceContext, sourcePath, options.MaximumDependencyBytes);
                dependency.ContentBase64 = Convert.ToBase64String(payload);
                dependency.ContentLength = payload.LongLength;
                dependency.ContentSha256 = ComputeBytesSha256(payload);
            }
            catch (Exception exception) when (exception is ServerException || exception is InvalidOperationException || exception is IOException)
            {
                dependency.CaptureStatus = EnterpriseWikiCaptureStatus.Failed;
                dependency.Diagnostics.Add(exception.Message);
                warnings.Add($"Resource '{absoluteUri}' could not be captured and may block a later plan: {exception.Message}");
            }

            return dependency;
        }

        private static byte[] CaptureFilePayload(ClientContext context, string serverRelativeUrl, long maximumBytes)
        {
            var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            context.Load(file, value => value.Exists, value => value.Length);
            var stream = file.OpenBinaryStream();
            context.ExecuteQueryRetry();
            if (!file.Exists || stream.Value == null)
            {
                throw new FileNotFoundException("The referenced SharePoint file was not found.", serverRelativeUrl);
            }

            if (file.Length > maximumBytes)
            {
                throw new InvalidOperationException($"The dependency is {file.Length} bytes, above the configured {maximumBytes}-byte limit.");
            }

            using (stream.Value)
            using (var output = new MemoryStream())
            {
                stream.Value.CopyTo(output);
                if (output.Length > maximumBytes)
                {
                    throw new InvalidOperationException($"The dependency is above the configured {maximumBytes}-byte limit.");
                }
                return output.ToArray();
            }
        }

        private static List<EnterpriseWikiDependencyAction> BuildDependencyActions(
            EnterpriseWikiSnapshot snapshot,
            string targetWebUrl,
            string targetWebServerRelativeUrl,
            EnterpriseWikiPlanningOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var sourceWebUri = new Uri(EnsureTrailingSlash(snapshot.Source.WebUrl));
            var targetWebUri = new Uri(EnsureTrailingSlash(targetWebUrl));
            var sourceWebPath = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/');
            var targetWebPath = targetWebServerRelativeUrl.TrimEnd('/');
            var result = new List<EnterpriseWikiDependencyAction>();
            foreach (var dependency in snapshot.Dependencies)
            {
                var action = new EnterpriseWikiDependencyAction
                {
                    SnapshotDependencyId = dependency.Id,
                    Disposition = EnterpriseWikiDependencyDisposition.PreserveExternal
                };
                result.Add(action);
                if (!Uri.TryCreate(dependency.SourceAbsoluteUrl, UriKind.Absolute, out var sourceUri))
                {
                    action.Disposition = EnterpriseWikiDependencyDisposition.Block;
                    action.Diagnostics.Add("The captured dependency URL is not an absolute HTTP(S) URL.");
                    blockers.Add($"Dependency '{dependency.OriginalValue}' has an invalid captured URL.");
                    continue;
                }

                if (!string.Equals(sourceWebUri.Host, sourceUri.Host, StringComparison.OrdinalIgnoreCase))
                {
                    if (dependency.IsRenderableResource && !options.AllowExternalResourceReferences)
                    {
                        action.Disposition = EnterpriseWikiDependencyDisposition.Block;
                        action.Diagnostics.Add("External renderable resources are blocked by planning policy.");
                        blockers.Add($"External resource '{sourceUri}' is blocked by policy.");
                    }
                    continue;
                }

                var sourcePath = dependency.SourceServerRelativeUrl ?? Uri.UnescapeDataString(sourceUri.AbsolutePath);
                var targetPath = IsPathWithin(sourcePath, sourceWebPath)
                    ? targetWebPath + sourcePath.Substring(sourceWebPath.Length)
                    : sourcePath;
                action.TargetServerRelativeUrl = targetPath;
                action.TargetAbsoluteUrl = targetWebUri.GetLeftPart(UriPartial.Authority) + EncodePath(targetPath) + sourceUri.Query + sourceUri.Fragment;
                action.Disposition = EnterpriseWikiDependencyDisposition.RewriteToTarget;
                if (!dependency.IsRenderableResource || IsSharePointRuntimePath(sourcePath))
                {
                    continue;
                }

                if (dependency.Kind == EnterpriseWikiDependencyKind.IFrame)
                {
                    action.Disposition = EnterpriseWikiDependencyDisposition.Block;
                    action.Diagnostics.Add("Same-tenant iframe dependencies require a separately reviewed page/application profile.");
                    blockers.Add($"Iframe dependency '{sourceUri}' is unsupported by the exact profile.");
                    continue;
                }

                if (!IsPathWithin(sourcePath, sourceWebPath))
                {
                    action.Disposition = EnterpriseWikiDependencyDisposition.Block;
                    action.Diagnostics.Add("The resource is outside the captured source web and cannot be safely materialized inside the approved target web.");
                    blockers.Add($"Same-tenant resource '{sourceUri}' is outside the source web boundary.");
                    continue;
                }

                if (dependency.CaptureStatus == EnterpriseWikiCaptureStatus.Failed
                    || string.IsNullOrWhiteSpace(dependency.ContentBase64)
                    || string.IsNullOrWhiteSpace(dependency.ContentSha256))
                {
                    action.Disposition = EnterpriseWikiDependencyDisposition.Block;
                    action.Diagnostics.Add("The source payload was not captured successfully.");
                    blockers.Add($"Resource '{sourceUri}' has no restorable payload in the source snapshot.");
                    continue;
                }

                action.Disposition = EnterpriseWikiDependencyDisposition.MaterializeAtTarget;
            }

            return result
                .OrderBy(action => action.SnapshotDependencyId, StringComparer.Ordinal)
                .ToList();
        }

        private static List<EnterpriseWikiFieldAction> BuildFieldActions(
            ClientContext targetContext,
            IEnumerable<EnterpriseWikiFieldValueSnapshot> fields,
            EnterpriseWikiPlanningOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var pages = targetContext.Web.GetPagesLibrary();
            var sourceFields = fields.OrderBy(field => field.InternalName, StringComparer.Ordinal).ToArray();
            if (pages == null)
            {
                return sourceFields.Select(field => new EnterpriseWikiFieldAction
                {
                    SourceInternalName = field.InternalName,
                    TargetInternalName = field.InternalName,
                    Disposition = EnterpriseWikiFieldDisposition.Block,
                    Reason = "The target publishing Pages library is unavailable."
                }).ToList();
            }

            targetContext.Load(pages.Fields, values => values.Include(
                field => field.InternalName,
                field => field.TypeAsString,
                field => field.ReadOnlyField));
            targetContext.ExecuteQueryRetry();
            var targetFields = pages.Fields.ToDictionary(field => field.InternalName, StringComparer.OrdinalIgnoreCase);
            var result = new List<EnterpriseWikiFieldAction>();
            foreach (var sourceField in sourceFields)
            {
                var action = new EnterpriseWikiFieldAction
                {
                    SourceInternalName = sourceField.InternalName,
                    TargetInternalName = sourceField.InternalName
                };
                result.Add(action);
                if (HandledFieldNames.Contains(sourceField.InternalName))
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.AlreadyHandled;
                    action.Reason = "The page creation workflow handles this field explicitly.";
                    continue;
                }

                if (sourceField.CaptureStatus == EnterpriseWikiCaptureStatus.Failed
                    || sourceField.CaptureStatus == EnterpriseWikiCaptureStatus.NotReturned)
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.CaptureUnavailable;
                    action.Reason = "The field definition is preserved, but no restorable value was captured.";
                    continue;
                }

                if (!sourceField.HasValue || sourceField.Kind == EnterpriseWikiFieldValueKind.Null)
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.SkipEmpty;
                    action.Reason = "The source item has no value for this field.";
                    continue;
                }

                if (!AdditionalFieldNames.Contains(sourceField.InternalName))
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.EvidenceOnly;
                    action.Reason = "The field is fully retained in the snapshot, but this importer does not recognize it yet.";
                    continue;
                }

                if (sourceField.ReadOnly)
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.SkipReadOnly;
                    action.Reason = "The source field is read-only.";
                    continue;
                }

                if (string.Equals(sourceField.TypeAsString, "Calculated", StringComparison.OrdinalIgnoreCase))
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.SkipCalculated;
                    action.Reason = "Calculated fields are recomputed by SharePoint.";
                    continue;
                }

                if (!targetFields.TryGetValue(sourceField.InternalName, out var targetField))
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.TargetFieldMissing;
                    action.Reason = "The recognized field is not present in the target Pages library.";
                    warnings.Add($"Recognized field '{sourceField.InternalName}' is absent from the target Pages library and will not be applied.");
                    continue;
                }

                action.TargetInternalName = targetField.InternalName;
                action.TargetTypeAsString = targetField.TypeAsString;
                if (targetField.ReadOnlyField)
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.SkipReadOnly;
                    action.Reason = "The target field is read-only.";
                    continue;
                }

                if (!string.Equals(sourceField.TypeAsString, targetField.TypeAsString, StringComparison.OrdinalIgnoreCase))
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.TargetTypeMismatch;
                    action.Reason = $"Source type '{sourceField.TypeAsString}' does not match target type '{targetField.TypeAsString}'.";
                    warnings.Add($"Recognized field '{sourceField.InternalName}' has a target type mismatch and will not be applied.");
                    continue;
                }

                if (sourceField.Kind == EnterpriseWikiFieldValueKind.Taxonomy
                    || sourceField.Kind == EnterpriseWikiFieldValueKind.TaxonomyCollection
                    || sourceField.Kind == EnterpriseWikiFieldValueKind.User
                    || sourceField.Kind == EnterpriseWikiFieldValueKind.UserCollection
                    || sourceField.Kind == EnterpriseWikiFieldValueKind.Lookup
                    || sourceField.Kind == EnterpriseWikiFieldValueKind.LookupCollection)
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.RequiresMapping;
                    action.Reason = "The value is captured, but its source identity must be mapped before it can be safely applied to another site.";
                    if (!(sourceField.Kind == EnterpriseWikiFieldValueKind.Taxonomy
                        || sourceField.Kind == EnterpriseWikiFieldValueKind.TaxonomyCollection)
                        || !options.BlockOnManagedMetadata)
                    {
                        warnings.Add($"Field '{sourceField.InternalName}' requires an identity or term mapping and remains evidence-only.");
                    }
                    continue;
                }

                if (!IsImportableFieldKind(sourceField.Kind))
                {
                    action.Disposition = EnterpriseWikiFieldDisposition.EvidenceOnly;
                    action.Reason = $"No importer is registered for value kind '{sourceField.Kind}'.";
                    continue;
                }

                action.Disposition = EnterpriseWikiFieldDisposition.Apply;
                action.Reason = "The field is recognized, writable, type-compatible, and has a supported captured value.";
            }

            return result;
        }

        private static bool IsImportableFieldKind(EnterpriseWikiFieldValueKind kind)
        {
            return kind == EnterpriseWikiFieldValueKind.String
                || kind == EnterpriseWikiFieldValueKind.StringCollection
                || kind == EnterpriseWikiFieldValueKind.Boolean
                || kind == EnterpriseWikiFieldValueKind.Number
                || kind == EnterpriseWikiFieldValueKind.DateTime
                || kind == EnterpriseWikiFieldValueKind.Guid
                || kind == EnterpriseWikiFieldValueKind.Url;
        }

        public static EnterpriseWikiTargetLifecycle DeriveTargetLifecycle(EnterpriseWikiLifecycleSnapshot sourceLifecycle)
        {
            var isPublished = string.Equals(sourceLifecycle?.Level, "Published", StringComparison.OrdinalIgnoreCase)
                && string.Equals(sourceLifecycle?.CheckOutType, "None", StringComparison.OrdinalIgnoreCase)
                && (!sourceLifecycle.ModerationStatus.HasValue || sourceLifecycle.ModerationStatus.Value == 0);
            return isPublished
                ? EnterpriseWikiTargetLifecycle.Published
                : EnterpriseWikiTargetLifecycle.Draft;
        }

        private static EnterpriseWikiTargetProbe ProbeTarget(
            ClientContext context,
            string targetPagePath,
            IEnumerable<EnterpriseWikiDependencyAction> dependencies,
            EnterpriseWikiTargetLifecycle targetLifecycle,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var web = context.Web;
            context.Load(web, value => value.Url, value => value.ServerRelativeUrl, value => value.WebTemplate, value => value.Configuration);
            context.Load(context.Site, site => site.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var probe = new EnterpriseWikiTargetProbe
            {
                WebUrl = web.Url.TrimEnd('/'),
                WebServerRelativeUrl = web.ServerRelativeUrl,
                WebTemplate = web.WebTemplate,
                WebConfiguration = web.Configuration
            };

            var pages = web.GetPagesLibrary();
            if (pages == null)
            {
                blockers.Add("The target web has no publishing Pages library.");
                return probe;
            }

            context.Load(pages,
                list => list.BaseTemplate,
                list => list.EnableVersioning,
                list => list.EnableMinorVersions,
                list => list.EnableModeration,
                list => list.ForceCheckout,
                list => list.DraftVersionVisibility);
            context.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            context.Load(pages.ContentTypes, contentTypes => contentTypes.Include(contentType => contentType.Id, contentType => contentType.Name));
            context.ExecuteQueryRetry();
            probe.PagesLibraryBaseTemplate = pages.BaseTemplate;
            probe.PagesLibraryServerRelativeUrl = pages.RootFolder.ServerRelativeUrl;
            probe.EnableVersioning = pages.EnableVersioning;
            probe.EnableMinorVersions = pages.EnableMinorVersions;
            probe.EnableModeration = pages.EnableModeration;
            probe.ForceCheckout = pages.ForceCheckout;
            probe.DraftVersionVisibility = pages.DraftVersionVisibility.ToString();
            if (pages.BaseTemplate != PublishingPagesListTemplate)
            {
                blockers.Add($"The target Pages library has base template {pages.BaseTemplate}; publishing Pages template {PublishingPagesListTemplate} is required.");
            }

            if (targetLifecycle == EnterpriseWikiTargetLifecycle.Draft
                && (!pages.EnableVersioning || !pages.EnableMinorVersions))
            {
                blockers.Add("The source maps to Draft, but the target Pages library cannot represent a checked-in minor draft deterministically.");
            }

            var enterpriseWikiContentType = pages.ContentTypes.FirstOrDefault(contentType => IsEnterpriseWikiContentType(contentType.Id.StringValue));
            probe.EnterpriseWikiContentTypeId = enterpriseWikiContentType?.Id.StringValue;
            if (enterpriseWikiContentType == null)
            {
                blockers.Add("The Enterprise Wiki Page content type is not available in the target Pages library.");
            }

            var siteRootPath = context.Site.ServerRelativeUrl == "/" ? string.Empty : context.Site.ServerRelativeUrl.TrimEnd('/');
            var layoutPath = $"{siteRootPath}/_catalogs/masterpage/EnterpriseWiki.aspx";
            var layoutFile = context.Site.RootWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(layoutPath));
            context.Load(layoutFile, file => file.Exists, file => file.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            probe.EnterpriseWikiLayoutExists = layoutFile.Exists;
            probe.EnterpriseWikiLayoutUrl = layoutFile.Exists
                ? new Uri(new Uri(web.Url).GetLeftPart(UriPartial.Authority) + EncodePath(layoutFile.ServerRelativeUrl)).AbsoluteUri
                : null;
            if (!layoutFile.Exists)
            {
                blockers.Add("EnterpriseWiki.aspx is not available in the target site collection master page gallery.");
            }

            var expectedDirectory = pages.RootFolder.ServerRelativeUrl;
            if (!string.Equals(GetDirectoryName(targetPagePath), expectedDirectory, StringComparison.OrdinalIgnoreCase))
            {
                blockers.Add($"The target page must be placed in the root of '{expectedDirectory}' for the exact profile.");
            }

            probe.TargetPageExists = FileExists(context, targetPagePath);
            if (probe.TargetPageExists)
            {
                blockers.Add($"Create-only target page already exists: {targetPagePath}");
            }

            foreach (var path in dependencies
                         .Where(dependency => dependency.Disposition == EnterpriseWikiDependencyDisposition.MaterializeAtTarget)
                         .Select(dependency => dependency.TargetServerRelativeUrl)
                         .Where(path => !string.IsNullOrWhiteSpace(path))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!IsPathWithin(path, web.ServerRelativeUrl))
                {
                    blockers.Add($"Planned dependency target escapes the target web boundary: {path}");
                    continue;
                }

                if (FileExists(context, path))
                {
                    probe.ExistingDependencyPaths.Add(path);
                    blockers.Add($"Create-only dependency target already exists: {path}");
                }
            }

            return probe;
        }

        private static int MaterializeDependencies(
            ClientContext context,
            IEnumerable<EnterpriseWikiDependencySnapshot> snapshots,
            IEnumerable<EnterpriseWikiDependencyAction> actions)
        {
            var web = context.Web;
            web.EnsureProperty(value => value.ServerRelativeUrl);
            var snapshotById = snapshots.ToDictionary(item => item.Id, StringComparer.Ordinal);
            var count = 0;
            foreach (var action in actions
                         .Where(item => item.Disposition == EnterpriseWikiDependencyDisposition.MaterializeAtTarget)
                         .GroupBy(item => item.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                if (!snapshotById.TryGetValue(action.SnapshotDependencyId, out var snapshot))
                {
                    throw new InvalidDataException($"Dependency action references missing snapshot '{action.SnapshotDependencyId}'.");
                }

                var bytes = Convert.FromBase64String(snapshot.ContentBase64 ?? string.Empty);
                if (!string.Equals(ComputeBytesSha256(bytes), snapshot.ContentSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Dependency payload digest mismatch: {action.TargetServerRelativeUrl}");
                }

                if (!IsPathWithin(action.TargetServerRelativeUrl, web.ServerRelativeUrl))
                {
                    throw new InvalidOperationException($"Dependency target escapes the target web boundary: {action.TargetServerRelativeUrl}");
                }

                var relativePath = action.TargetServerRelativeUrl.Substring(web.ServerRelativeUrl.TrimEnd('/').Length).TrimStart('/');
                var separator = relativePath.LastIndexOf('/');
                var folderPath = separator < 0 ? string.Empty : relativePath.Substring(0, separator);
                var fileName = separator < 0 ? relativePath : relativePath.Substring(separator + 1);
                var folder = string.IsNullOrEmpty(folderPath) ? web.RootFolder : web.EnsureFolderPath(folderPath);
                var created = folder.Files.Add(new FileCreationInformation
                {
                    Content = bytes,
                    Overwrite = false,
                    Url = fileName
                });
                context.Load(created, file => file.Exists, file => file.ServerRelativeUrl);
                context.ExecuteQueryRetry();
                if (!created.Exists)
                {
                    throw new InvalidOperationException($"SharePoint did not persist dependency '{action.TargetServerRelativeUrl}'.");
                }
                count++;
            }

            return count;
        }

        private static List<EnterpriseWikiFieldImportResult> ApplyPlannedFields(
            ClientContext context,
            ListItem targetItem,
            IEnumerable<EnterpriseWikiFieldValueSnapshot> fields,
            IEnumerable<EnterpriseWikiFieldAction> actions,
            IEnumerable<EnterpriseWikiTextReplacement> replacements,
            ICollection<string> warnings)
        {
            var fieldByName = fields.ToDictionary(field => field.InternalName, StringComparer.OrdinalIgnoreCase);
            var results = new List<EnterpriseWikiFieldImportResult>();
            foreach (var action in actions)
            {
                var result = new EnterpriseWikiFieldImportResult
                {
                    InternalName = action.SourceInternalName,
                    PlannedDisposition = action.Disposition,
                    Attempted = action.Disposition == EnterpriseWikiFieldDisposition.Apply,
                    Succeeded = action.Disposition != EnterpriseWikiFieldDisposition.Apply,
                    Message = action.Reason
                };
                results.Add(result);
                if (!result.Attempted)
                {
                    continue;
                }

                if (!fieldByName.TryGetValue(action.SourceInternalName, out var field))
                {
                    result.Message = "The planned source field is missing from the sealed snapshot.";
                    warnings.Add($"Planned field '{action.SourceInternalName}' is missing from the sealed snapshot.");
                    continue;
                }

                try
                {
                    SetFieldValue(targetItem, action.TargetInternalName, field, replacements);
                    targetItem.Update();
                    context.ExecuteQueryRetry();
                    result.Succeeded = true;
                    result.Message = "Applied successfully.";
                }
                catch (Exception exception)
                {
                    result.Message = exception.Message;
                    warnings.Add($"Field '{action.SourceInternalName}' could not be applied: {exception.Message}");
                }
            }

            return results;
        }

        private static void SetFieldValue(
            ListItem targetItem,
            string targetInternalName,
            EnterpriseWikiFieldValueSnapshot field,
            IEnumerable<EnterpriseWikiTextReplacement> replacements)
        {
            switch (field.Kind)
            {
                case EnterpriseWikiFieldValueKind.String:
                    targetItem[targetInternalName] = RewriteContent(field.Value, replacements);
                    break;
                case EnterpriseWikiFieldValueKind.StringCollection:
                    targetItem[targetInternalName] = field.StringValues.ToArray();
                    break;
                case EnterpriseWikiFieldValueKind.Boolean:
                    targetItem[targetInternalName] = bool.Parse(field.Value);
                    break;
                case EnterpriseWikiFieldValueKind.Number:
                    targetItem[targetInternalName] = double.Parse(field.Value, NumberStyles.Any, CultureInfo.InvariantCulture);
                    break;
                case EnterpriseWikiFieldValueKind.DateTime:
                    targetItem[targetInternalName] = DateTime.Parse(field.Value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                    break;
                case EnterpriseWikiFieldValueKind.Guid:
                    targetItem[targetInternalName] = Guid.Parse(field.Value);
                    break;
                case EnterpriseWikiFieldValueKind.Url:
                    targetItem[targetInternalName] = new FieldUrlValue
                    {
                        Url = RewriteContent(field.UrlValue?.Url, replacements),
                        Description = field.UrlValue?.Description
                    };
                    break;
                default:
                    throw new NotSupportedException($"Field value kind '{field.Kind}' is not importable.");
            }
        }

        private static EnterpriseWikiImportReceipt ReadImportReceipt(
            ClientContext targetContext,
            EnterpriseWikiMigrationPackage package,
            string approvedPlanDigest,
            DateTimeOffset startedAt,
            int materializedDependencyCount,
            IList<EnterpriseWikiFieldImportResult> fieldResults,
            IEnumerable<string> warnings)
        {
            using (var verificationContext = targetContext.Clone(package.Plan.TargetWebUrl))
            {
                var pages = verificationContext.Web.GetPagesLibrary();
                var file = verificationContext.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
                var items = pages.GetItems(new CamlQuery
                {
                    ViewXml = $@"<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='FileRef' />
        <Value Type='Text'>{System.Security.SecurityElement.Escape(package.Plan.TargetPageServerRelativeUrl)}</Value>
      </Eq>
    </Where>
  </Query>
  <ViewFields>
    <FieldRef Name='ID' />
    <FieldRef Name='ContentTypeId' />
    <FieldRef Name='PublishingPageContent' />
    <FieldRef Name='_ModerationStatus' />
  </ViewFields>
  <RowLimit>1</RowLimit>
</View>"
                });
                verificationContext.Load(file,
                    value => value.Exists,
                    value => value.UniqueId,
                    value => value.UIVersionLabel,
                    value => value.Level,
                    value => value.CheckOutType);
                verificationContext.Load(items);
                verificationContext.ExecuteQueryRetry();
                if (!file.Exists)
                {
                    throw new InvalidOperationException("Fresh target readback could not find the imported page.");
                }

                var item = items.SingleOrDefault();
                if (item == null)
                {
                    throw new InvalidOperationException("Fresh target readback could not find the imported page list item.");
                }

                var content = GetFieldString(item, "PublishingPageContent") ?? string.Empty;
                var contentTypeId = GetFieldString(item, "ContentTypeId") ?? string.Empty;
                var webParts = verificationContext.Web.GetWebParts(package.Plan.TargetPageServerRelativeUrl).ToArray();
                var persistedDigest = EnterpriseWikiPackageSerializer.ComputeSha256(content);
                var receiptWarnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList();
                var storageContentEqual = string.Equals(persistedDigest, package.Plan.ExpectedPublishingPageContentSha256, StringComparison.OrdinalIgnoreCase);
                if (!storageContentEqual)
                {
                    receiptWarnings.Add("PublishingPageContent storage bytes differ after SharePoint serialization; DOM + visual acceptance remains the required fidelity gate.");
                }

                var expectedContentPresent = !string.IsNullOrWhiteSpace(package.Snapshot.PublishingPageContent);
                var persistedContentPresent = !string.IsNullOrWhiteSpace(content);
                if (expectedContentPresent && !persistedContentPresent)
                {
                    receiptWarnings.Add("Fresh target readback found empty PublishingPageContent even though the approved source snapshot was non-empty.");
                }

                var actualLevel = file.Level.ToString();
                var actualCheckOutType = file.CheckOutType.ToString();
                var lifecycleMatched = package.Plan.TargetLifecycle == EnterpriseWikiTargetLifecycle.Published
                    ? string.Equals(actualLevel, "Published", StringComparison.OrdinalIgnoreCase)
                    : string.Equals(actualLevel, "Draft", StringComparison.OrdinalIgnoreCase)
                        && string.Equals(actualCheckOutType, "None", StringComparison.OrdinalIgnoreCase);
                if (!lifecycleMatched)
                {
                    receiptWarnings.Add($"Target lifecycle mismatch. Expected {package.Plan.TargetLifecycle}; actual level is {actualLevel} and checkout state is {actualCheckOutType}.");
                }

                var plannedFieldsPassed = fieldResults.All(result => !result.Attempted || result.Succeeded);
                var readbackPassed = IsEnterpriseWikiContentType(contentTypeId)
                    && webParts.Length == package.Snapshot.WebParts.Count
                    && (!expectedContentPresent || persistedContentPresent)
                    && lifecycleMatched
                    && plannedFieldsPassed;
                return new EnterpriseWikiImportReceipt
                {
                    StartedAtUtc = startedAt,
                    CompletedAtUtc = DateTimeOffset.UtcNow,
                    ApprovedPlanDigest = approvedPlanDigest,
                    TargetWebUrl = package.Plan.TargetWebUrl,
                    TargetPageServerRelativeUrl = package.Plan.TargetPageServerRelativeUrl,
                    TargetFileUniqueId = file.UniqueId,
                    TargetListItemId = item.Id,
                    TargetContentTypeId = contentTypeId,
                    TargetVersionLabel = file.UIVersionLabel,
                    ExpectedLifecycle = package.Plan.TargetLifecycle,
                    ActualFileLevel = actualLevel,
                    ActualCheckOutType = actualCheckOutType,
                    ActualModerationStatus = TryGetInt32(item, "_ModerationStatus"),
                    LifecycleMatched = lifecycleMatched,
                    ExpectedPublishingPageContentSha256 = package.Plan.ExpectedPublishingPageContentSha256,
                    PersistedPublishingPageContentSha256 = persistedDigest,
                    StorageContentEqual = storageContentEqual,
                    ImportedWebPartCount = webParts.Length,
                    MaterializedDependencyCount = materializedDependencyCount,
                    FieldResults = fieldResults,
                    FreshReadbackPassed = readbackPassed,
                    Warnings = receiptWarnings,
                    Succeeded = readbackPassed
                };
            }
        }

        private static IList<EnterpriseWikiTextReplacement> BuildReplacements(
            EnterpriseWikiPageIdentity source,
            string targetWebUrl,
            string targetWebServerRelativeUrl)
        {
            var sourceWebUri = new Uri(source.WebUrl);
            var targetWebUri = new Uri(targetWebUrl);
            var candidates = new[]
            {
                new EnterpriseWikiTextReplacement
                {
                    Source = source.WebUrl.TrimEnd('/'),
                    Target = targetWebUrl.TrimEnd('/'),
                    Reason = "Map authored absolute URLs from the source web to the target web."
                },
                new EnterpriseWikiTextReplacement
                {
                    Source = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/'),
                    Target = targetWebServerRelativeUrl.TrimEnd('/'),
                    Reason = "Map authored server-relative URLs from the source web to the target web."
                },
                new EnterpriseWikiTextReplacement
                {
                    Source = sourceWebUri.AbsolutePath.TrimEnd('/'),
                    Target = new Uri(targetWebUrl).AbsolutePath.TrimEnd('/'),
                    Reason = "Map URL-encoded source web paths to the target web."
                },
                new EnterpriseWikiTextReplacement
                {
                    Source = sourceWebUri.GetLeftPart(UriPartial.Authority),
                    Target = targetWebUri.GetLeftPart(UriPartial.Authority),
                    Reason = "Map remaining same-tenant absolute references to the target tenant origin."
                }
            };
            return candidates
                .Where(item => !string.IsNullOrEmpty(item.Source)
                    && !string.Equals(item.Source, item.Target, StringComparison.OrdinalIgnoreCase))
                .GroupBy(item => item.Source, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderByDescending(item => item.Source.Length)
                .ToList();
        }

        private static IList<string> BuildStorageAssertions(
            EnterpriseWikiSnapshot snapshot,
            string targetPagePath,
            IEnumerable<EnterpriseWikiDependencyAction> dependencyActions,
            string expectedContentDigest,
            EnterpriseWikiTargetLifecycle targetLifecycle)
        {
            var result = new List<string>
            {
                $"target-page={targetPagePath}",
                "fresh-read-target-file-identity",
                "fresh-read-target-enterprise-wiki-content-type",
                "fresh-read-target-version-and-lifecycle",
                $"expected-target-lifecycle={targetLifecycle}",
                $"source-publishing-page-content-sha256={snapshot.PublishingPageContentSha256}",
                $"expected-target-publishing-page-content-sha256={expectedContentDigest}",
                $"expected-shared-webparts={snapshot.WebParts.Count}"
            };
            var dependencyById = snapshot.Dependencies.ToDictionary(item => item.Id, StringComparer.Ordinal);
            result.AddRange(dependencyActions
                .Where(item => item.Disposition == EnterpriseWikiDependencyDisposition.MaterializeAtTarget)
                .Select(item => $"dependency={item.TargetServerRelativeUrl}|sha256={dependencyById[item.SnapshotDependencyId].ContentSha256}"));
            return result.OrderBy(item => item, StringComparer.Ordinal).ToList();
        }

        private static List<DependencyCandidate> ExtractDependencyCandidates(string html)
        {
            var result = new List<DependencyCandidate>();
            if (string.IsNullOrWhiteSpace(html))
            {
                return result;
            }

            var document = new HtmlParser().ParseDocument(html);
            foreach (var element in document.All)
            {
                AddAttributeCandidate(result, element, "href", GetKind(element, "href"));
                AddAttributeCandidate(result, element, "src", GetKind(element, "src"));
                AddAttributeCandidate(result, element, "poster", EnterpriseWikiDependencyKind.Media);
                AddAttributeCandidate(result, element, "data", EnterpriseWikiDependencyKind.Object);
                var style = element.GetAttribute("style");
                if (!string.IsNullOrWhiteSpace(style))
                {
                    foreach (Match match in CssUrlPattern.Matches(style))
                    {
                        result.Add(new DependencyCandidate
                        {
                            Consumer = $"{element.LocalName}[style]",
                            Kind = EnterpriseWikiDependencyKind.Image,
                            Value = match.Groups["url"].Value.Trim(),
                            IsRenderableResource = true
                        });
                    }
                }
            }

            return result;
        }

        private static IEnumerable<DependencyCandidate> ExtractUrlCandidates(string text, string consumer)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return Array.Empty<DependencyCandidate>();
            }

            return Regex.Matches(text, @"https?://[^\s'""<>]+|(?<quote>['""])(?<path>/[^'""<>\s]+)\k<quote>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Cast<Match>()
                .Select(match => new DependencyCandidate
                {
                    Consumer = consumer,
                    Kind = EnterpriseWikiDependencyKind.Unknown,
                    Value = (match.Groups["path"].Success ? match.Groups["path"].Value : match.Value).TrimEnd('.', ',', ';', ')'),
                    IsRenderableResource = false
                })
                .ToArray();
        }

        private static void AddAttributeCandidate(
            ICollection<DependencyCandidate> result,
            IElement element,
            string attributeName,
            EnterpriseWikiDependencyKind kind)
        {
            var value = element.GetAttribute(attributeName);
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            result.Add(new DependencyCandidate
            {
                Consumer = $"{element.LocalName}[{attributeName}]",
                Kind = kind,
                Value = value.Trim(),
                IsRenderableResource = kind != EnterpriseWikiDependencyKind.Anchor
                    && kind != EnterpriseWikiDependencyKind.Unknown
            });
        }

        private static EnterpriseWikiDependencyKind GetKind(IElement element, string attributeName)
        {
            switch (element.LocalName.ToLowerInvariant())
            {
                case "a":
                case "area":
                    return EnterpriseWikiDependencyKind.Anchor;
                case "img":
                    return EnterpriseWikiDependencyKind.Image;
                case "script":
                    return EnterpriseWikiDependencyKind.Script;
                case "link":
                    return EnterpriseWikiDependencyKind.StyleSheet;
                case "iframe":
                    return EnterpriseWikiDependencyKind.IFrame;
                case "object":
                    return EnterpriseWikiDependencyKind.Object;
                case "audio":
                case "source":
                case "video":
                    return EnterpriseWikiDependencyKind.Media;
                default:
                    return attributeName == "href"
                        ? EnterpriseWikiDependencyKind.Anchor
                        : EnterpriseWikiDependencyKind.Unknown;
            }
        }

        private static bool TryResolveDependencyUri(Uri sourcePageUri, string value, out Uri result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(value)
                || value.StartsWith("#", StringComparison.Ordinal)
                || value.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("tel:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return Uri.TryCreate(sourcePageUri, value, out result)
                && (result.Scheme == Uri.UriSchemeHttps || result.Scheme == Uri.UriSchemeHttp);
        }

        private static bool IsSharePointRuntimePath(string serverRelativeUrl)
        {
            return serverRelativeUrl.StartsWith("/_layouts/", StringComparison.OrdinalIgnoreCase)
                || serverRelativeUrl.StartsWith("/_vti_bin/", StringComparison.OrdinalIgnoreCase)
                || serverRelativeUrl.StartsWith("/_api/", StringComparison.OrdinalIgnoreCase)
                || serverRelativeUrl.IndexOf("/_catalogs/masterpage/", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string NormalizePagePath(string webServerRelativeUrl, string value, string defaultLibrary)
        {
            var candidate = value.Trim();
            if (Uri.TryCreate(candidate, UriKind.Absolute, out var absolute))
            {
                candidate = Uri.UnescapeDataString(absolute.AbsolutePath);
            }
            else
            {
                candidate = Uri.UnescapeDataString(candidate.Split(new[] { '?', '#' }, 2)[0]).Replace('\\', '/');
            }

            if (!candidate.StartsWith("/", StringComparison.Ordinal))
            {
                if (!candidate.Contains("/"))
                {
                    candidate = $"{defaultLibrary}/{candidate}";
                }
                candidate = $"{webServerRelativeUrl.TrimEnd('/')}/{candidate.TrimStart('/')}";
            }

            if (!candidate.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
            {
                candidate += ".aspx";
            }

            return candidate;
        }

        private static EnterpriseWikiSourceFence ReadFence(ClientContext context, string pagePath)
        {
            var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            context.Load(file, value => value.Exists, value => value.UniqueId, value => value.UIVersionLabel, value => value.Length, value => value.TimeLastModified);
            context.ExecuteQueryRetry();
            if (!file.Exists)
            {
                throw new FileNotFoundException("The source page disappeared during capture.", pagePath);
            }
            return ToFence(file);
        }

        private static EnterpriseWikiSourceFence ToFence(CsomFile file)
        {
            return new EnterpriseWikiSourceFence
            {
                FileUniqueId = file.UniqueId,
                VersionLabel = file.UIVersionLabel,
                Length = file.Length,
                ModifiedUtc = file.TimeLastModified.ToUniversalTime()
            };
        }

        private static bool FenceEquals(EnterpriseWikiSourceFence left, EnterpriseWikiSourceFence right)
        {
            return left.FileUniqueId == right.FileUniqueId
                && left.Length == right.Length
                && left.ModifiedUtc == right.ModifiedUtc
                && string.Equals(left.VersionLabel, right.VersionLabel, StringComparison.Ordinal);
        }

        private static bool FileExists(ClientContext context, string serverRelativeUrl)
        {
            var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            context.Load(file, value => value.Exists);
            try
            {
                context.ExecuteQueryRetry();
                return file.Exists;
            }
            catch (ServerException exception) when (exception.ServerErrorTypeName == "System.IO.FileNotFoundException")
            {
                return false;
            }
        }

        private static string ComputeBytesSha256(byte[] bytes)
        {
            using (var algorithm = SHA256.Create())
            {
                var digest = algorithm.ComputeHash(bytes);
                return string.Concat(digest.Select(item => item.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }

        private static string GetFieldString(ListItem item, string internalName)
        {
            return item.FieldValues.TryGetValue(internalName, out var value)
                ? Convert.ToString(value, CultureInfo.InvariantCulture)
                : null;
        }

        private static int? TryGetInt32(ListItem item, string internalName)
        {
            if (!item.FieldValues.TryGetValue(internalName, out var value) || value == null)
            {
                return null;
            }

            return int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out var result)
                ? result
                : (int?)null;
        }

        private static bool IsPathWithin(string candidate, string root)
        {
            var normalizedCandidate = (candidate ?? string.Empty).TrimEnd('/');
            var normalizedRoot = (root ?? string.Empty).TrimEnd('/');
            return string.Equals(normalizedCandidate, normalizedRoot, StringComparison.OrdinalIgnoreCase)
                || normalizedCandidate.StartsWith(normalizedRoot + "/", StringComparison.OrdinalIgnoreCase);
        }

        private static bool UriEquals(string left, string right)
        {
            return string.Equals(left?.TrimEnd('/'), right?.TrimEnd('/'), StringComparison.OrdinalIgnoreCase);
        }

        private static string GetDirectoryName(string serverRelativeUrl)
        {
            var separator = serverRelativeUrl.LastIndexOf('/');
            return separator <= 0 ? "/" : serverRelativeUrl.Substring(0, separator);
        }

        private static string GetFileName(string serverRelativeUrl)
        {
            var separator = serverRelativeUrl.LastIndexOf('/');
            return separator < 0 ? serverRelativeUrl : serverRelativeUrl.Substring(separator + 1);
        }

        private static string EnsureTrailingSlash(string value)
        {
            return value.EndsWith("/", StringComparison.Ordinal) ? value : value + "/";
        }

        private static string EncodePath(string decodedPath)
        {
            return string.Join("/", decodedPath.Split('/').Select(Uri.EscapeDataString));
        }

        private sealed class SourceCapture
        {
            public EnterpriseWikiPageIdentity Identity { get; set; }

            public string PublishingPageContent { get; set; }

            public List<EnterpriseWikiFieldValueSnapshot> Fields { get; set; }

            public List<EnterpriseWikiWebPartSnapshot> WebParts { get; set; }

            public EnterpriseWikiSecuritySnapshot Security { get; set; }

            public EnterpriseWikiLifecycleSnapshot Lifecycle { get; set; }

            public EnterpriseWikiSourceFence SourceFence { get; set; }
        }

        private sealed class DependencyCandidate
        {
            public string Value { get; set; }

            public string Consumer { get; set; }

            public EnterpriseWikiDependencyKind Kind { get; set; }

            public bool IsRenderableResource { get; set; }
        }
    }
}
