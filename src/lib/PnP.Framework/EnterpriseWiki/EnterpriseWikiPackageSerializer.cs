using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PnP.Framework.EnterpriseWiki
{
    public static class EnterpriseWikiPackageSerializer
    {
        public const string ExportSchemaVersion = "pnp-enterprise-wiki-export/v2";
        public const string MigrationSchemaVersion = "pnp-enterprise-wiki-migration-package/v2";
        public const string DefaultExportFileName = "enterprise-wiki-export.json";
        public const string DefaultPackageFileName = "enterprise-wiki-package.json";
        public const string DefaultReportFileName = "enterprise-wiki-report.md";
        public const string DefaultReceiptFileName = "enterprise-wiki-import-receipt.json";

        private static readonly JsonSerializerOptions CanonicalOptions = CreateOptions(false);
        private static readonly JsonSerializerOptions IndentedOptions = CreateOptions(true);

        public static string ComputeSnapshotDigest(EnterpriseWikiSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            return ComputeSha256(JsonSerializer.Serialize(snapshot, CanonicalOptions));
        }

        public static string ComputePlanDigest(EnterpriseWikiMigrationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return ComputeSha256(JsonSerializer.Serialize(plan, CanonicalOptions));
        }

        public static string ComputeSha256(string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            using (var algorithm = SHA256.Create())
            {
                var bytes = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value));
                var builder = new StringBuilder(bytes.Length * 2);
                foreach (var item in bytes)
                {
                    builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        public static string SaveExport(string path, EnterpriseWikiExportPackage package, bool overwrite = false)
        {
            ValidateExport(package);
            var exportPath = ResolvePath(path, DefaultExportFileName);
            SaveJson(exportPath, package, overwrite);
            return exportPath;
        }

        public static EnterpriseWikiExportPackage LoadExport(string path)
        {
            var exportPath = ResolveExistingPath(path, DefaultExportFileName, "Enterprise Wiki export");
            var package = JsonSerializer.Deserialize<EnterpriseWikiExportPackage>(File.ReadAllText(exportPath), IndentedOptions);
            ValidateExport(package);
            return package;
        }

        public static string SaveMigration(string path, EnterpriseWikiMigrationPackage package, bool overwrite = false)
        {
            ValidateMigration(package);
            var packagePath = ResolvePath(path, DefaultPackageFileName);
            var reportPath = Path.Combine(Path.GetDirectoryName(packagePath) ?? string.Empty, DefaultReportFileName);
            if (File.Exists(reportPath) && !overwrite)
            {
                throw new IOException($"The report file already exists: {reportPath}");
            }

            SaveJson(packagePath, package, overwrite);
            File.WriteAllText(reportPath, BuildReport(package), new UTF8Encoding(false));
            return packagePath;
        }

        public static EnterpriseWikiMigrationPackage LoadMigration(string path)
        {
            var packagePath = ResolveExistingPath(path, DefaultPackageFileName, "Enterprise Wiki migration package");
            var package = JsonSerializer.Deserialize<EnterpriseWikiMigrationPackage>(File.ReadAllText(packagePath), IndentedOptions);
            ValidateMigration(package);
            return package;
        }

        public static string SaveReceipt(string path, EnterpriseWikiImportReceipt receipt, bool overwrite = false)
        {
            if (receipt == null)
            {
                throw new ArgumentNullException(nameof(receipt));
            }

            var receiptPath = ResolvePath(path, DefaultReceiptFileName);
            SaveJson(receiptPath, receipt, overwrite);
            return receiptPath;
        }

        public static void ValidateExport(EnterpriseWikiExportPackage package)
        {
            if (package == null)
            {
                throw new InvalidDataException("The Enterprise Wiki export is empty.");
            }

            if (!string.Equals(package.SchemaVersion, ExportSchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported Enterprise Wiki export schema '{package.SchemaVersion}'.");
            }

            if (package.Snapshot == null)
            {
                throw new InvalidDataException("The Enterprise Wiki export must contain a source snapshot.");
            }

            if (package.Snapshot.Source == null || package.Snapshot.CapturePolicy == null)
            {
                throw new InvalidDataException("The Enterprise Wiki source snapshot is missing identity or capture policy data.");
            }

            if (package.Snapshot.Fields == null
                || package.Snapshot.WebParts == null
                || package.Snapshot.Dependencies == null
                || package.Snapshot.Blockers == null
                || package.Snapshot.Warnings == null)
            {
                throw new InvalidDataException("The Enterprise Wiki source snapshot contains a null inventory collection.");
            }

            var contentDigest = ComputeSha256(package.Snapshot.PublishingPageContent ?? string.Empty);
            if (!string.Equals(contentDigest, package.Snapshot.PublishingPageContentSha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The PublishingPageContent digest does not match the source HTML.");
            }

            var duplicateField = package.Snapshot.Fields
                .GroupBy(item => item.InternalName, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateField != null)
            {
                throw new InvalidDataException($"The source field inventory contains a missing or duplicate internal name '{duplicateField.Key}'.");
            }

            foreach (var webPart in package.Snapshot.WebParts)
            {
                if (!string.Equals(ComputeSha256(webPart.ExportXml ?? string.Empty), webPart.ExportSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Web Part export digest mismatch: {webPart.Id}");
                }
            }

            var duplicateDependency = package.Snapshot.Dependencies
                .GroupBy(item => item.Id, StringComparer.Ordinal)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateDependency != null)
            {
                throw new InvalidDataException($"The dependency inventory contains a missing or duplicate ID '{duplicateDependency.Key}'.");
            }

            foreach (var dependency in package.Snapshot.Dependencies.Where(item => !string.IsNullOrWhiteSpace(item.ContentBase64)))
            {
                byte[] payload;
                try
                {
                    payload = Convert.FromBase64String(dependency.ContentBase64);
                }
                catch (FormatException exception)
                {
                    throw new InvalidDataException($"Dependency payload is not valid Base64: {dependency.Id}", exception);
                }

                if (payload.LongLength != dependency.ContentLength
                    || !string.Equals(ComputeBytesSha256(payload), dependency.ContentSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Dependency payload length or digest mismatch: {dependency.Id}");
                }
            }

            var snapshotDigest = ComputeSnapshotDigest(package.Snapshot);
            if (!string.Equals(snapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The Enterprise Wiki snapshot digest does not match the export payload.");
            }
        }

        public static void ValidateMigration(EnterpriseWikiMigrationPackage package)
        {
            if (package == null)
            {
                throw new InvalidDataException("The Enterprise Wiki migration package is empty.");
            }

            if (!string.Equals(package.SchemaVersion, MigrationSchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported Enterprise Wiki migration package schema '{package.SchemaVersion}'.");
            }

            if (!string.Equals(package.ExportSchemaVersion, ExportSchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported embedded Enterprise Wiki export schema '{package.ExportSchemaVersion}'.");
            }

            if (package.Snapshot == null || package.Plan == null)
            {
                throw new InvalidDataException("The Enterprise Wiki migration package must contain both a snapshot and a plan.");
            }

            ValidateExport(new EnterpriseWikiExportPackage
            {
                SchemaVersion = package.ExportSchemaVersion,
                ExportedAtUtc = package.ExportedAtUtc,
                Snapshot = package.Snapshot,
                SnapshotDigest = package.SnapshotDigest
            });

            var snapshotDigest = ComputeSnapshotDigest(package.Snapshot);
            if (!string.Equals(snapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The Enterprise Wiki snapshot digest does not match the package payload.");
            }

            if (!string.Equals(package.Plan.SourceSnapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The migration plan does not reference the sealed snapshot in this package.");
            }

            var planDigest = ComputePlanDigest(package.Plan);
            if (!string.Equals(planDigest, package.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The Enterprise Wiki plan digest does not match the package payload.");
            }

            if (package.Plan.FieldActions == null || package.Plan.DependencyActions == null)
            {
                throw new InvalidDataException("The migration plan contains a null action collection.");
            }

            var sourceFieldNames = new HashSet<string>(package.Snapshot.Fields.Select(item => item.InternalName), StringComparer.OrdinalIgnoreCase);
            var plannedFieldNames = new HashSet<string>(package.Plan.FieldActions.Select(item => item.SourceInternalName), StringComparer.OrdinalIgnoreCase);
            if (package.Plan.FieldActions.Count != sourceFieldNames.Count
                || plannedFieldNames.Count != sourceFieldNames.Count
                || !sourceFieldNames.SetEquals(plannedFieldNames))
            {
                throw new InvalidDataException("The plan must contain exactly one field action for every captured source field.");
            }

            var dependencyIds = new HashSet<string>(package.Snapshot.Dependencies.Select(item => item.Id), StringComparer.Ordinal);
            var plannedDependencyIds = new HashSet<string>(package.Plan.DependencyActions.Select(item => item.SnapshotDependencyId), StringComparer.Ordinal);
            if (package.Plan.DependencyActions.Count != dependencyIds.Count
                || plannedDependencyIds.Count != dependencyIds.Count
                || !dependencyIds.SetEquals(plannedDependencyIds))
            {
                throw new InvalidDataException("The plan must contain exactly one dependency action for every captured dependency.");
            }

            var expectedContent = EnterpriseWikiMigrationService.RewriteContent(
                package.Snapshot.PublishingPageContent,
                package.Plan.Replacements);
            if (!string.Equals(
                    ComputeSha256(expectedContent),
                    package.Plan.ExpectedPublishingPageContentSha256,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The expected target PublishingPageContent digest does not match the approved replacements.");
            }

            var derivedLifecycle = EnterpriseWikiMigrationService.DeriveTargetLifecycle(package.Snapshot.Lifecycle);
            if (package.Plan.TargetLifecycle != derivedLifecycle)
            {
                throw new InvalidDataException($"The planned lifecycle '{package.Plan.TargetLifecycle}' does not match the source-derived lifecycle '{derivedLifecycle}'.");
            }

            var expectedState = package.Plan.IsExecutable
                ? EnterpriseWikiPackageState.ApprovalReady
                : EnterpriseWikiPackageState.Blocked;
            if (package.State != expectedState)
            {
                throw new InvalidDataException($"Package state '{package.State}' does not match plan executability '{expectedState}'.");
            }
        }

        private static string ComputeBytesSha256(byte[] value)
        {
            using (var algorithm = SHA256.Create())
            {
                var bytes = algorithm.ComputeHash(value);
                var builder = new StringBuilder(bytes.Length * 2);
                foreach (var item in bytes)
                {
                    builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        public static string BuildReport(EnterpriseWikiMigrationPackage package)
        {
            ValidateMigration(package);
            var snapshot = package.Snapshot;
            var plan = package.Plan;
            var report = package.Report ?? new EnterpriseWikiCustomerReport();
            var builder = new StringBuilder();
            builder.AppendLine("# Enterprise Wiki migration report");
            builder.AppendLine();
            builder.AppendLine(report.Summary ?? "This report describes the sealed source snapshot and target-specific migration plan.");
            builder.AppendLine();
            builder.AppendLine("The JSON package is authoritative. This Markdown is a complete review view: large HTML, XML, schema, JSON, and Base64 values are represented by length, SHA-256, and a preview while their full values remain in the package.");
            builder.AppendLine();

            AppendTable(builder, "Package envelope", new[] { "JSON path", "Value", "How to read it" }, new[]
            {
                Row("schemaVersion", package.SchemaVersion, "Version of the planned/importable package contract."),
                Row("plannedAtUtc", Format(package.PlannedAtUtc), "UTC time at which the target-specific plan was sealed."),
                Row("exportSchemaVersion", package.ExportSchemaVersion, "Version of the embedded source-only export contract."),
                Row("exportedAtUtc", Format(package.ExportedAtUtc), "UTC time at which the source snapshot was captured."),
                Row("state", package.State, "ApprovalReady can be imported after digest approval; Blocked cannot."),
                Row("snapshotDigest", package.SnapshotDigest, "SHA-256 over the complete source snapshot."),
                Row("planDigest", package.PlanDigest, "SHA-256 over all target decisions; this is the approval token."),
                Row("plan.isExecutable", plan.IsExecutable, "True only when the blocker list is empty.")
            });

            AppendTable(builder, "Source page identity", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("webUrl", snapshot.Source.WebUrl, "Absolute source web URL."),
                Row("webServerRelativeUrl", snapshot.Source.WebServerRelativeUrl, "Server-relative source web boundary used for URL mapping."),
                Row("pageServerRelativeUrl", snapshot.Source.PageServerRelativeUrl, "Exact source page path."),
                Row("listItemId", snapshot.Source.ListItemId, "Source Pages-library item ID; evidence only across sites."),
                Row("fileUniqueId", snapshot.Source.FileUniqueId, "Source file GUID used as identity evidence."),
                Row("contentTypeId", snapshot.Source.ContentTypeId, "Must derive from Enterprise Wiki Page and not Project Page."),
                Row("contentTypeName", snapshot.Source.ContentTypeName, "Human-readable source content type."),
                Row("versionLabel", snapshot.Source.VersionLabel, "Version captured by the export."),
                Row("length", snapshot.Source.Length, "Source ASPX file length in bytes."),
                Row("modifiedUtc", Format(snapshot.Source.ModifiedUtc), "Source file modified time at capture."),
                Row("title", snapshot.Source.Title, "Title assigned during target page creation."),
                Row("pageLayoutUrl", snapshot.Source.PageLayoutUrl, "Source publishing layout; exact profile requires EnterpriseWiki.aspx."),
                Row("pageLayoutDescription", snapshot.Source.PageLayoutDescription, "Source layout description, if present.")
            });

            AppendTable(builder, "Capture policy and source fence", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("capturePolicy.sourcePageServerRelativeUrl", snapshot.CapturePolicy?.SourcePageServerRelativeUrl, "Normalized source page requested at export."),
                Row("capturePolicy.includeWebParts", snapshot.CapturePolicy?.IncludeWebParts, "Whether shared Web Parts were inventoried and exported."),
                Row("capturePolicy.maximumDependencyBytes", snapshot.CapturePolicy?.MaximumDependencyBytes, "Maximum bytes captured for each restorable dependency."),
                Row("sourceFence.fileUniqueId", snapshot.SourceFence?.FileUniqueId, "Identity checked before and after export."),
                Row("sourceFence.versionLabel", snapshot.SourceFence?.VersionLabel, "Version checked before and after export."),
                Row("sourceFence.length", snapshot.SourceFence?.Length, "File length checked before and after export."),
                Row("sourceFence.modifiedUtc", Format(snapshot.SourceFence?.ModifiedUtc), "Modified time checked before and after export.")
            });

            AppendTable(builder, "Lifecycle decision", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("snapshot.lifecycle.checkOutType", snapshot.Lifecycle?.CheckOutType, "Source checkout state captured as evidence."),
                Row("snapshot.lifecycle.level", snapshot.Lifecycle?.Level, "Only Published maps to Published; every other value maps to Draft."),
                Row("snapshot.lifecycle.moderationStatus", snapshot.Lifecycle?.ModerationStatus, "Source moderation status value, if available."),
                Row("snapshot.lifecycle.createdUtc", Format(snapshot.Lifecycle?.CreatedUtc), "Source creation time; not replayed as a system field."),
                Row("snapshot.lifecycle.modifiedUtc", Format(snapshot.Lifecycle?.ModifiedUtc), "Source modified time; not replayed as a system field."),
                Row("plan.targetLifecycle", plan.TargetLifecycle, "Derived lifecycle to enforce at the target."),
                Row("plan.lifecycleReason", plan.LifecycleReason, "Human-readable derivation rule; there is no publish Boolean input.")
            });

            AppendTable(builder, "Publishing content", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("snapshot.publishingPageContent", SummarizePayload(snapshot.PublishingPageContent), "Full HTML remains in JSON; this row gives length, digest, and preview."),
                Row("snapshot.publishingPageContentSha256", snapshot.PublishingPageContentSha256, "Digest of source HTML before URL replacements."),
                Row("plan.expectedPublishingPageContentSha256", plan.ExpectedPublishingPageContentSha256, "Digest expected after approved replacements and before SharePoint serialization.")
            });

            var actionByField = plan.FieldActions.ToDictionary(item => item.SourceInternalName, StringComparer.OrdinalIgnoreCase);
            AppendTable(builder, $"Complete list-item field inventory ({snapshot.Fields.Count})",
                new[] { "#", "Internal name", "Title / ID", "Source type", "Capture", "Value", "Flags", "Plan action", "Target", "Reason" },
                snapshot.Fields.Select((field, index) =>
                {
                    actionByField.TryGetValue(field.InternalName, out var action);
                    return Row(
                        index + 1,
                        field.InternalName,
                        $"{Format(field.Title)} / {field.Id:D}",
                        field.TypeAsString,
                        $"{field.CaptureStatus}; hasValue={field.HasValue}; rawType={Format(field.RawType)}",
                        SummarizeFieldValue(field),
                        $"readOnly={field.ReadOnly}; hidden={field.Hidden}; required={field.Required}",
                        action?.Disposition,
                        action == null ? null : $"{Format(action.TargetInternalName)} ({Format(action.TargetTypeAsString)})",
                        action?.Reason);
                }));

            builder.AppendLine("## Per-field recovery evidence");
            builder.AppendLine();
            foreach (var field in snapshot.Fields)
            {
                builder.AppendLine($"### {EscapeHeading(field.InternalName)}");
                builder.AppendLine();
                AppendTable(builder, null, new[] { "Property", "Value" }, new[]
                {
                    Row("id", field.Id),
                    Row("title", field.Title),
                    Row("schemaXml", SummarizePayload(field.SchemaXml)),
                    Row("kind", field.Kind),
                    Row("value", field.Value),
                    Row("stringValues", Join(field.StringValues)),
                    Row("urlValue", field.UrlValue == null ? null : $"url={Format(field.UrlValue.Url)}; description={Format(field.UrlValue.Description)}"),
                    Row("lookupValues", Join(field.LookupValues.Select(value => $"id={value.LookupId}; value={Format(value.LookupValue)}"))),
                    Row("taxonomyValues", Join(field.TaxonomyValues.Select(value => $"label={Format(value.Label)}; termGuid={Format(value.TermGuid)}; wssId={value.WssId}"))),
                    Row("binaryBase64", SummarizePayload(field.BinaryBase64)),
                    Row("rawType", field.RawType),
                    Row("rawValue", SummarizePayload(field.RawValue)),
                    Row("rawValueJson", SummarizePayload(field.RawValueJson)),
                    Row("captureStatus", field.CaptureStatus),
                    Row("diagnostics", Join(field.Diagnostics))
                });
            }

            AppendTable(builder, $"Shared Web Parts ({snapshot.WebParts.Count})",
                new[] { "ID", "Title", "Zone", "Index", "Hidden", "Export evidence" },
                snapshot.WebParts.Select(item => Row(item.Id, item.Title, item.ZoneId, item.ZoneIndex, item.Hidden, $"sha256={item.ExportSha256}; {SummarizePayload(item.ExportXml)}")));

            var actionByDependency = plan.DependencyActions.ToDictionary(item => item.SnapshotDependencyId, StringComparer.Ordinal);
            AppendTable(builder, $"Dependencies ({snapshot.Dependencies.Count})",
                new[] { "ID", "Consumer / kind", "Source", "Capture", "Payload", "Plan action", "Target", "Diagnostics" },
                snapshot.Dependencies.Select(item =>
                {
                    actionByDependency.TryGetValue(item.Id, out var action);
                    return Row(
                        item.Id,
                        $"{item.Consumer} / {item.Kind}; renderable={item.IsRenderableResource}",
                        $"original={Format(item.OriginalValue)}; absolute={Format(item.SourceAbsoluteUrl)}; path={Format(item.SourceServerRelativeUrl)}",
                        item.CaptureStatus,
                        string.IsNullOrWhiteSpace(item.ContentBase64) ? null : $"bytes={item.ContentLength}; sha256={item.ContentSha256}",
                        action?.Disposition,
                        action == null ? null : $"absolute={Format(action.TargetAbsoluteUrl)}; path={Format(action.TargetServerRelativeUrl)}",
                        Join(item.Diagnostics.Concat(action?.Diagnostics ?? Array.Empty<string>())));
                }));

            AppendTable(builder, "Security snapshot", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("hasUniqueRoleAssignments", snapshot.Security?.HasUniqueRoleAssignments, "Unique permissions are captured; the current importer does not replay them."),
                Row("roleAssignments", snapshot.Security == null ? null : Join(snapshot.Security.RoleAssignments.Select(item => $"{item.PrincipalTitle} ({item.PrincipalLoginName}): {Join(item.RoleDefinitionNames)}")), "Complete source role-assignment evidence when inheritance is broken.")
            });

            AppendTable(builder, "Plan and target probe", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("plan.operation", plan.Operation, "CreatePage is executable now; ApplyDeferredFields reserves future snapshot-based recovery."),
                Row("plan.sourceSnapshotDigest", plan.SourceSnapshotDigest, "Binds this plan to the exact source snapshot."),
                Row("plan.sourceWebUrl", plan.SourceWebUrl, "Source web used for mapping."),
                Row("plan.sourcePageServerRelativeUrl", plan.SourcePageServerRelativeUrl, "Source page represented by the snapshot."),
                Row("plan.targetWebUrl", plan.TargetWebUrl, "Target connection must point here during import."),
                Row("plan.targetWebServerRelativeUrl", plan.TargetWebServerRelativeUrl, "Target web boundary for paths."),
                Row("plan.targetPageServerRelativeUrl", plan.TargetPageServerRelativeUrl, "Exact target page to create."),
                Row("plan.pageLayoutName", plan.PageLayoutName, "Publishing layout name passed to page creation."),
                Row("plan.createOnly", plan.CreateOnly, "Existing pages and dependency files are blockers; Force does not change this."),
                Row("planningPolicy.requireInheritedPermissions", plan.PlanningPolicy?.RequireInheritedPermissions, "Blocks source pages with unique permissions when true."),
                Row("planningPolicy.blockOnManagedMetadata", plan.PlanningPolicy?.BlockOnManagedMetadata, "Blocks non-empty taxonomy values without a reviewed term mapping when true."),
                Row("planningPolicy.allowExternalResourceReferences", plan.PlanningPolicy?.AllowExternalResourceReferences, "Allows external authored resource references to remain external when true."),
                Row("targetProbe.webUrl", plan.TargetProbe?.WebUrl, "Web actually probed while planning."),
                Row("targetProbe.webServerRelativeUrl", plan.TargetProbe?.WebServerRelativeUrl, "Target web path boundary."),
                Row("targetProbe.webTemplate", plan.TargetProbe?.WebTemplate, "Target web template identifier."),
                Row("targetProbe.webConfiguration", plan.TargetProbe?.WebConfiguration, "Target web configuration number."),
                Row("targetProbe.pagesLibraryServerRelativeUrl", plan.TargetProbe?.PagesLibraryServerRelativeUrl, "Publishing Pages library root."),
                Row("targetProbe.pagesLibraryBaseTemplate", plan.TargetProbe?.PagesLibraryBaseTemplate, "Expected publishing template value is 850."),
                Row("targetProbe.enableVersioning", plan.TargetProbe?.EnableVersioning, "Required for deterministic lifecycle handling."),
                Row("targetProbe.enableMinorVersions", plan.TargetProbe?.EnableMinorVersions, "Required when the derived target lifecycle is Draft."),
                Row("targetProbe.enableModeration", plan.TargetProbe?.EnableModeration, "Controls whether a published page also needs approval."),
                Row("targetProbe.forceCheckout", plan.TargetProbe?.ForceCheckout, "Controls whether the importer must check out before updates."),
                Row("targetProbe.draftVersionVisibility", plan.TargetProbe?.DraftVersionVisibility, "Who can see draft/minor versions."),
                Row("targetProbe.enterpriseWikiContentTypeId", plan.TargetProbe?.EnterpriseWikiContentTypeId, "Enterprise Wiki Page content type found in the target library."),
                Row("targetProbe.enterpriseWikiLayoutUrl", plan.TargetProbe?.EnterpriseWikiLayoutUrl, "EnterpriseWiki.aspx found in the target site collection."),
                Row("targetProbe.enterpriseWikiLayoutExists", plan.TargetProbe?.EnterpriseWikiLayoutExists, "False is a blocker."),
                Row("targetProbe.targetPageExists", plan.TargetProbe?.TargetPageExists, "True is a blocker for CreatePage."),
                Row("targetProbe.existingDependencyPaths", Join(plan.TargetProbe?.ExistingDependencyPaths), "Create-only dependency collisions.")
            });

            AppendTable(builder, "Approved text replacements", new[] { "Source", "Target", "Reason" },
                plan.Replacements.Select(item => Row(item.Source, item.Target, item.Reason)));
            AppendList(builder, "Storage assertions", plan.StorageAssertions);
            AppendList(builder, "Browser acceptance assertions", plan.BrowserAssertions);
            AppendList(builder, "Blockers", plan.Blockers);
            AppendList(builder, "Warnings", plan.Warnings);
            AppendList(builder, "Captured ingredients", report.CapturedIngredients);

            builder.AppendLine("## Field-action legend");
            builder.AppendLine();
            builder.AppendLine("- `Apply`: recognized and supported; Import writes it.");
            builder.AppendLine("- `AlreadyHandled`: page creation, content, layout, or content type logic handles it outside the generic field loop.");
            builder.AppendLine("- `EvidenceOnly`: the complete source value remains available for a future mapper, but Import does not write it today.");
            builder.AppendLine("- `RequiresMapping`: user, lookup, or taxonomy identity cannot safely cross sites without an explicit mapping.");
            builder.AppendLine("- `Skip*`, `Target*`, or `CaptureUnavailable`: retained and explained, but not written.");
            builder.AppendLine("- `Block`: the plan cannot be imported.");
            builder.AppendLine();
            return builder.ToString();
        }

        private static string SummarizeFieldValue(EnterpriseWikiFieldValueSnapshot field)
        {
            switch (field.Kind)
            {
                case EnterpriseWikiFieldValueKind.Url:
                    return field.UrlValue == null ? null : $"url={Format(field.UrlValue.Url)}; description={Format(field.UrlValue.Description)}";
                case EnterpriseWikiFieldValueKind.StringCollection:
                    return Join(field.StringValues);
                case EnterpriseWikiFieldValueKind.User:
                case EnterpriseWikiFieldValueKind.UserCollection:
                case EnterpriseWikiFieldValueKind.Lookup:
                case EnterpriseWikiFieldValueKind.LookupCollection:
                    return Join(field.LookupValues.Select(value => $"{value.LookupId}:{value.LookupValue}"));
                case EnterpriseWikiFieldValueKind.Taxonomy:
                case EnterpriseWikiFieldValueKind.TaxonomyCollection:
                    return Join(field.TaxonomyValues.Select(value => $"{value.Label}|{value.TermGuid}|{value.WssId}"));
                case EnterpriseWikiFieldValueKind.ByteArray:
                    return SummarizePayload(field.BinaryBase64);
                case EnterpriseWikiFieldValueKind.Unsupported:
                    return SummarizePayload(field.RawValueJson ?? field.RawValue);
                default:
                    return SummarizePayload(field.Value);
            }
        }

        private static string SummarizePayload(string value)
        {
            if (value == null)
            {
                return null;
            }

            var preview = value.Replace("\r", " ").Replace("\n", " ");
            if (preview.Length > 160)
            {
                preview = preview.Substring(0, 160) + "…";
            }

            return $"length={value.Length}; sha256={ComputeSha256(value)}; preview={preview}";
        }

        private static string Join(IEnumerable<string> values)
        {
            var items = (values ?? Array.Empty<string>()).Where(value => value != null).ToArray();
            return items.Length == 0 ? null : string.Join("; ", items);
        }

        private static string Format(object value)
        {
            if (value == null)
            {
                return "(null)";
            }

            if (value is DateTime dateTime)
            {
                return dateTime.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }

            if (value is DateTimeOffset dateTimeOffset)
            {
                return dateTimeOffset.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }

            var result = Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.IsNullOrEmpty(result) ? "(empty)" : result;
        }

        private static string EscapeHeading(string value)
        {
            return (value ?? "(unnamed)").Replace("#", "\\#").Replace("`", "\\`");
        }

        private static string[] Row(params object[] values)
        {
            return values.Select(Format).ToArray();
        }

        private static void AppendTable(StringBuilder builder, string heading, string[] headers, IEnumerable<string[]> rows)
        {
            if (!string.IsNullOrWhiteSpace(heading))
            {
                builder.AppendLine($"## {heading}");
                builder.AppendLine();
            }

            builder.AppendLine("| " + string.Join(" | ", headers.Select(EscapeTableCell)) + " |");
            builder.AppendLine("| " + string.Join(" | ", headers.Select(_ => "---")) + " |");
            var any = false;
            foreach (var row in rows ?? Array.Empty<string[]>())
            {
                any = true;
                builder.AppendLine("| " + string.Join(" | ", row.Select(EscapeTableCell)) + " |");
            }
            if (!any)
            {
                builder.AppendLine("| " + string.Join(" | ", headers.Select((_, index) => index == 0 ? "None" : string.Empty)) + " |");
            }
            builder.AppendLine();
        }

        private static string EscapeTableCell(string value)
        {
            return Format(value)
                .Replace("|", "\\|")
                .Replace("\r", " ")
                .Replace("\n", " ");
        }

        private static void AppendList(StringBuilder builder, string heading, IEnumerable<string> values)
        {
            var items = (values ?? Array.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
            builder.AppendLine($"## {heading}");
            builder.AppendLine();
            if (items.Length == 0)
            {
                builder.AppendLine("- None");
            }
            else
            {
                foreach (var item in items)
                {
                    builder.AppendLine($"- {item}");
                }
            }

            builder.AppendLine();
        }

        private static string ResolveExistingPath(string path, string defaultFileName, string description)
        {
            var resolved = ResolvePath(path, defaultFileName);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"{description} not found.", resolved);
            }

            return resolved;
        }

        private static string ResolvePath(string path, string defaultFileName)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("A file path or directory is required.", nameof(path));
            }

            var fullPath = Path.GetFullPath(path);
            return Directory.Exists(fullPath) || string.IsNullOrEmpty(Path.GetExtension(fullPath))
                ? Path.Combine(fullPath, defaultFileName)
                : fullPath;
        }

        private static void SaveJson<T>(string path, T value, bool overwrite)
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (File.Exists(path) && !overwrite)
            {
                throw new IOException($"The file already exists: {path}");
            }

            File.WriteAllText(path, JsonSerializer.Serialize(value, IndentedOptions) + Environment.NewLine, new UTF8Encoding(false));
        }

        private static JsonSerializerOptions CreateOptions(bool writeIndented)
        {
            var options = new JsonSerializerOptions
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.Never,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                PropertyNameCaseInsensitive = false,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = writeIndented
            };
            options.Converters.Add(new JsonStringEnumConverter());
            return options;
        }
    }
}
