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
using System.Text.RegularExpressions;
using System.Xml.Linq;
using CsomFile = Microsoft.SharePoint.Client.File;

namespace PnP.Framework.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationService
    {
        private const int PublishingPagesListTemplate = 850;

        private static readonly string[] AdditionalFieldNames =
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

        public EnterpriseWikiMigrationPackage Capture(
            ClientContext sourceContext,
            ClientContext targetContext,
            EnterpriseWikiCaptureOptions options)
        {
            if (sourceContext == null)
            {
                throw new ArgumentNullException(nameof(sourceContext));
            }

            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            ValidateOptions(options);
            var sourceWeb = sourceContext.Web;
            var targetWeb = targetContext.Web;
            sourceContext.Load(sourceWeb, web => web.Url, web => web.ServerRelativeUrl, web => web.Title);
            targetContext.Load(targetWeb, web => web.Url, web => web.ServerRelativeUrl, web => web.Title, web => web.WebTemplate, web => web.Configuration);
            sourceContext.ExecuteQueryRetry();
            targetContext.ExecuteQueryRetry();

            var sourcePagePath = NormalizePagePath(sourceWeb.ServerRelativeUrl, options.SourcePageServerRelativeUrl, "Pages");
            var targetPagePath = NormalizePagePath(targetWeb.ServerRelativeUrl, options.TargetPageServerRelativeUrl, "Pages");
            var blockers = new List<string>();
            var warnings = new List<string>();

            var sourceCapture = CaptureSourcePage(sourceContext, sourcePagePath, options, blockers, warnings);
            var dependencies = CaptureDependencies(
                sourceContext,
                sourceCapture.Identity,
                sourceCapture.PublishingPageContent,
                sourceCapture.WebParts,
                new Uri(targetWeb.Url),
                options,
                blockers,
                warnings);
            var targetProbe = ProbeTarget(targetContext, targetPagePath, dependencies, blockers, warnings);

            if (!sourceCapture.Identity.PageLayoutUrl.EndsWith("/EnterpriseWiki.aspx", StringComparison.OrdinalIgnoreCase))
            {
                blockers.Add($"The source page uses layout '{sourceCapture.Identity.PageLayoutUrl}'. The v1 exact profile requires EnterpriseWiki.aspx.");
            }

            if (sourceCapture.Security.HasUniqueRoleAssignments && options.RequireInheritedPermissions)
            {
                blockers.Add("The source page has unique role assignments. The v1 exact profile requires inherited permissions.");
            }

            var managedMetadata = sourceCapture.Fields
                .Where(field => field.Kind == EnterpriseWikiFieldValueKind.Taxonomy || field.Kind == EnterpriseWikiFieldValueKind.TaxonomyCollection)
                .Where(field => !string.IsNullOrWhiteSpace(field.Value))
                .ToArray();
            if (managedMetadata.Length > 0 && options.BlockOnManagedMetadata)
            {
                blockers.Add($"The source page contains {managedMetadata.Length} non-empty managed metadata field value(s), but no reviewed target term mapping was supplied.");
            }

            var afterFence = ReadFence(sourceContext, sourcePagePath);
            if (!FenceEquals(sourceCapture.SourceFence, afterFence))
            {
                blockers.Add("The source page changed while it was being captured. Discard this package and capture again.");
            }

            var snapshot = new EnterpriseWikiSnapshot
            {
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
                SourceFence = sourceCapture.SourceFence
            };
            var snapshotDigest = EnterpriseWikiPackageSerializer.ComputeSnapshotDigest(snapshot);
            var plan = new EnterpriseWikiMigrationPlan
            {
                SourceSnapshotDigest = snapshotDigest,
                SourceWebUrl = sourceCapture.Identity.WebUrl,
                SourcePageServerRelativeUrl = sourceCapture.Identity.PageServerRelativeUrl,
                TargetWebUrl = targetWeb.Url.TrimEnd('/'),
                TargetWebServerRelativeUrl = targetWeb.ServerRelativeUrl,
                TargetPageServerRelativeUrl = targetPagePath,
                PageLayoutName = "EnterpriseWiki",
                Publish = options.Publish,
                CreateOnly = true,
                TargetProbe = targetProbe,
                Replacements = BuildReplacements(sourceCapture.Identity, targetWeb.Url, targetWeb.ServerRelativeUrl),
                StorageAssertions = BuildStorageAssertions(snapshot, targetPagePath),
                BrowserAssertions = BrowserAssertions.ToList(),
                Blockers = blockers.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()
            };
            var planDigest = EnterpriseWikiPackageSerializer.ComputePlanDigest(plan);
            var state = plan.IsExecutable ? EnterpriseWikiPackageState.ApprovalReady : EnterpriseWikiPackageState.Blocked;
            var report = new EnterpriseWikiCustomerReport
            {
                Summary = plan.IsExecutable
                    ? "Capture, target analysis, and deterministic planning completed. Copy requires explicit approval of the plan digest."
                    : "The package is sealed for review but cannot be copied until every blocker is resolved and a new package is captured.",
                CapturedIngredients = new List<string>
                {
                    "Page/file/list item identity and source stability fence",
                    "Enterprise Wiki content type and EnterpriseWiki.aspx layout",
                    "PublishingPageContent and selected publishing/search metadata",
                    $"{snapshot.WebParts.Count} shared Web Part export(s) with zone placement",
                    $"{snapshot.Dependencies.Count} authored dependency/link decision(s)",
                    "Page security inheritance and lifecycle state",
                    "Target publishing library, Enterprise Wiki content type, layout, and create-only collision probe"
                },
                Blockers = plan.Blockers.ToList(),
                Warnings = plan.Warnings.ToList()
            };

            return new EnterpriseWikiMigrationPackage
            {
                CreatedAtUtc = DateTimeOffset.UtcNow,
                State = state,
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = snapshotDigest,
                PlanDigest = planDigest,
                Report = report
            };
        }

        public EnterpriseWikiCopyReceipt Copy(
            ClientContext targetContext,
            EnterpriseWikiMigrationPackage package,
            string approvedPlanDigest)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            EnterpriseWikiPackageSerializer.Validate(package);
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
                package.Snapshot.Dependencies,
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

            var materialized = MaterializeDependencies(targetContext, package.Snapshot.Dependencies);
            var rewrittenContent = RewriteContent(package.Snapshot.PublishingPageContent, package.Plan.Replacements);
            var pages = targetWeb.GetPagesLibrary();
            targetContext.Load(pages, list => list.EnableModeration, list => list.ForceCheckout);
            targetContext.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            targetContext.ExecuteQueryRetry();
            var targetDirectory = GetDirectoryName(package.Plan.TargetPageServerRelativeUrl);
            if (!string.Equals(targetDirectory, pages.RootFolder.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new NotSupportedException("The v1 Enterprise Wiki copy profile supports pages in the root of the target Pages library only.");
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
            ApplyAdditionalFields(targetContext, pages, targetItem, package.Snapshot.Fields, package.Plan.Replacements, preflightWarnings);
            targetItem.Update();
            targetContext.ExecuteQueryRetry();

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
            if (targetFile.CheckOutType != CheckOutType.None)
            {
                targetFile.CheckIn("PnP Enterprise Wiki copy", CheckinType.MajorCheckIn);
                targetContext.ExecuteQueryRetry();
            }

            if (package.Plan.Publish)
            {
                targetFile.Publish("PnP Enterprise Wiki copy");
                targetContext.ExecuteQueryRetry();
                if (pages.EnableModeration)
                {
                    targetFile.Approve("PnP Enterprise Wiki copy");
                    targetContext.ExecuteQueryRetry();
                }
            }

            return ReadCopyReceipt(
                targetContext,
                package,
                approvedPlanDigest,
                startedAt,
                materialized,
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
                return $"type '{typeName}' is not supported by the v1 deterministic import profile and must be replaced or explicitly mapped";
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
                return $"type '{typeName}' is bound to a source list; the v1 profile requires a reviewed target-list and view mapping before copy";
            }

            return null;
        }

        private static void ValidateOptions(EnterpriseWikiCaptureOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.SourcePageServerRelativeUrl))
            {
                throw new ArgumentException("A source page path is required.", nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.TargetPageServerRelativeUrl))
            {
                throw new ArgumentException("A target page path is required.", nameof(options));
            }

            if (options.MaximumDependencyBytes <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(options), "MaximumDependencyBytes must be greater than zero.");
            }
        }

        private static SourceCapture CaptureSourcePage(
            ClientContext context,
            string pagePath,
            EnterpriseWikiCaptureOptions options,
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
            var security = CaptureSecurity(context, item);
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
                field => field.InternalName,
                field => field.TypeAsString,
                field => field.ReadOnlyField,
                field => field.Hidden,
                field => field.Required));
            context.ExecuteQueryRetry();
            var fieldByName = parentList.Fields.ToDictionary(field => field.InternalName, StringComparer.OrdinalIgnoreCase);
            var result = new List<EnterpriseWikiFieldValueSnapshot>();
            foreach (var fieldName in AdditionalFieldNames)
            {
                if (!fieldByName.TryGetValue(fieldName, out var field)
                    || !item.FieldValues.TryGetValue(fieldName, out var value)
                    || value == null)
                {
                    continue;
                }

                var snapshot = SerializeFieldValue(field, value);
                result.Add(snapshot);
                if (snapshot.Kind == EnterpriseWikiFieldValueKind.Unsupported && !string.IsNullOrWhiteSpace(snapshot.Value))
                {
                    warnings.Add($"Field '{fieldName}' has unsupported value type '{value.GetType().FullName}' and will be evidence-only.");
                }
            }

            return result;
        }

        private static EnterpriseWikiFieldValueSnapshot SerializeFieldValue(Field field, object value)
        {
            var kind = EnterpriseWikiFieldValueKind.Unsupported;
            string serialized;
            if (value is FieldUrlValue url)
            {
                kind = EnterpriseWikiFieldValueKind.Url;
                serialized = $"{url.Url ?? string.Empty}\n{url.Description ?? string.Empty}";
            }
            else if (value is TaxonomyFieldValue taxonomy)
            {
                kind = EnterpriseWikiFieldValueKind.Taxonomy;
                serialized = $"{taxonomy.Label}|{taxonomy.TermGuid}|{taxonomy.WssId}";
            }
            else if (value is TaxonomyFieldValueCollection taxonomyCollection)
            {
                kind = EnterpriseWikiFieldValueKind.TaxonomyCollection;
                serialized = string.Join(";#", taxonomyCollection.Select(term => $"{term.Label}|{term.TermGuid}|{term.WssId}"));
            }
            else if (value is DateTime dateTime)
            {
                kind = EnterpriseWikiFieldValueKind.DateTime;
                serialized = dateTime.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }
            else if (value is bool boolean)
            {
                kind = EnterpriseWikiFieldValueKind.Boolean;
                serialized = boolean ? "true" : "false";
            }
            else if (value is byte || value is short || value is int || value is long || value is float || value is double || value is decimal)
            {
                kind = EnterpriseWikiFieldValueKind.Number;
                serialized = Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            else if (value is string text)
            {
                kind = EnterpriseWikiFieldValueKind.String;
                serialized = text;
            }
            else
            {
                serialized = Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return new EnterpriseWikiFieldValueSnapshot
            {
                InternalName = field.InternalName,
                TypeAsString = field.TypeAsString,
                Kind = kind,
                Value = serialized,
                ReadOnly = field.ReadOnlyField,
                Hidden = field.Hidden,
                Required = field.Required
            };
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

        private static EnterpriseWikiSecuritySnapshot CaptureSecurity(ClientContext context, ListItem item)
        {
            var result = new EnterpriseWikiSecuritySnapshot
            {
                HasUniqueRoleAssignments = item.HasUniqueRoleAssignments
            };
            if (!item.HasUniqueRoleAssignments)
            {
                return result;
            }

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

            return result;
        }

        private static List<EnterpriseWikiDependencySnapshot> CaptureDependencies(
            ClientContext sourceContext,
            EnterpriseWikiPageIdentity source,
            string publishingPageContent,
            IEnumerable<EnterpriseWikiWebPartSnapshot> webParts,
            Uri targetWebUri,
            EnterpriseWikiCaptureOptions options,
            ICollection<string> blockers,
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

                var dependency = PlanDependency(sourceContext, sourceWebUri, targetWebUri, candidate, absoluteUri, options, blockers, warnings);
                result.Add(dependency);
            }

            return result;
        }

        private static EnterpriseWikiDependencySnapshot PlanDependency(
            ClientContext sourceContext,
            Uri sourceWebUri,
            Uri targetWebUri,
            DependencyCandidate candidate,
            Uri absoluteUri,
            EnterpriseWikiCaptureOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var dependency = new EnterpriseWikiDependencySnapshot
            {
                OriginalValue = candidate.Value,
                SourceAbsoluteUrl = absoluteUri.AbsoluteUri,
                Consumer = candidate.Consumer,
                Kind = candidate.Kind,
                Disposition = EnterpriseWikiDependencyDisposition.PreserveExternal
            };
            if (!string.Equals(sourceWebUri.Host, absoluteUri.Host, StringComparison.OrdinalIgnoreCase))
            {
                if (candidate.IsRenderableResource && !options.AllowExternalResourceReferences)
                {
                    dependency.Disposition = EnterpriseWikiDependencyDisposition.Block;
                    dependency.Diagnostics.Add("External renderable resources are blocked by capture policy.");
                    blockers.Add($"External resource '{absoluteUri}' is blocked by policy.");
                }
                return dependency;
            }

            var sourcePath = Uri.UnescapeDataString(absoluteUri.AbsolutePath);
            var sourceWebPath = Uri.UnescapeDataString(sourceWebUri.AbsolutePath).TrimEnd('/');
            var targetWebPath = Uri.UnescapeDataString(targetWebUri.AbsolutePath).TrimEnd('/');
            dependency.SourceServerRelativeUrl = sourcePath;
            string targetPath;
            if (IsPathWithin(sourcePath, sourceWebPath))
            {
                targetPath = targetWebPath + sourcePath.Substring(sourceWebPath.Length);
            }
            else
            {
                targetPath = sourcePath;
            }

            dependency.TargetServerRelativeUrl = targetPath;
            dependency.TargetAbsoluteUrl = targetWebUri.GetLeftPart(UriPartial.Authority) + EncodePath(targetPath) + absoluteUri.Query + absoluteUri.Fragment;
            dependency.Disposition = EnterpriseWikiDependencyDisposition.RewriteToTarget;
            if (!candidate.IsRenderableResource || IsSharePointRuntimePath(sourcePath))
            {
                return dependency;
            }

            if (candidate.Kind == EnterpriseWikiDependencyKind.IFrame)
            {
                dependency.Disposition = EnterpriseWikiDependencyDisposition.Block;
                dependency.Diagnostics.Add("Same-tenant iframe dependencies require a separately reviewed page/application profile.");
                blockers.Add($"Iframe dependency '{absoluteUri}' is unsupported by the v1 exact profile.");
                return dependency;
            }

            if (!IsPathWithin(sourcePath, sourceWebPath))
            {
                dependency.Disposition = EnterpriseWikiDependencyDisposition.Block;
                dependency.Diagnostics.Add("The resource is outside the captured source web and cannot be safely materialized inside the approved target web.");
                blockers.Add($"Same-tenant resource '{absoluteUri}' is outside the source web boundary.");
                return dependency;
            }

            try
            {
                var payload = CaptureFilePayload(sourceContext, sourcePath, options.MaximumDependencyBytes);
                dependency.ContentBase64 = Convert.ToBase64String(payload);
                dependency.ContentLength = payload.LongLength;
                dependency.ContentSha256 = ComputeBytesSha256(payload);
                dependency.Disposition = EnterpriseWikiDependencyDisposition.MaterializeAtTarget;
            }
            catch (Exception exception) when (exception is ServerException || exception is InvalidOperationException || exception is IOException)
            {
                dependency.Disposition = EnterpriseWikiDependencyDisposition.Block;
                dependency.Diagnostics.Add(exception.Message);
                blockers.Add($"Resource '{absoluteUri}' could not be captured: {exception.Message}");
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

        private static EnterpriseWikiTargetProbe ProbeTarget(
            ClientContext context,
            string targetPagePath,
            IEnumerable<EnterpriseWikiDependencySnapshot> dependencies,
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

            context.Load(pages, list => list.BaseTemplate);
            context.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            context.Load(pages.ContentTypes, contentTypes => contentTypes.Include(contentType => contentType.Id, contentType => contentType.Name));
            context.ExecuteQueryRetry();
            probe.PagesLibraryBaseTemplate = pages.BaseTemplate;
            probe.PagesLibraryServerRelativeUrl = pages.RootFolder.ServerRelativeUrl;
            if (pages.BaseTemplate != PublishingPagesListTemplate)
            {
                blockers.Add($"The target Pages library has base template {pages.BaseTemplate}; publishing Pages template {PublishingPagesListTemplate} is required.");
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
                blockers.Add($"The target page must be placed in the root of '{expectedDirectory}' for the v1 profile.");
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

        private static int MaterializeDependencies(ClientContext context, IEnumerable<EnterpriseWikiDependencySnapshot> dependencies)
        {
            var web = context.Web;
            web.EnsureProperty(value => value.ServerRelativeUrl);
            var count = 0;
            foreach (var dependency in dependencies
                         .Where(item => item.Disposition == EnterpriseWikiDependencyDisposition.MaterializeAtTarget)
                         .GroupBy(item => item.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                var bytes = Convert.FromBase64String(dependency.ContentBase64 ?? string.Empty);
                if (!string.Equals(ComputeBytesSha256(bytes), dependency.ContentSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Dependency payload digest mismatch: {dependency.TargetServerRelativeUrl}");
                }

                if (!IsPathWithin(dependency.TargetServerRelativeUrl, web.ServerRelativeUrl))
                {
                    throw new InvalidOperationException($"Dependency target escapes the target web boundary: {dependency.TargetServerRelativeUrl}");
                }

                var relativePath = dependency.TargetServerRelativeUrl.Substring(web.ServerRelativeUrl.TrimEnd('/').Length).TrimStart('/');
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
                    throw new InvalidOperationException($"SharePoint did not persist dependency '{dependency.TargetServerRelativeUrl}'.");
                }
                count++;
            }

            return count;
        }

        private static void ApplyAdditionalFields(
            ClientContext context,
            Microsoft.SharePoint.Client.List pages,
            ListItem targetItem,
            IEnumerable<EnterpriseWikiFieldValueSnapshot> fields,
            IEnumerable<EnterpriseWikiTextReplacement> replacements,
            ICollection<string> warnings)
        {
            var applicableFields = new List<EnterpriseWikiFieldValueSnapshot>();
            foreach (var field in fields)
            {
                if (field.ReadOnly)
                {
                    continue;
                }

                Field targetField;
                try
                {
                    targetField = pages.Fields.GetByInternalNameOrTitle(field.InternalName);
                    context.Load(targetField, value => value.InternalName, value => value.ReadOnlyField);
                    context.ExecuteQueryRetry();
                }
                catch (ServerException exception)
                {
                    warnings.Add($"Target field '{field.InternalName}' could not be resolved and was not applied: {exception.Message}");
                    continue;
                }

                if (targetField.ReadOnlyField)
                {
                    continue;
                }

                applicableFields.Add(field);
            }

            foreach (var field in applicableFields)
            {
                switch (field.Kind)
                {
                    case EnterpriseWikiFieldValueKind.String:
                        targetItem[field.InternalName] = RewriteContent(field.Value, replacements);
                        break;
                    case EnterpriseWikiFieldValueKind.Boolean:
                        targetItem[field.InternalName] = string.Equals(field.Value, "true", StringComparison.OrdinalIgnoreCase);
                        break;
                    case EnterpriseWikiFieldValueKind.Number:
                        if (double.TryParse(field.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                        {
                            targetItem[field.InternalName] = number;
                        }
                        break;
                    case EnterpriseWikiFieldValueKind.DateTime:
                        if (DateTime.TryParse(field.Value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dateTime))
                        {
                            targetItem[field.InternalName] = dateTime;
                        }
                        break;
                    case EnterpriseWikiFieldValueKind.Url:
                        var parts = (field.Value ?? string.Empty).Split(new[] { '\n' }, 2);
                        targetItem[field.InternalName] = new FieldUrlValue
                        {
                            Url = RewriteContent(parts[0], replacements),
                            Description = parts.Length > 1 ? parts[1] : string.Empty
                        };
                        break;
                    case EnterpriseWikiFieldValueKind.Taxonomy:
                    case EnterpriseWikiFieldValueKind.TaxonomyCollection:
                        warnings.Add($"Managed metadata field '{field.InternalName}' was captured but not applied because target term mapping is outside the v1 profile.");
                        break;
                    default:
                        warnings.Add($"Field '{field.InternalName}' was evidence-only and was not applied.");
                        break;
                }
            }
        }

        private static EnterpriseWikiCopyReceipt ReadCopyReceipt(
            ClientContext targetContext,
            EnterpriseWikiMigrationPackage package,
            string approvedPlanDigest,
            DateTimeOffset startedAt,
            int materializedDependencyCount,
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
  </ViewFields>
  <RowLimit>1</RowLimit>
</View>"
                });
                verificationContext.Load(file, value => value.Exists, value => value.UniqueId, value => value.UIVersionLabel);
                verificationContext.Load(items);
                verificationContext.ExecuteQueryRetry();
                if (!file.Exists)
                {
                    throw new InvalidOperationException("Fresh target readback could not find the copied page.");
                }

                var item = items.SingleOrDefault();
                if (item == null)
                {
                    throw new InvalidOperationException("Fresh target readback could not find the copied page list item.");
                }

                var content = GetFieldString(item, "PublishingPageContent") ?? string.Empty;
                var contentTypeId = GetFieldString(item, "ContentTypeId") ?? string.Empty;
                var webParts = verificationContext.Web.GetWebParts(package.Plan.TargetPageServerRelativeUrl).ToArray();
                var persistedDigest = EnterpriseWikiPackageSerializer.ComputeSha256(content);
                var receiptWarnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList();
                var storageContentEqual = string.Equals(persistedDigest, package.Snapshot.PublishingPageContentSha256, StringComparison.OrdinalIgnoreCase);
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

                var readbackPassed = IsEnterpriseWikiContentType(contentTypeId)
                    && webParts.Length == package.Snapshot.WebParts.Count
                    && (!expectedContentPresent || persistedContentPresent);
                return new EnterpriseWikiCopyReceipt
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
                    PersistedPublishingPageContentSha256 = persistedDigest,
                    StorageContentEqual = storageContentEqual,
                    ImportedWebPartCount = webParts.Length,
                    MaterializedDependencyCount = materializedDependencyCount,
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

        private static IList<string> BuildStorageAssertions(EnterpriseWikiSnapshot snapshot, string targetPagePath)
        {
            var result = new List<string>
            {
                $"target-page={targetPagePath}",
                "fresh-read-target-file-identity",
                "fresh-read-target-enterprise-wiki-content-type",
                "fresh-read-target-version-and-lifecycle",
                $"source-publishing-page-content-sha256={snapshot.PublishingPageContentSha256}",
                $"expected-shared-webparts={snapshot.WebParts.Count}"
            };
            result.AddRange(snapshot.Dependencies
                .Where(item => item.Disposition == EnterpriseWikiDependencyDisposition.MaterializeAtTarget)
                .Select(item => $"dependency={item.TargetServerRelativeUrl}|sha256={item.ContentSha256}"));
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
