using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    public sealed class TopologyPlanner
    {
        public const string SiteOriginalIdentifierPropertyName = "pnp_reserved_site_original_identifier";
        public const string WebOriginalIdentifierPropertyName = "pnp_reserved_web_original_identifier";
        public const string SitePlanDigestPropertyName = "pnp_reserved_site_migration_digest";
        public const string WebPlanDigestPropertyName = "pnp_reserved_web_migration_digest";

        public TopologyPlanBuildResult Build(
            IEnumerable<SourceSiteCollectionSnapshot> sourceSites,
            IEnumerable<TargetSiteCollectionSpec> targetSites,
            TopologyPlanningPolicy policy = null)
        {
            if (sourceSites == null)
            {
                throw new ArgumentNullException(nameof(sourceSites));
            }
            if (targetSites == null)
            {
                throw new ArgumentNullException(nameof(targetSites));
            }

            policy = policy ?? new TopologyPlanningPolicy();
            var sources = sourceSites.ToArray();
            var targets = targetSites.ToArray();
            var issues = new List<MigrationIssue>();
            AddDuplicateIssues(sources, targets, policy, issues);
            var targetBySource = targets
                .GroupBy(value => value.SourceSiteId)
                .ToDictionary(group => group.Key, group => group.ToArray());
            var overrideByWeb = (policy.WebOverrides ?? new List<TargetWebOverride>())
                .GroupBy(value => value.SourceWebId)
                .ToDictionary(group => group.Key, group => group.ToArray());
            var sitePlans = new List<SiteCollectionMappingPlan>();
            foreach (var source in sources.OrderBy(value => value.SiteCollectionUrl, StringComparer.OrdinalIgnoreCase))
            {
                ValidateSourceSite(source, issues);
                TargetSiteCollectionSpec[] candidates;
                if (!targetBySource.TryGetValue(source.SiteId, out candidates) || candidates.Length != 1)
                {
                    AddBlocker(issues, "MissingTargetSiteSpec", "source-site:" + source.SiteId.ToString("D"), "Exactly one target Site Collection specification is required for this source Site Collection.");
                    continue;
                }

                var target = candidates[0];
                ValidateTargetSite(target, issues);
                var webPlans = BuildWebPlans(source, target, policy, overrideByWeb, issues);
                sitePlans.Add(new SiteCollectionMappingPlan
                {
                    SourceSiteId = source.SiteId,
                    SourceSiteCollectionUrl = NormalizeAbsoluteUrl(source.SiteCollectionUrl),
                    TargetMode = target.Mode,
                    PreferredTargetSiteCollectionUrl = NormalizeAbsoluteUrl(target.TargetSiteUrl),
                    TargetSiteCollectionUrl = NormalizeAbsoluteUrl(target.TargetSiteUrl),
                    ExpectedTargetSiteId = target.ExpectedTargetSiteId,
                    TargetTitle = target.Title,
                    TargetOwner = target.Owner,
                    TargetTemplate = target.Template,
                    TargetLanguage = target.Language,
                    TargetTimeZone = target.TimeZone,
                    OriginalIdentifier = SiteOriginalIdentifier(source.SiteId),
                    Webs = webPlans
                });
            }

            var orderedIssues = issues
                .OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Subject, StringComparer.Ordinal)
                .ThenBy(value => value.Message, StringComparer.Ordinal)
                .ToList();
            if (orderedIssues.Any(value => value.Severity == MigrationIssueSeverity.Blocker || value.Severity == MigrationIssueSeverity.Error))
            {
                return new TopologyPlanBuildResult { Issues = orderedIssues };
            }

            var plan = new TopologyPlan
            {
                SiteCollections = sitePlans.OrderBy(value => value.SourceSiteId).ToList()
            };
            plan.PlanDigest = ComputeDigest(plan);
            return new TopologyPlanBuildResult { Plan = plan, Issues = orderedIssues };
        }

        public static string ComputeDigest(TopologyPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            // Digesting must be safe for a sealed plan shared by parallel page
            // assessments. Mutating PlanDigest, even briefly, lets another reader
            // observe an invalid plan or seal an assessment without topology
            // provenance. A shallow envelope is sufficient because only the
            // self-referential top-level digest is excluded from canonical form.
            var canonical = new TopologyPlan
            {
                SchemaVersion = plan.SchemaVersion,
                SiteCollections = plan.SiteCollections,
                PlanDigest = null
            };
            return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(canonical));
        }

        public static string ComputeWebMappingDigest(WebMappingPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(plan));
        }

        public static string ComputeSiteMappingDigest(SiteCollectionMappingPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            var stablePlan = new SiteCollectionMappingPlan
            {
                SourceSiteId = plan.SourceSiteId,
                SourceSiteCollectionUrl = plan.SourceSiteCollectionUrl,
                TargetMode = TargetSiteMode.CreateTargetSite,
                PreferredTargetSiteCollectionUrl = plan.PreferredTargetSiteCollectionUrl,
                TargetSiteCollectionUrl = plan.TargetSiteCollectionUrl,
                TargetSiteCollisionResolved = plan.TargetSiteCollisionResolved,
                TargetSiteResolutionReason = plan.TargetSiteResolutionReason,
                ExpectedTargetSiteId = null,
                TargetTitle = plan.TargetTitle,
                TargetOwner = plan.TargetOwner,
                TargetTemplate = plan.TargetTemplate,
                TargetLanguage = plan.TargetLanguage,
                TargetTimeZone = plan.TargetTimeZone,
                OriginalIdentifier = plan.OriginalIdentifier,
                Webs = plan.Webs
            };
            return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(stablePlan));
        }

        public static string MapWebOwnedServerRelativePath(
            string sourceObjectServerRelativeUrl,
            string sourceOwnerWebServerRelativeUrl,
            string targetOwnerWebServerRelativeUrl)
        {
            var sourceObject = NormalizeServerRelativePath(sourceObjectServerRelativeUrl);
            var sourceWeb = NormalizeServerRelativePath(sourceOwnerWebServerRelativeUrl).TrimEnd('/');
            var targetWeb = NormalizeServerRelativePath(targetOwnerWebServerRelativeUrl).TrimEnd('/');
            if (!string.Equals(sourceObject, sourceWeb, StringComparison.OrdinalIgnoreCase)
                && !sourceObject.StartsWith(sourceWeb + "/", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("The source object is outside its declared owner Web.", nameof(sourceObjectServerRelativeUrl));
            }

            return targetWeb + sourceObject.Substring(sourceWeb.Length);
        }

        public static string SiteOriginalIdentifier(Guid siteId)
        {
            return "urn:pnp:spo-site:v1:" + siteId.ToString("D");
        }

        public static string WebOriginalIdentifier(Guid siteId, Guid webId)
        {
            return "urn:pnp:spo-web:v1:" + siteId.ToString("D") + ":" + webId.ToString("D");
        }

        private static IList<WebMappingPlan> BuildWebPlans(
            SourceSiteCollectionSnapshot source,
            TargetSiteCollectionSpec target,
            TopologyPlanningPolicy policy,
            IDictionary<Guid, TargetWebOverride[]> overrideByWeb,
            ICollection<MigrationIssue> issues)
        {
            var issueCount = issues.Count;
            var webs = (source.Webs ?? new List<SourceWebSnapshot>())
                .GroupBy(value => value.WebId)
                .ToDictionary(group => group.Key, group => group.ToArray());
            SourceWebSnapshot[] roots;
            if (!webs.TryGetValue(source.RootWebId, out roots) || roots.Length != 1)
            {
                AddBlocker(issues, "MissingRootWeb", "source-site:" + source.SiteId.ToString("D"), "The source topology must contain exactly one declared root Web.");
                return new List<WebMappingPlan>();
            }

            foreach (var web in source.Webs ?? new List<SourceWebSnapshot>())
            {
                ValidateSourceWeb(source, web, issues);
                SourceWebSnapshot[] parents;
                if (web.WebId != source.RootWebId
                    && (!web.ParentWebId.HasValue || !webs.TryGetValue(web.ParentWebId.Value, out parents) || parents.Length != 1))
                {
                    AddBlocker(issues, "MissingParentWeb", "source-web:" + web.WebId.ToString("D"), "Every child Web must reference one captured parent Web in the same Site Collection.");
                }
            }
            if (issues.Count > issueCount)
            {
                return new List<WebMappingPlan>();
            }

            var result = new Dictionary<Guid, WebMappingPlan>();
            var pending = source.Webs.OrderBy(value => PathDepth(value.ServerRelativeUrl)).ToList();
            while (pending.Count > 0)
            {
                var progressed = false;
                foreach (var web in pending.ToArray())
                {
                    if (web.WebId == source.RootWebId)
                    {
                        var targetRoot = NormalizeAbsoluteUrl(target.TargetSiteUrl);
                        result[web.WebId] = CreateRootMapping(source, web, target, targetRoot);
                        pending.Remove(web);
                        progressed = true;
                        continue;
                    }

                    WebMappingPlan parent;
                    if (!web.ParentWebId.HasValue || !result.TryGetValue(web.ParentWebId.Value, out parent))
                    {
                        continue;
                    }

                    TargetWebOverride[] overrides;
                    var targetOverride = overrideByWeb.TryGetValue(web.WebId, out overrides) && overrides.Length == 1 ? overrides[0] : null;
                    var hasOverride = targetOverride != null && !string.IsNullOrWhiteSpace(targetOverride.TargetUrlSegment);
                    var relativePath = !hasOverride
                        ? DefaultTargetRelativePath(web, webs[web.ParentWebId.Value].Single())
                        : targetOverride.TargetUrlSegment;
                    // A SharePoint child Web is created relative to its captured direct parent.
                    // Therefore its creation URL must be exactly one path segment. Accepting a
                    // multi-segment value here would silently invent missing intermediate Webs
                    // and make the materializer call Webs.Add with an unsupported URL.
                    var pathIsSafe = IsSafeWebSegment(relativePath);
                    if (!pathIsSafe)
                    {
                        AddBlocker(issues, "InvalidTargetWebPath", "source-web:" + web.WebId.ToString("D"), "Target Web relative path '" + relativePath + "' is not safe.");
                        pending.Remove(web);
                        progressed = true;
                        continue;
                    }

                    var targetPath = parent.TargetServerRelativeUrl.TrimEnd('/') + "/" + relativePath;
                    var targetUrl = new Uri(new Uri(target.TargetSiteUrl).GetLeftPart(UriPartial.Authority) + targetPath).AbsoluteUri.TrimEnd('/');
                    result[web.WebId] = new WebMappingPlan
                    {
                        Kind = TopologyNodeKind.ChildWeb,
                        SourceSiteId = source.SiteId,
                        SourceWebId = web.WebId,
                        SourceParentWebId = web.ParentWebId,
                        SourceSiteCollectionUrl = NormalizeAbsoluteUrl(source.SiteCollectionUrl),
                        SourceWebUrl = NormalizeAbsoluteUrl(web.WebUrl),
                        SourceServerRelativeUrl = NormalizeServerRelativePath(web.ServerRelativeUrl),
                        TargetSiteCollectionUrl = NormalizeAbsoluteUrl(target.TargetSiteUrl),
                        PreferredTargetWebUrl = targetUrl,
                        TargetWebUrl = targetUrl,
                        PreferredTargetServerRelativeUrl = NormalizeServerRelativePath(targetPath),
                        TargetServerRelativeUrl = NormalizeServerRelativePath(targetPath),
                        TargetParentWebUrl = parent.TargetWebUrl,
                        TargetTitle = targetOverride == null || string.IsNullOrWhiteSpace(targetOverride.TargetTitle) ? web.Title : targetOverride.TargetTitle,
                        TargetTemplate = targetOverride != null && !string.IsNullOrWhiteSpace(targetOverride.TargetTemplate)
                            ? targetOverride.TargetTemplate
                            : policy.PreserveSourceChildWebTemplate && !string.IsNullOrWhiteSpace(web.WebTemplate)
                                ? web.WebTemplate
                                : policy.DefaultChildWebTemplate,
                        TargetConfiguration = targetOverride != null && targetOverride.TargetConfiguration.HasValue
                            ? targetOverride.TargetConfiguration.Value
                            : policy.PreserveSourceChildWebTemplate
                                ? web.Configuration
                                : policy.DefaultChildWebConfiguration,
                        OriginalIdentifier = WebOriginalIdentifier(source.SiteId, web.WebId)
                    };
                    pending.Remove(web);
                    progressed = true;
                }

                if (!progressed)
                {
                    foreach (var web in pending)
                    {
                        AddBlocker(issues, "MissingParentWeb", "source-web:" + web.WebId.ToString("D"), "The source Web graph is cyclic or its parent closure is incomplete.");
                    }
                    break;
                }
            }

            foreach (var collision in result.Values.GroupBy(value => value.TargetWebUrl, StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1))
            {
                AddBlocker(issues, "TargetPathCollision", collision.Key, "More than one source Web maps to the same target Web URL.");
            }
            return result.Values.OrderBy(value => PathDepth(value.TargetServerRelativeUrl)).ThenBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static WebMappingPlan CreateRootMapping(SourceSiteCollectionSnapshot source, SourceWebSnapshot web, TargetSiteCollectionSpec target, string targetRoot)
        {
            return new WebMappingPlan
            {
                Kind = TopologyNodeKind.SiteCollectionRoot,
                SourceSiteId = source.SiteId,
                SourceWebId = web.WebId,
                SourceSiteCollectionUrl = NormalizeAbsoluteUrl(source.SiteCollectionUrl),
                SourceWebUrl = NormalizeAbsoluteUrl(web.WebUrl),
                SourceServerRelativeUrl = NormalizeServerRelativePath(web.ServerRelativeUrl),
                TargetSiteCollectionUrl = targetRoot,
                PreferredTargetWebUrl = targetRoot,
                TargetWebUrl = targetRoot,
                PreferredTargetServerRelativeUrl = NormalizeServerRelativePath(new Uri(targetRoot).AbsolutePath),
                TargetServerRelativeUrl = NormalizeServerRelativePath(new Uri(targetRoot).AbsolutePath),
                TargetTitle = target.Title,
                TargetTemplate = target.Template,
                TargetConfiguration = 0,
                OriginalIdentifier = WebOriginalIdentifier(source.SiteId, web.WebId)
            };
        }

        private static void AddDuplicateIssues(SourceSiteCollectionSnapshot[] sources, TargetSiteCollectionSpec[] targets, TopologyPlanningPolicy policy, ICollection<MigrationIssue> issues)
        {
            foreach (var group in sources.GroupBy(value => value.SiteId).Where(group => group.Count() != 1))
            {
                AddBlocker(issues, "DuplicateSourceIdentity", "source-site:" + group.Key.ToString("D"), "The source topology contains more than one record for the same Site Collection ID.");
            }
            foreach (var group in targets.GroupBy(value => value.SourceSiteId).Where(group => group.Count() != 1))
            {
                AddBlocker(issues, "DuplicateTargetSiteSpec", "target-site-for:" + group.Key.ToString("D"), "Exactly one target Site Collection specification is required per source Site Collection.");
            }
            foreach (var group in (policy.WebOverrides ?? new List<TargetWebOverride>()).GroupBy(value => value.SourceWebId).Where(group => group.Count() != 1))
            {
                AddBlocker(issues, "InvalidTargetWebOverride", "target-web-override:" + group.Key.ToString("D"), "A source Web has more than one target Web override.");
            }
        }

        private static void ValidateSourceSite(SourceSiteCollectionSnapshot source, ICollection<MigrationIssue> issues)
        {
            Uri url;
            if (source == null || source.SiteId == Guid.Empty || source.RootWebId == Guid.Empty
                || !Uri.TryCreate(source.SiteCollectionUrl, UriKind.Absolute, out url) || url.Scheme != Uri.UriSchemeHttps
                || source.Availability == EvidenceAvailability.Unavailable || source.Availability == EvidenceAvailability.Conflict)
            {
                AddBlocker(issues, "InvalidSourceSite", "source-site:" + (source == null ? Guid.Empty : source.SiteId).ToString("D"), "A source Site Collection requires captured HTTPS URL, site/root identity, and non-conflicting evidence.");
            }
        }

        private static void ValidateSourceWeb(SourceSiteCollectionSnapshot source, SourceWebSnapshot web, ICollection<MigrationIssue> issues)
        {
            Uri webUrl;
            Uri siteUrl;
            var invalid = web == null || web.SiteId != source.SiteId || web.WebId == Guid.Empty
                || !Uri.TryCreate(web == null ? null : web.WebUrl, UriKind.Absolute, out webUrl)
                || !Uri.TryCreate(source.SiteCollectionUrl, UriKind.Absolute, out siteUrl)
                || webUrl.Scheme != Uri.UriSchemeHttps
                || !string.Equals(webUrl.Authority, siteUrl.Authority, StringComparison.OrdinalIgnoreCase)
                || web.Availability != EvidenceAvailability.Captured;
            if (invalid)
            {
                AddBlocker(issues, "InvalidSourceWeb", "source-web:" + (web == null ? Guid.Empty : web.WebId).ToString("D"), "The source Web must be a captured member of its declared Site Collection.");
                return;
            }

            var sitePath = NormalizeServerRelativePath(source.ServerRelativeUrl).TrimEnd('/');
            var webPath = NormalizeServerRelativePath(web.ServerRelativeUrl).TrimEnd('/');
            if (!string.Equals(webPath, sitePath, StringComparison.OrdinalIgnoreCase) && !webPath.StartsWith(sitePath + "/", StringComparison.OrdinalIgnoreCase))
            {
                AddBlocker(issues, "InvalidSourceWeb", "source-web:" + web.WebId.ToString("D"), "The source Web path is outside its Site Collection.");
            }
        }

        private static void ValidateTargetSite(TargetSiteCollectionSpec target, ICollection<MigrationIssue> issues)
        {
            Uri url;
            if (target == null || !Uri.TryCreate(target.TargetSiteUrl, UriKind.Absolute, out url) || url.Scheme != Uri.UriSchemeHttps
                || string.IsNullOrWhiteSpace(target.Title) || target.Language <= 0)
            {
                AddBlocker(issues, "InvalidTargetSiteSpec", target == null ? "target-site" : target.TargetSiteUrl, "The target Site Collection specification requires an HTTPS URL, title, and language.");
            }
            if (target != null && target.Mode == TargetSiteMode.ExistingTargetSite && (!target.ExpectedTargetSiteId.HasValue || target.ExpectedTargetSiteId.Value == Guid.Empty))
            {
                AddBlocker(issues, "InvalidTargetSiteSpec", target.TargetSiteUrl, "ExistingTargetSite mode requires the expected target Site Collection ID.");
            }
            if (target != null && target.Mode == TargetSiteMode.CreateTargetSite && (string.IsNullOrWhiteSpace(target.Owner) || string.IsNullOrWhiteSpace(target.Template)))
            {
                AddBlocker(issues, "InvalidTargetSiteSpec", target.TargetSiteUrl, "CreateTargetSite mode requires an owner and template.");
            }
        }

        private static string DefaultTargetRelativePath(
            SourceWebSnapshot web,
            SourceWebSnapshot parent)
        {
            var webPath = NormalizeServerRelativePath(web.ServerRelativeUrl);
            var parentPath = NormalizeServerRelativePath(parent.ServerRelativeUrl).TrimEnd('/');
            if (!webPath.StartsWith(parentPath + "/", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("The source Web path is outside its captured direct parent Web.", nameof(web));
            }

            return Uri.UnescapeDataString(webPath.Substring(parentPath.Length + 1));
        }

        private static bool IsSafeWebSegment(string value)
        {
            return !string.IsNullOrWhiteSpace(value) && value.Length <= 64
                && value.IndexOf('/') < 0 && value.IndexOf('\\') < 0 && value.IndexOf("..", StringComparison.Ordinal) < 0;
        }

        private static string NormalizeAbsoluteUrl(string value)
        {
            return new Uri(value).AbsoluteUri.TrimEnd('/');
        }

        private static string NormalizeServerRelativePath(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("A server-relative path is required.", nameof(value));
            }
            var normalized = Uri.UnescapeDataString(value.Trim()).Replace('\\', '/');
            if (!normalized.StartsWith("/", StringComparison.Ordinal))
            {
                throw new ArgumentException("A server-relative path is required.", nameof(value));
            }
            if (normalized.Split('/').Any(segment => segment == "." || segment == ".."))
            {
                throw new ArgumentException("Relative traversal is not allowed.", nameof(value));
            }
            return normalized.Length == 1 ? normalized : normalized.TrimEnd('/');
        }

        private static int PathDepth(string value)
        {
            return (value ?? string.Empty).Count(character => character == '/');
        }

        private static void AddBlocker(ICollection<MigrationIssue> issues, string code, string subject, string message)
        {
            issues.Add(new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = subject,
                Message = message
            });
        }
    }
}
