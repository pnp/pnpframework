using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace PnP.Framework.Migration.Lists.Capture
{
    internal static class ListDependencySnapshotReader
    {
        public static ListDependencySnapshot Read(
            ClientContext context,
            Web sourceWeb,
            Guid sourceListId,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            ICollection<string> warnings)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (sourceWeb == null)
            {
                throw new ArgumentNullException(nameof(sourceWeb));
            }
            if (sourceListId == Guid.Empty)
            {
                throw new ArgumentException("A source List ID is required.", nameof(sourceListId));
            }

            var site = context.Site;
            var rootWeb = site.RootWeb;
            var list = sourceWeb.Lists.GetById(sourceListId);
            context.Load(site, value => value.Id);
            context.Load(sourceWeb, value => value.Id, value => value.Url, value => value.ServerRelativeUrl);
            context.Load(rootWeb, value => value.Id, value => value.Url, value => value.ServerRelativeUrl);
            context.Load(list,
                value => value.Id,
                value => value.Title,
                value => value.Description,
                value => value.TemplateFeatureId,
                value => value.BaseTemplate,
                value => value.BaseType,
                value => value.Hidden,
                value => value.ContentTypesEnabled,
                value => value.EnableAttachments,
                value => value.EnableFolderCreation,
                value => value.EnableVersioning,
                value => value.EnableMinorVersions,
                value => value.EnableModeration,
                value => value.ForceCheckout,
                value => value.IrmEnabled,
                value => value.IrmExpire,
                value => value.IrmReject,
                value => value.ItemCount);
            context.Load(list.RootFolder,
                value => value.ServerRelativeUrl,
                value => value.UniqueContentTypeOrder);
            context.Load(list.Fields, values => values.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.Title,
                value => value.TypeAsString,
                value => value.Group,
                value => value.SchemaXml,
                value => value.Hidden,
                value => value.ReadOnlyField,
                value => value.Required,
                value => value.FromBaseType,
                value => value.Sealed));
            context.Load(list.ContentTypes, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.Description,
                value => value.Group,
                value => value.Hidden,
                value => value.ReadOnly,
                value => value.Sealed));
            context.Load(list.Views, values => values.Include(
                value => value.Id,
                value => value.Title,
                value => value.ServerRelativeUrl,
                value => value.Hidden,
                value => value.DefaultView,
                value => value.PersonalView,
                value => value.ViewType,
                value => value.RowLimit,
                value => value.Paged,
                value => value.ViewQuery,
                value => value.ListViewXml,
                value => value.JSLink));
            context.ExecuteQueryRetry();

            foreach (var contentType in list.ContentTypes)
            {
                context.Load(contentType.Parent, value => value.Id);
                context.Load(contentType.FieldLinks, values => values.Include(
                    value => value.Id,
                    value => value.Name,
                    value => value.DisplayName,
                    value => value.Required,
                    value => value.Hidden,
                    value => value.ReadOnly));
            }
            foreach (var view in list.Views)
            {
                context.Load(view.ViewFields);
            }
            context.ExecuteQueryRetry();

            var informationRightsManagement = ReadInformationRightsManagement(
                context,
                sourceWeb.Url,
                sourceListId,
                list.IrmEnabled,
                list.IrmExpire,
                list.IrmReject);
            var items = ListItemSnapshotReader.Read(context, list, maximumBytes, artifactStore, warnings);
            var views = ListViewSnapshotReader.Read(list.Views, list.RootFolder.ServerRelativeUrl);
            var viewRenderingResources = ListViewRenderingResourceSnapshotReader.Read(
                context,
                sourceWeb,
                rootWeb,
                views,
                maximumBytes,
                artifactStore,
                warnings);
            var contentTypeDiagnostics = new List<string>();
            var listContentTypes = ListContentTypeSnapshotReader.Read(list.ContentTypes, contentTypeDiagnostics);
            var siteContentTypes = ContentTypeClosureSnapshotReader.Read(context, sourceWeb, listContentTypes, contentTypeDiagnostics);
            var availability = EvidenceAvailability.Captured;
            if (items.Count != list.ItemCount || items.Any(value => value.Availability != EvidenceAvailability.Captured))
            {
                availability = EvidenceAvailability.Partial;
            }
            if (viewRenderingResources.Any(value => value.Availability != EvidenceAvailability.Captured))
            {
                availability = EvidenceAvailability.Partial;
            }
            if (contentTypeDiagnostics.Any(value => value.StartsWith("ConflictingListContentTypeFieldLink:", StringComparison.Ordinal)))
            {
                availability = EvidenceAvailability.Partial;
            }
            if (informationRightsManagement.Availability != EvidenceAvailability.Captured)
            {
                availability = EvidenceAvailability.Partial;
            }
            var snapshot = new ListDependencySnapshot
            {
                SourceSiteId = site.Id,
                SourceWebId = sourceWeb.Id,
                SourceWebUrl = sourceWeb.Url.TrimEnd('/'),
                SourceListId = list.Id,
                Title = list.Title,
                Description = list.Description ?? string.Empty,
                TemplateFeatureId = list.TemplateFeatureId,
                BaseTemplate = list.BaseTemplate,
                BaseType = list.BaseType.ToString(),
                RootFolderServerRelativeUrl = list.RootFolder.ServerRelativeUrl,
                Hidden = list.Hidden,
                ContentTypesEnabled = list.ContentTypesEnabled,
                EnableAttachments = list.EnableAttachments,
                EnableFolderCreation = list.EnableFolderCreation,
                EnableVersioning = list.EnableVersioning,
                EnableMinorVersions = list.EnableMinorVersions,
                EnableModeration = list.EnableModeration,
                ForceCheckout = list.ForceCheckout,
                InformationRightsManagement = informationRightsManagement,
                SourceItemCount = list.ItemCount,
                Fields = ListFieldSnapshotReader.Read(context, list.Fields),
                ContentTypes = listContentTypes,
                HasExplicitUniqueContentTypeOrder = list.RootFolder.UniqueContentTypeOrder != null,
                UniqueContentTypeOrder = (list.RootFolder.UniqueContentTypeOrder ?? new ContentTypeId[0])
                    .Select(value => value.StringValue).ToList(),
                SiteContentTypes = siteContentTypes,
                Views = views,
                ViewRenderingResources = viewRenderingResources,
                Items = items,
                Availability = availability
            };
            if (items.Count != list.ItemCount)
            {
                snapshot.Diagnostics.Add("Captured item count " + items.Count + " differs from source List.ItemCount " + list.ItemCount + ".");
                warnings.Add("List '" + list.Title + "' item count changed or was not captured completely.");
            }
            foreach (var diagnostic in contentTypeDiagnostics)
            {
                snapshot.Diagnostics.Add(diagnostic);
                warnings.Add(diagnostic);
            }
            foreach (var diagnostic in informationRightsManagement.Diagnostics)
            {
                snapshot.Diagnostics.Add(diagnostic);
                warnings.Add(diagnostic);
            }
            return snapshot;
        }

        private static ListInformationRightsManagementSnapshot ReadInformationRightsManagement(
            ClientContext context,
            string sourceWebUrl,
            Guid sourceListId,
            bool irmEnabled,
            bool irmExpire,
            bool irmReject)
        {
            var snapshot = new ListInformationRightsManagementSnapshot
            {
                IrmEnabled = irmEnabled,
                IrmExpire = irmExpire,
                IrmReject = irmReject
            };
            if (!irmEnabled)
            {
                return snapshot;
            }

            try
            {
                using (var policyContext = context.Clone(sourceWebUrl))
                {
                    var sourceList = policyContext.Web.Lists.GetById(sourceListId);
                    var settings = sourceList.InformationRightsManagementSettings;
                    policyContext.Load(settings,
                        value => value.AllowPrint,
                        value => value.AllowScript,
                        value => value.AllowWriteCopy,
                        value => value.DisableDocumentBrowserView,
                        value => value.DocumentAccessExpireDays,
                        value => value.DocumentLibraryProtectionExpireDate,
                        value => value.EnableDocumentAccessExpire,
                        value => value.EnableDocumentBrowserPublishingView,
                        value => value.EnableGroupProtection,
                        value => value.EnableLicenseCacheExpire,
                        value => value.GroupName,
                        value => value.LicenseCacheExpireDays,
                        value => value.PolicyDescription,
                        value => value.PolicyTitle,
                        value => value.TemplateId);
                    policyContext.ExecuteQueryRetry();
                    snapshot.Policy = new ListInformationRightsManagementPolicySnapshot
                    {
                        AllowPrint = settings.AllowPrint,
                        AllowScript = settings.AllowScript,
                        AllowWriteCopy = settings.AllowWriteCopy,
                        DisableDocumentBrowserView = settings.DisableDocumentBrowserView,
                        DocumentAccessExpireDays = settings.DocumentAccessExpireDays,
                        DocumentLibraryProtectionExpireDate = settings.DocumentLibraryProtectionExpireDate,
                        EnableDocumentAccessExpire = settings.EnableDocumentAccessExpire,
                        EnableDocumentBrowserPublishingView = settings.EnableDocumentBrowserPublishingView,
                        EnableGroupProtection = settings.EnableGroupProtection,
                        EnableLicenseCacheExpire = settings.EnableLicenseCacheExpire,
                        GroupName = settings.GroupName,
                        LicenseCacheExpireDays = settings.LicenseCacheExpireDays,
                        PolicyDescription = settings.PolicyDescription,
                        PolicyTitle = settings.PolicyTitle,
                        TemplateId = settings.TemplateId
                    };
                }
            }
            catch (WebException exception)
            {
                using (var response = exception.Response as HttpWebResponse)
                {
                    var statusCode = response == null ? 0 : (int)response.StatusCode;
                    if (statusCode == 401 || statusCode == 403)
                    {
                        snapshot.Availability = EvidenceAvailability.Unavailable;
                        snapshot.AuthorizationEvidence = LiteralHttpAuthorizationEvidence.Create(
                            "capture-list-irm-policy",
                            response.ResponseUri?.AbsoluteUri ?? sourceWebUrl,
                            statusCode,
                            DateTimeOffset.UtcNow);
                        snapshot.Diagnostics.Add(
                            "List IRM policy request returned literal HTTP " + statusCode + ".");
                        return snapshot;
                    }
                }
                snapshot.Availability = EvidenceAvailability.Partial;
                snapshot.Diagnostics.Add(
                    "List IRM policy capture failed without literal HTTP 401/403 evidence: " + exception.Message);
            }
            catch (Exception exception)
            {
                snapshot.Availability = EvidenceAvailability.Partial;
                snapshot.Diagnostics.Add(
                    "List IRM policy capture failed without literal HTTP 401/403 evidence: " + exception.Message);
            }
            return snapshot;
        }
    }
}
