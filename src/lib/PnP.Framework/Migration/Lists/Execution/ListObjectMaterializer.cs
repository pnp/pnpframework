using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal sealed class ListObjectMaterializationResult
    {
        public List List { get; set; }

        public ListMaterializationDisposition Disposition { get; set; }
    }

    internal static class ListObjectMaterializer
    {
        public static ListObjectMaterializationResult Ensure(ClientContext context, ListDependencySnapshot source, ListMaterializationPlan plan)
        {
            var fresh = ListTargetInspector.Inspect(context, source, plan);
            if (!fresh.IsAdmitted)
            {
                throw new InvalidOperationException("Fresh target List preflight no longer permits the approved operation: "
                    + string.Join("; ", fresh.Issues.Select(value => value.Message)));
            }
            if (fresh.Disposition == ListMaterializationDisposition.ReuseOwned)
            {
                var existing = context.Web.Lists.GetById(fresh.TargetListId.Value);
                Load(context, existing);
                return new ListObjectMaterializationResult
                {
                    List = existing,
                    Disposition = fresh.Disposition
                };
            }
            if (fresh.Disposition != ListMaterializationDisposition.CreateOwned)
            {
                throw new InvalidOperationException("Unexpected List materialization disposition: " + fresh.Disposition + ".");
            }

            context.Load(context.Web, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var relativeUrl = plan.TargetRootFolderServerRelativeUrl.Substring(context.Web.ServerRelativeUrl.TrimEnd('/').Length).TrimStart('/');
            if (string.IsNullOrWhiteSpace(relativeUrl))
            {
                throw new InvalidDataException("A target List URL below the mapped Web is required.");
            }
            var creation = new ListCreationInformation
            {
                Title = plan.TargetTitle,
                Description = source.Description,
                Url = relativeUrl,
                TemplateType = source.BaseTemplate
            };
            if (source.TemplateFeatureId != Guid.Empty)
            {
                creation.TemplateFeatureId = source.TemplateFeatureId;
            }
            var list = context.Web.Lists.Add(creation);
            list.Hidden = source.Hidden;
            list.ContentTypesEnabled = source.ContentTypesEnabled;
            list.EnableAttachments = source.EnableAttachments;
            list.EnableFolderCreation = source.EnableFolderCreation;
            list.EnableVersioning = source.EnableVersioning;
            list.EnableMinorVersions = source.EnableMinorVersions;
            list.EnableModeration = source.EnableModeration;
            list.ForceCheckout = source.ForceCheckout;
            list.Update();
            context.ExecuteQueryRetry();
            Load(context, list);
            list.RootFolder.Properties[ListTargetInspector.OriginalIdentifierPropertyName] = plan.OriginalIdentifier;
            list.RootFolder.Properties[ListTargetInspector.PlanDigestPropertyName] = plan.PlanDigest;
            list.RootFolder.Update();
            context.ExecuteQueryRetry();
            Load(context, list);
            if (!string.Equals(Property(list.RootFolder.Properties, ListTargetInspector.OriginalIdentifierPropertyName), plan.OriginalIdentifier, StringComparison.Ordinal)
                || !string.Equals(Property(list.RootFolder.Properties, ListTargetInspector.PlanDigestPropertyName), plan.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Fresh target List provenance readback differs from the sealed plan.");
            }
            return new ListObjectMaterializationResult
            {
                List = list,
                Disposition = fresh.Disposition
            };
        }

        private static void Load(ClientContext context, List list)
        {
            context.Load(list, value => value.Id, value => value.Title, value => value.BaseTemplate, value => value.RootFolder.ServerRelativeUrl, value => value.RootFolder.Properties);
            context.ExecuteQueryRetry();
        }

        private static string Property(PropertyValues values, string name)
        {
            object value;
            return values.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value) : null;
        }
    }
}
