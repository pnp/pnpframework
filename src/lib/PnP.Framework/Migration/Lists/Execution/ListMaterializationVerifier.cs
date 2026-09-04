using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PnP.Framework.Migration.Features;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListMaterializationVerifier
    {
        public static void Verify(
            ClientContext context,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationReceipt receipt,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts)
        {
            Verify(context, source, plan, receipt, dependencyReceipts, null);
        }

        public static void Verify(
            ClientContext context,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationReceipt receipt,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            ListMaterializationExecutionScope.ListSelection selection)
        {
            var diagnostics = new List<string>();
            try
            {
                VerifyCore(context, source, plan, receipt, dependencyReceipts, selection, diagnostics);
            }
            catch (Exception exception)
            {
                diagnostics.Add("Fresh List readback failed: " + exception.Message);
            }
            receipt.Diagnostics = diagnostics;
            receipt.FreshReadbackPassed = diagnostics.Count == 0;
        }

        private static void VerifyCore(
            ClientContext context,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationReceipt receipt,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            ListMaterializationExecutionScope.ListSelection selection,
            ICollection<string> diagnostics)
        {
            var list = context.Web.Lists.GetById(receipt.TargetListId);
            context.Load(context.Web, value => value.Id, value => value.Url);
            context.Load(list,
                value => value.Id,
                value => value.Title,
                value => value.Description,
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
                value => value.ItemCount);
            context.Load(list.RootFolder,
                value => value.ServerRelativeUrl,
                value => value.Properties,
                value => value.UniqueContentTypeOrder);
            context.ExecuteQueryRetry();

            VerifyIdentityAndSettings(context, list, source, plan, receipt, diagnostics);
            VerifyRequiredFeatures(context, plan, diagnostics);
            ListSchemaVerifier.Verify(context, list, source, plan, receipt, dependencyReceipts, selection, diagnostics);
            ListItemVerifier.Verify(context, list, source, plan, receipt, dependencyReceipts, selection, diagnostics);
        }

        private static void VerifyRequiredFeatures(
            ClientContext context,
            ListMaterializationPlan plan,
            ICollection<string> diagnostics)
        {
            var probes = PlatformFeatureTargetInspector.Inspect(context, plan.RequiredFeatures);
            foreach (var feature in plan.RequiredFeatures)
            {
                PlatformFeatureTargetProbe probe;
                if (!probes.TryGetValue(feature.FeatureId, out probe) || !probe.IsActive || !probe.IsAdmitted)
                {
                    diagnostics.Add("Target platform feature readback failed: " + feature.Name + " ("
                        + feature.FeatureId.ToString("D") + ")."
                        + (probe == null ? string.Empty : " " + string.Join("; ", probe.Issues.Select(value => value.Message))));
                }
            }
        }

        private static void VerifyIdentityAndSettings(
            ClientContext context,
            List list,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationReceipt receipt,
            ICollection<string> diagnostics)
        {
            Equal(context.Web.Id, receipt.TargetWebId, "Target Web ID", diagnostics);
            Equal(list.Id, receipt.TargetListId, "Target List ID", diagnostics);
            Equal(list.Title, plan.TargetTitle, "List title", diagnostics);
            Equal(list.Description ?? string.Empty, source.Description ?? string.Empty, "List description", diagnostics);
            Equal(list.BaseTemplate, source.BaseTemplate, "List base template", diagnostics);
            Equal(list.BaseType.ToString(), source.BaseType, "List base type", diagnostics);
            Equal(NormalizePath(list.RootFolder.ServerRelativeUrl), NormalizePath(plan.TargetRootFolderServerRelativeUrl), "List root path", diagnostics);
            Equal(list.Hidden, source.Hidden, "List Hidden", diagnostics);
            Equal(list.ContentTypesEnabled, source.ContentTypesEnabled, "List ContentTypesEnabled", diagnostics);
            Equal(list.EnableAttachments, source.EnableAttachments, "List EnableAttachments", diagnostics);
            Equal(list.EnableFolderCreation, source.EnableFolderCreation, "List EnableFolderCreation", diagnostics);
            Equal(list.EnableVersioning, source.EnableVersioning, "List EnableVersioning", diagnostics);
            Equal(list.EnableMinorVersions, source.EnableMinorVersions, "List EnableMinorVersions", diagnostics);
            Equal(list.EnableModeration, source.EnableModeration, "List EnableModeration", diagnostics);
            Equal(list.ForceCheckout, source.ForceCheckout, "List ForceCheckout", diagnostics);
            Equal(Property(list.RootFolder.Properties, ListTargetInspector.OriginalIdentifierPropertyName), plan.OriginalIdentifier, "List provenance identifier", diagnostics);
            Equal(Property(list.RootFolder.Properties, ListTargetInspector.PlanDigestPropertyName), plan.PlanDigest, "List provenance plan digest", diagnostics, true);
            Equal(receipt.PlanDigest, plan.PlanDigest, "Receipt plan digest", diagnostics, true);
        }

        private static string NormalizePath(string value)
        {
            return Uri.UnescapeDataString(value ?? string.Empty).TrimEnd('/');
        }

        private static string Property(PropertyValues values, string name)
        {
            object value;
            return values != null && values.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value, CultureInfo.InvariantCulture) : null;
        }

        private static void Equal<T>(T actual, T expected, string subject, ICollection<string> diagnostics)
        {
            if (!EqualityComparer<T>.Default.Equals(actual, expected))
            {
                diagnostics.Add(subject + " differs. Expected '" + expected + "', observed '" + actual + "'.");
            }
        }

        private static void Equal(string actual, string expected, string subject, ICollection<string> diagnostics, bool ignoreCase = false)
        {
            if (!string.Equals(actual ?? string.Empty, expected ?? string.Empty, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            {
                diagnostics.Add(subject + " differs. Expected '" + expected + "', observed '" + actual + "'.");
            }
        }
    }
}
