using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListViewMaterializer
    {
        public static IDictionary<Guid, Guid> Ensure(ClientContext context, List list, ListMaterializationPlan plan)
        {
            context.Load(list.Views, values => values.Include(
                value => value.Id,
                value => value.Title,
                value => value.PersonalView,
                value => value.ViewQuery,
                value => value.RowLimit,
                value => value.Paged,
                value => value.JSLink));
            context.ExecuteQueryRetry();
            foreach (var view in list.Views)
            {
                context.Load(view.ViewFields);
            }
            context.ExecuteQueryRetry();
            var publicViewsByTitle = list.Views
                .AsEnumerable()
                .Where(value => !value.PersonalView)
                .GroupBy(value => value.Title, StringComparer.Ordinal)
                .ToDictionary(
                    value => value.Key,
                    value => value.ToList(),
                    StringComparer.Ordinal);
            var result = new Dictionary<Guid, Guid>();
            foreach (var viewPlan in plan.Views.Where(value => value.Disposition == ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView
                || value.Disposition == ListViewMaterializationDisposition.CreateOrReuseWebPartView)
                .OrderByDescending(value => value.Source.DefaultView)
                .ThenBy(value => value.Source.Title, StringComparer.OrdinalIgnoreCase)
                .ThenBy(value => value.SourceViewId))
            {
                var source = viewPlan.Source;
                var title = TargetTitle(viewPlan);
                publicViewsByTitle.TryGetValue(title, out var candidates);
                if (candidates != null && candidates.Count > 1)
                {
                    throw new InvalidOperationException("Target List contains multiple public Views named '" + title + "'.");
                }
                var target = candidates?.SingleOrDefault();
                if (target == null)
                {
                    target = list.Views.Add(new ViewCreationInformation
                    {
                        Title = title,
                        Query = source.ViewQuery,
                        RowLimit = source.RowLimit,
                        Paged = source.Paged,
                        PersonalView = false,
                        SetAsDefaultView = source.DefaultView,
                        ViewFields = source.ViewFields.ToArray(),
                        ViewTypeKind = ParseViewType(source.ViewType)
                    });
                    context.ExecuteQueryRetry();
                    publicViewsByTitle[title] = new List<View> { target };
                }
                var expectedJsLink = ListViewRenderingResourceMaterializer.RewriteJsLink(
                    source.JsLink,
                    source,
                    plan.ViewRenderingResources);
                target.ViewQuery = source.ViewQuery;
                target.RowLimit = source.RowLimit;
                target.Paged = source.Paged;
                target.JSLink = expectedJsLink;
                target.ViewFields.RemoveAll();
                foreach (var field in source.ViewFields)
                {
                    target.ViewFields.Add(field);
                }
                target.Update();
                context.ExecuteQueryRetry();
                context.Load(target, value => value.Id, value => value.Title, value => value.ViewType, value => value.ViewQuery, value => value.RowLimit, value => value.Paged, value => value.JSLink);
                context.Load(target.ViewFields);
                context.ExecuteQueryRetry();
                if (!string.Equals(target.Title, title, StringComparison.Ordinal)
                    || !string.Equals(target.ViewType, source.ViewType, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(target.ViewQuery ?? string.Empty, source.ViewQuery ?? string.Empty, StringComparison.Ordinal)
                    || target.RowLimit != source.RowLimit
                    || target.Paged != source.Paged
                    || !string.Equals(target.JSLink ?? string.Empty, expectedJsLink ?? string.Empty, StringComparison.Ordinal)
                    || !target.ViewFields.SequenceEqual(source.ViewFields, StringComparer.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("Fresh target View readback differs from the sealed source View: " + source.Id.ToString("D") + ".");
                }
                result[source.Id] = target.Id;
            }
            return result;
        }

        internal static string TargetTitle(ListViewMaterializationPlan plan)
        {
            return plan.Disposition == ListViewMaterializationDisposition.CreateOrReuseWebPartView
                ? "PnP migration " + plan.Source.Id.ToString("N")
                : plan.Source.Title;
        }

        private static ViewType ParseViewType(string value)
        {
            ViewType result;
            return Enum.TryParse(value, true, out result) ? result : ViewType.Html;
        }
    }
}
