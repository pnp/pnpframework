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
            foreach (var view in list.Views)
            {
                context.Load(view.ViewFields);
            }
            context.ExecuteQueryRetry();
            var result = new Dictionary<Guid, Guid>();
            foreach (var viewPlan in plan.Views.Where(value => value.Disposition == ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView
                || value.Disposition == ListViewMaterializationDisposition.CreateOrReuseWebPartView))
            {
                var source = viewPlan.Source;
                var title = viewPlan.Disposition == ListViewMaterializationDisposition.CreateOrReuseWebPartView
                    ? "PnP migration " + source.Id.ToString("N")
                    : source.Title;
                var target = list.Views.AsEnumerable().FirstOrDefault(value => !value.PersonalView && string.Equals(value.Title, title, StringComparison.Ordinal));
                if (target == null)
                {
                    target = list.Views.Add(new ViewCreationInformation
                    {
                        Title = title,
                        Query = source.ViewQuery,
                        RowLimit = source.RowLimit,
                        Paged = source.Paged,
                        PersonalView = false,
                        SetAsDefaultView = false,
                        ViewFields = source.ViewFields.ToArray(),
                        ViewTypeKind = ParseViewType(source.ViewType)
                    });
                    context.ExecuteQueryRetry();
                }
                target.ViewQuery = source.ViewQuery;
                target.RowLimit = source.RowLimit;
                target.Paged = source.Paged;
                target.JSLink = source.JsLink;
                target.ViewFields.RemoveAll();
                foreach (var field in source.ViewFields)
                {
                    target.ViewFields.Add(field);
                }
                target.Update();
                context.ExecuteQueryRetry();
                context.Load(target, value => value.Id, value => value.ViewQuery, value => value.RowLimit, value => value.Paged, value => value.JSLink);
                context.Load(target.ViewFields);
                context.ExecuteQueryRetry();
                if (!string.Equals(target.ViewQuery ?? string.Empty, source.ViewQuery ?? string.Empty, StringComparison.Ordinal)
                    || target.RowLimit != source.RowLimit
                    || target.Paged != source.Paged
                    || !string.Equals(target.JSLink ?? string.Empty, source.JsLink ?? string.Empty, StringComparison.Ordinal)
                    || !target.ViewFields.SequenceEqual(source.ViewFields, StringComparer.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("Fresh target View readback differs from the sealed source View: " + source.Id.ToString("D") + ".");
                }
                result[source.Id] = target.Id;
            }
            return result;
        }

        private static ViewType ParseViewType(string value)
        {
            ViewType result;
            return Enum.TryParse(value, true, out result) ? result : ViewType.Html;
        }
    }
}
