using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PnP.Framework.Migration.Pages.Publishing.Profiles;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public sealed class EnterpriseWikiPageDiscovery
    {
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

            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            ListItemCollectionPosition position = null;
            do
            {
                var query = new CamlQuery
                {
                    ListItemCollectionPosition = position,
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
  <RowLimit Paged='TRUE'>500</RowLimit>
</View>"
                };
                var items = pages.GetItems(query);
                sourceContext.Load(items, collection => collection.ListItemCollectionPosition);
                sourceContext.ExecuteQueryRetry();
                foreach (var item in items)
                {
                    var contentTypeId = Convert.ToString(item["ContentTypeId"], CultureInfo.InvariantCulture);
                    var fileRef = Convert.ToString(item["FileRef"], CultureInfo.InvariantCulture);
                    if (EnterpriseWikiV1CohortPolicy.IsIncludedContentType(contentTypeId) && !string.IsNullOrWhiteSpace(fileRef))
                    {
                        result.Add(fileRef);
                    }
                }

                position = items.ListItemCollectionPosition;
            }
            while (position != null);

            return result.OrderBy(item => item, StringComparer.OrdinalIgnoreCase).ToArray();
        }

        public static bool IsEnterpriseWikiContentType(string contentTypeId)
        {
            return EnterpriseWikiV1CohortPolicy.IsIncludedContentType(contentTypeId);
        }
    }
}
