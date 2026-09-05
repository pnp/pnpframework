using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    internal sealed class ContentTypeTargetAdmissionContext
    {
        private readonly HashSet<string> plannedContentTypeIds =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, HashSet<Guid>> plannedFieldLinks =
            new Dictionary<string, HashSet<Guid>>(StringComparer.OrdinalIgnoreCase);

        private readonly HashSet<Guid> provisionedRuntimeFieldIds;

        public ContentTypeTargetAdmissionContext(IEnumerable<Guid> provisionedRuntimeFieldIds)
        {
            this.provisionedRuntimeFieldIds = new HashSet<Guid>(
                provisionedRuntimeFieldIds ?? Enumerable.Empty<Guid>());
        }

        public bool WillProvideContentType(string contentTypeId)
        {
            return !string.IsNullOrWhiteSpace(contentTypeId)
                && plannedContentTypeIds.Contains(contentTypeId);
        }

        public bool WillProvideParentFieldLink(string parentContentTypeId, Guid fieldId)
        {
            return !string.IsNullOrWhiteSpace(parentContentTypeId)
                && plannedFieldLinks.TryGetValue(parentContentTypeId, out var fields)
                && fields.Contains(fieldId);
        }

        public bool WillProvisionRuntimeField(Guid fieldId)
        {
            return provisionedRuntimeFieldIds.Contains(fieldId);
        }

        public void RegisterAdmitted(ContentTypeMaterializationPlan plan)
        {
            if (plan == null || string.IsNullOrWhiteSpace(plan.ContentTypeId))
            {
                throw new ArgumentException("An admitted content type plan with an ID is required.", nameof(plan));
            }

            plannedContentTypeIds.Add(plan.ContentTypeId);
            plannedFieldLinks[plan.ContentTypeId] = new HashSet<Guid>(
                (plan.RequiredFieldLinks ?? Enumerable.Empty<ContentTypeFieldLinkSnapshot>())
                    .Where(value => value != null)
                    .Select(value => value.FieldId));
        }
    }
}
