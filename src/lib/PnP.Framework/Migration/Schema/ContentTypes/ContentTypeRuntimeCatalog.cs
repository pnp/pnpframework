using System;
using System.Collections.Generic;
using System.Linq;
using PnP.Framework.Migration.Features;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    internal static class ContentTypeRuntimeCatalog
    {
        internal static readonly Guid AssetLibraryFeatureId = new Guid("4bcccd62-dcaf-46dc-a7d4-e38277ef33f4");

        internal static readonly Guid DocumentSetFeatureId = new Guid("3bae86a2-776d-499d-9db8-fa4cdc7884f8");

        internal static readonly Guid VideoAndRichMediaFeatureId = new Guid("6e1e5426-2ebd-4871-8027-c5ca86371ead");

        internal static readonly Guid DocumentIdServiceFeatureId = new Guid("b50e3104-6812-424f-a011-cc90e6327318");

        private static readonly HashSet<Guid> DocumentIdFieldIds = new HashSet<Guid>
        {
            new Guid("ae3e2a36-125d-45d3-9051-744b513536a6"),
            new Guid("3b63724f-3418-461f-868b-7706f69b029c"),
            new Guid("c010d384-479c-494f-968c-c413dbe3de29")
        };

        private const string SystemMediaCollectionContentTypeId = "0x0120D520A8";

        private const string VideoContentTypeId = "0x0120D520A808";

        private const string RichMediaAssetContentTypeId = "0x0101009148F5A04DDD49CBA7127AADA5FB792B";

        private const string VideoRenditionContentTypeId = RichMediaAssetContentTypeId + "00291D173ECE694D56B19D111489C4369D";

        private const string AudioContentTypeId = RichMediaAssetContentTypeId + "006973ACD696DC4858A76371B2FB2F439A";

        private const string ImageContentTypeId = RichMediaAssetContentTypeId + "00AADE34325A8B49CDA8BB4DB53328F214";

        private static readonly HashSet<string> KnownTargetRuntimeIds = new HashSet<string>(new[]
        {
            BuiltInContentTypeId.Item,
            BuiltInContentTypeId.Document,
            BuiltInContentTypeId.Folder,
            BuiltInContentTypeId.DocumentSet,
            BuiltInContentTypeId.Link,
            BuiltInContentTypeId.LinkToDocument,
            // SharePoint runtime children used by Document Set and Asset Library templates.
            // Keep this list exact: a prefix match would incorrectly classify custom
            // document content types as target-runtime schema.
            SystemMediaCollectionContentTypeId,
            VideoContentTypeId,
            RichMediaAssetContentTypeId,
            VideoRenditionContentTypeId,
            AudioContentTypeId,
            ImageContentTypeId
        }, StringComparer.OrdinalIgnoreCase);

        public static bool IsTargetRuntime(string contentTypeId)
        {
            return !string.IsNullOrWhiteSpace(contentTypeId) && KnownTargetRuntimeIds.Contains(contentTypeId);
        }

        public static IList<PlatformFeatureMaterializationPlan> CreateFeatureRequirements(
            IEnumerable<string> contentTypeIds,
            string targetWebUrl)
        {
            return CreateFeatureRequirements(
                contentTypeIds,
                Enumerable.Empty<ContentTypeSchemaSnapshot>(),
                targetWebUrl);
        }

        public static IList<PlatformFeatureMaterializationPlan> CreateFeatureRequirements(
            IEnumerable<string> contentTypeIds,
            IEnumerable<ContentTypeSchemaSnapshot> contentTypeSchemas,
            string targetWebUrl)
        {
            var requirements = new Dictionary<Guid, PlatformFeatureMaterializationPlan>();
            foreach (var contentTypeId in (contentTypeIds ?? Enumerable.Empty<string>())
                         .Where(value => !string.IsNullOrWhiteSpace(value))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (IsAssetContentType(contentTypeId))
                {
                    Add(requirements, AssetLibrary(), contentTypeId, contentTypeId);
                }
                if (string.Equals(contentTypeId, BuiltInContentTypeId.DocumentSet, StringComparison.OrdinalIgnoreCase))
                {
                    Add(requirements, DocumentSets(), contentTypeId, BuiltInContentTypeId.DocumentSet);
                }
                if (string.Equals(contentTypeId, SystemMediaCollectionContentTypeId, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(contentTypeId, VideoContentTypeId, StringComparison.OrdinalIgnoreCase))
                {
                    Add(requirements, AssetLibrary(), contentTypeId, null);
                    Add(requirements, DocumentSets(), contentTypeId, BuiltInContentTypeId.DocumentSet);
                    Add(requirements, VideoAndRichMedia(), contentTypeId, SystemMediaCollectionContentTypeId);
                    Add(requirements, VideoAndRichMedia(), contentTypeId, contentTypeId);
                }
            }

            foreach (var schema in (contentTypeSchemas ?? Enumerable.Empty<ContentTypeSchemaSnapshot>())
                         .Where(value => value != null && !string.IsNullOrWhiteSpace(value.ContentTypeId)))
            {
                if ((schema.RequiredFieldClosure ?? new List<Schema.Fields.FieldSchemaSnapshot>())
                    .Any(value => value != null && DocumentIdFieldIds.Contains(value.Id)))
                {
                    Add(requirements, DocumentIdService(), schema.ContentTypeId, BuiltInContentTypeId.LinkToDocument);
                }
            }

            foreach (var requirement in requirements.Values)
            {
                requirement.TargetWebUrl = targetWebUrl;
                requirement.RequiredByContentTypeIds = requirement.RequiredByContentTypeIds
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList();
                requirement.ExpectedContentTypeIds = requirement.ExpectedContentTypeIds
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList();
            }
            return requirements.Values.OrderBy(value => value.DependencyOrder).ThenBy(value => value.FeatureId).ToList();
        }

        internal static IEnumerable<Guid> RuntimeFieldIdsProvidedBy(
            IEnumerable<PlatformFeatureMaterializationPlan> requirements)
        {
            return (requirements ?? Enumerable.Empty<PlatformFeatureMaterializationPlan>())
                .Where(value => value != null && value.FeatureId == DocumentIdServiceFeatureId)
                .SelectMany(_ => DocumentIdFieldIds)
                .Distinct();
        }

        internal static bool IsDocumentIdField(Guid fieldId)
        {
            return DocumentIdFieldIds.Contains(fieldId);
        }

        private static bool IsAssetContentType(string contentTypeId)
        {
            return string.Equals(contentTypeId, RichMediaAssetContentTypeId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentTypeId, VideoRenditionContentTypeId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentTypeId, AudioContentTypeId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentTypeId, ImageContentTypeId, StringComparison.OrdinalIgnoreCase);
        }

        private static void Add(
            IDictionary<Guid, PlatformFeatureMaterializationPlan> requirements,
            PlatformFeatureMaterializationPlan definition,
            string requiredByContentTypeId,
            string expectedContentTypeId)
        {
            PlatformFeatureMaterializationPlan requirement;
            if (!requirements.TryGetValue(definition.FeatureId, out requirement))
            {
                requirement = definition;
                requirements[definition.FeatureId] = requirement;
            }
            requirement.RequiredByContentTypeIds.Add(requiredByContentTypeId);
            if (!string.IsNullOrWhiteSpace(expectedContentTypeId))
            {
                requirement.ExpectedContentTypeIds.Add(expectedContentTypeId);
            }
        }

        private static PlatformFeatureMaterializationPlan AssetLibrary()
        {
            return Feature(
                AssetLibraryFeatureId,
                "Asset Library",
                100,
                new Guid[0],
                "Ensure the SharePoint-owned Asset Library site feature that provides rich-media runtime content types.");
        }

        private static PlatformFeatureMaterializationPlan DocumentSets()
        {
            return Feature(
                DocumentSetFeatureId,
                "Document Sets",
                200,
                new Guid[0],
                "Ensure the SharePoint-owned Document Sets site feature that provides the Document Set runtime content type.");
        }

        private static PlatformFeatureMaterializationPlan VideoAndRichMedia()
        {
            return Feature(
                VideoAndRichMediaFeatureId,
                "Video and Rich Media",
                300,
                new[] { AssetLibraryFeatureId, DocumentSetFeatureId },
                "Ensure the SharePoint-owned Video and Rich Media site feature that provides System Media Collection and Video runtime content types.");
        }

        private static PlatformFeatureMaterializationPlan DocumentIdService()
        {
            return Feature(
                DocumentIdServiceFeatureId,
                "Document ID Service",
                50,
                new Guid[0],
                "Ensure the SharePoint-owned Document ID Service site feature that provides the sealed Document ID fields and inherited field links.");
        }

        private static PlatformFeatureMaterializationPlan Feature(
            Guid id,
            string name,
            int dependencyOrder,
            IEnumerable<Guid> dependencies,
            string reason)
        {
            return new PlatformFeatureMaterializationPlan
            {
                FeatureId = id,
                Name = name,
                Scope = PlatformFeatureScope.SiteCollection,
                DependencyOrder = dependencyOrder,
                DependsOnFeatureIds = dependencies.ToList(),
                Disposition = PlatformFeatureMaterializationDisposition.EnsureActive,
                Reason = reason
            };
        }
    }
}
