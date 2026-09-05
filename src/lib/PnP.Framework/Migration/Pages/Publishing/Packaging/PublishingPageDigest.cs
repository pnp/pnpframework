using PnP.Framework.Migration.Pages.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public static class PublishingPageDigest
    {
        public static string ComputeSelectionDigest(PublishingPageWorkflowSelection selection)
        {
            if (selection == null)
            {
                throw new ArgumentNullException(nameof(selection));
            }

            return PageDigest.ComputeSha256(PublishingPagePackageSerializer.SerializeCanonical(selection));
        }

        public static string ComputeSnapshotDigest(PublishingPageCaptureBundle snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            return PageDigest.ComputeSha256(PublishingPagePackageSerializer.SerializeCanonical(snapshot));
        }

        internal static string ComputeLegacySnapshotDigestWithoutViewRenderingResources(PublishingPageCaptureBundle snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            var dependencies = (snapshot.ListDependencies ?? Array.Empty<Lists.Capture.ListDependencySnapshot>())
                .Where(value => value != null)
                .ToArray();
            var resourceInventories = dependencies
                .Select(value => value.ViewRenderingResources)
                .ToArray();
            var views = dependencies
                .SelectMany(value => value.Views ?? Array.Empty<Lists.Views.ListViewSnapshot>())
                .Where(value => value != null)
                .ToArray();
            var bindingInventories = views
                .Select(value => value.RenderingResourceBindings)
                .ToArray();
            try
            {
                foreach (var dependency in dependencies)
                {
                    dependency.ViewRenderingResources = null;
                }
                foreach (var view in views)
                {
                    view.RenderingResourceBindings = null;
                }
                return ComputeSnapshotDigest(snapshot);
            }
            finally
            {
                for (var index = 0; index < dependencies.Length; index++)
                {
                    dependencies[index].ViewRenderingResources = resourceInventories[index];
                }
                for (var index = 0; index < views.Length; index++)
                {
                    views[index].RenderingResourceBindings = bindingInventories[index];
                }
            }
        }

        public static string ComputePlanDigest(PublishingPageMigrationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return PageDigest.ComputeSha256(PublishingPagePackageSerializer.SerializeCanonical(plan));
        }

        public static string ComputeSha256(string value)
        {
            return PageDigest.ComputeSha256(value);
        }

        public static string ComputeSha256(byte[] value)
        {
            return PageDigest.ComputeSha256(value);
        }
    }
}
