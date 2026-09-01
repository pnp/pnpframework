using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Packaging;
using System;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutMaterializer
    {
        private const string SystemPageLayoutContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC1";

        public static bool Ensure(
            ClientContext context,
            PublishingPageLayoutSnapshot snapshot,
            PublishingPageLayoutMaterializationPlan plan,
            PublishingPageLayoutTargetAdmission admission,
            IMigrationArtifactStore artifactStore)
        {
            if (admission == null || !admission.IsEligible)
            {
                throw new InvalidOperationException("A blocked Page Layout cannot be materialized.");
            }

            if (admission.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                || admission.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseOwned)
            {
                Verify(context, plan);
                return false;
            }

            if (snapshot?.Bytes == null || plan.TargetBytes == null)
            {
                throw new InvalidOperationException("Custom Page Layout source and target byte evidence is incomplete.");
            }

            var sourceBytes = MigrationArtifact.ReadAllBytes(snapshot.Bytes, snapshot.ContentBase64, artifactStore);
            var targetBytes = PublishingPageLayoutResourceRewriter.Rewrite(sourceBytes, plan.ResourceRewrites);
            if (targetBytes.LongLength != plan.TargetBytes.Length
                || !string.Equals(MigrationDigest.ComputeSha256(targetBytes), plan.TargetBytes.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Rewritten Page Layout bytes differ from the sealed target artifact.");
            }

            var freshProbe = PublishingPageLayoutTargetInspector.Inspect(context, plan);
            var freshAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(plan, freshProbe);
            if (!freshAdmission.IsEligible)
            {
                throw new InvalidOperationException("Fresh Page Layout preflight no longer satisfies the sealed plan.");
            }

            if (freshAdmission.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseOwned)
            {
                return false;
            }

            if (freshAdmission.Disposition != PublishingPageLayoutMaterializationDisposition.CreateOwned)
            {
                throw new InvalidOperationException($"Unexpected Page Layout admission disposition: {freshAdmission.Disposition}.");
            }

            var rootWeb = context.Site.RootWeb;
            var gallery = rootWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            context.Load(gallery, value => value.ForceCheckout, value => value.EnableVersioning, value => value.EnableModeration);
            context.Load(gallery.RootFolder, value => value.ServerRelativeUrl);
            context.Load(rootWeb.AvailableContentTypes, values => values.Include(value => value.Id, value => value.Name));
            context.ExecuteQueryRetry();
            var targetDirectory = plan.TargetServerRelativeUrl.Substring(0, plan.TargetServerRelativeUrl.LastIndexOf('/'));
            if (!string.Equals(targetDirectory, gallery.RootFolder.ServerRelativeUrl.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"The create-only Page Layout target must be directly under '{gallery.RootFolder.ServerRelativeUrl}', not '{targetDirectory}'.");
            }

            var associatedContentTypeId = plan.ContentTypeSchema?.ContentTypeId ?? freshProbe.ResolvedAssociatedContentTypeId;
            var associatedContentType = rootWeb.AvailableContentTypes.FirstOrDefault(value =>
                string.Equals(value.Id.StringValue, associatedContentTypeId, StringComparison.OrdinalIgnoreCase));
            if (associatedContentType == null)
            {
                throw new InvalidOperationException($"Target associated content type is unavailable: {plan.AssociatedContentTypeName}.");
            }

            var upload = gallery.RootFolder.Files.Add(new FileCreationInformation
            {
                Content = targetBytes,
                Url = plan.TargetServerRelativeUrl,
                Overwrite = false
            });
            context.Load(upload, value => value.CheckOutType, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            if ((gallery.ForceCheckout || gallery.EnableVersioning) && upload.CheckOutType == CheckOutType.None)
            {
                upload.CheckOut();
                context.ExecuteQueryRetry();
            }

            var item = upload.ListItemAllFields;
            item["Title"] = snapshot.Title ?? plan.TargetPageLayoutName;
            item["MasterPageDescription"] = $"PnP migration digest-owned Page Layout for {plan.AssociatedContentTypeName}.";
            item["ContentTypeId"] = SystemPageLayoutContentTypeId;
            item["PublishingAssociatedContentType"] = $";#{associatedContentType.Name};#{associatedContentType.Id};#";
            item["UIVersion"] = "15";
            item.Update();
            if (gallery.ForceCheckout || gallery.EnableVersioning)
            {
                upload.CheckIn("PnP migration Page Layout materialization.", CheckinType.MajorCheckIn);
                if (gallery.EnableModeration)
                {
                    upload.Publish("PnP migration Page Layout materialization.");
                }
            }

            context.ExecuteQueryRetry();
            Verify(context, plan);
            return true;
        }

        private static void Verify(ClientContext context, PublishingPageLayoutMaterializationPlan plan)
        {
            var readback = PublishingPageLayoutTargetInspector.Inspect(context, plan);
            var admission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(plan, readback);
            var expected = plan.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                ? PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                : PublishingPageLayoutMaterializationDisposition.ReuseOwned;
            if (!admission.IsEligible || admission.Disposition != expected)
            {
                throw new InvalidOperationException("Fresh Page Layout readback differs from the sealed plan.");
            }
        }

    }
}
