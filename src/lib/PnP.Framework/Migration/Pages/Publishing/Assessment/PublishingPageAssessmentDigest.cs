using PnP.Framework.Migration.Pages.Publishing.Packaging;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageAssessmentDigest
    {
        public static string Compute(PublishingPageMigrationAssessment assessment)
        {
            var value = assessment.AssessmentDigest;
            assessment.AssessmentDigest = null;
            try
            {
                return PublishingPageDigest.ComputeSha256(
                    PublishingPagePackageSerializer.SerializeCanonical(assessment));
            }
            finally
            {
                assessment.AssessmentDigest = value;
            }
        }
    }
}
