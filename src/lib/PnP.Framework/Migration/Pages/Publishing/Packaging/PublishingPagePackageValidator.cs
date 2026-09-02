using PnP.Framework.Migration.Packaging;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public static class PublishingPagePackageValidator
    {
        public static void ValidateExport(PublishingPageExportPackage package)
        {
            PublishingPageExportPackageValidator.Validate(package, null);
        }

        public static void ValidateExport(
            PublishingPageExportPackage package,
            IMigrationArtifactStore artifactStore)
        {
            PublishingPageExportPackageValidator.Validate(package, artifactStore);
        }

        public static void ValidateMigration(PublishingPageMigrationPackage package)
        {
            PublishingPageMigrationPackageValidator.Validate(package, null);
        }

        public static void ValidateMigration(
            PublishingPageMigrationPackage package,
            IMigrationArtifactStore artifactStore)
        {
            PublishingPageMigrationPackageValidator.Validate(package, artifactStore);
        }
    }
}
