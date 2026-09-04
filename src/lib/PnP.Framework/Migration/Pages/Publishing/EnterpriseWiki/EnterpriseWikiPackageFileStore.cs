using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Assessment;
using PnP.Framework.Migration.Pages.Publishing.Reporting;
using PnP.Framework.Migration.Packaging;
using System;
using System.IO;
using System.Text;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public static class EnterpriseWikiPackageFileStore
    {
        public const string DefaultExportFileName = "enterprise-wiki-export.json";

        public const string DefaultPackageFileName = "enterprise-wiki-package.json";

        public const string DefaultAssessmentFileName = "enterprise-wiki-assessment.json";

        public const string DefaultReportFileName = "enterprise-wiki-report.md";

        public const string DefaultReceiptFileName = "enterprise-wiki-import-receipt.json";

        public static string SaveExport(string path, PublishingPageExportPackage package, bool overwrite = false)
        {
            return SaveExport(path, package, null, overwrite);
        }

        public static string SaveExport(
            string path,
            PublishingPageExportPackage package,
            IMigrationArtifactStore artifactStore,
            bool overwrite = false)
        {
            PublishingPagePackageValidator.ValidateExport(package, artifactStore);
            var exportPath = ResolvePath(path, DefaultExportFileName);
            SaveText(exportPath, PublishingPagePackageSerializer.Serialize(package), overwrite);
            return exportPath;
        }

        public static PublishingPageExportPackage LoadExport(string path)
        {
            return LoadExport(path, null);
        }

        public static PublishingPageExportPackage LoadExport(string path, IMigrationArtifactStore artifactStore)
        {
            var exportPath = ResolveExistingPath(path, DefaultExportFileName, "Enterprise Wiki export");
            using var stream = OpenPackageReadStream(exportPath);
            var package = PublishingPagePackageSerializer.Deserialize<PublishingPageExportPackage>(stream);
            PublishingPagePackageValidator.ValidateExport(package, artifactStore);
            return package;
        }

        public static string SaveAssessment(
            string path,
            PublishingPageMigrationAssessment assessment,
            bool overwrite = false)
        {
            PublishingPageMigrationAssessmentValidator.Validate(assessment);
            var assessmentPath = ResolvePath(path, DefaultAssessmentFileName);
            SaveText(assessmentPath, PublishingPagePackageSerializer.Serialize(assessment), overwrite);
            return assessmentPath;
        }

        public static PublishingPageMigrationAssessment LoadAssessment(string path)
        {
            var assessmentPath = ResolveExistingPath(path, DefaultAssessmentFileName, "Enterprise Wiki assessment");
            using var stream = OpenPackageReadStream(assessmentPath);
            var assessment = PublishingPagePackageSerializer.Deserialize<PublishingPageMigrationAssessment>(stream);
            PublishingPageMigrationAssessmentValidator.Validate(assessment);
            return assessment;
        }

        public static string SaveMigration(string path, PublishingPageMigrationPackage package, bool overwrite = false)
        {
            return SaveMigration(path, package, null, overwrite);
        }

        public static string SaveMigration(
            string path,
            PublishingPageMigrationPackage package,
            IMigrationArtifactStore artifactStore,
            bool overwrite = false)
        {
            PublishingPagePackageValidator.ValidateMigration(package, artifactStore);
            var packagePath = ResolvePath(path, DefaultPackageFileName);
            var reportPath = Path.Combine(Path.GetDirectoryName(packagePath) ?? string.Empty, DefaultReportFileName);
            EnsureWritable(packagePath, overwrite);
            EnsureWritable(reportPath, overwrite);
            SaveText(packagePath, PublishingPagePackageSerializer.Serialize(package), true);
            SaveText(reportPath, PublishingPageMigrationReportBuilder.Build(package, artifactStore), true);
            return packagePath;
        }

        public static PublishingPageMigrationPackage LoadMigration(string path)
        {
            return LoadMigration(path, null);
        }

        public static PublishingPageMigrationPackage LoadMigration(string path, IMigrationArtifactStore artifactStore)
        {
            var packagePath = ResolveExistingPath(path, DefaultPackageFileName, "Enterprise Wiki migration package");
            using var stream = OpenPackageReadStream(packagePath);
            var package = PublishingPagePackageSerializer.Deserialize<PublishingPageMigrationPackage>(stream);
            PublishingPagePackageValidator.ValidateMigration(package, artifactStore);
            return package;
        }

        public static string SaveReceipt(string path, PublishingPageImportReceipt receipt, bool overwrite = false)
        {
            if (receipt == null)
            {
                throw new ArgumentNullException(nameof(receipt));
            }

            var receiptPath = ResolvePath(path, DefaultReceiptFileName);
            SaveText(receiptPath, PublishingPagePackageSerializer.Serialize(receipt), overwrite);
            return receiptPath;
        }

        private static string ResolveExistingPath(string path, string defaultFileName, string description)
        {
            var resolved = ResolvePath(path, defaultFileName);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"{description} not found.", resolved);
            }

            return resolved;
        }

        private static string ResolvePath(string path, string defaultFileName)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("A file path or directory is required.", nameof(path));
            }

            var fullPath = Path.GetFullPath(path);
            return Directory.Exists(fullPath) || string.IsNullOrEmpty(Path.GetExtension(fullPath))
                ? Path.Combine(fullPath, defaultFileName)
                : fullPath;
        }

        private static void SaveText(string path, string value, bool overwrite)
        {
            EnsureWritable(path, overwrite);
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(path, value, new UTF8Encoding(false));
        }

        private static FileStream OpenPackageReadStream(string path)
        {
            return new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                128 * 1024,
                FileOptions.SequentialScan);
        }

        private static void EnsureWritable(string path, bool overwrite)
        {
            if (File.Exists(path) && !overwrite)
            {
                throw new IOException($"The file already exists: {path}");
            }
        }
    }
}
