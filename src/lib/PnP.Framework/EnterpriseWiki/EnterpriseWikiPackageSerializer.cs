using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PnP.Framework.EnterpriseWiki
{
    public static class EnterpriseWikiPackageSerializer
    {
        public const string SchemaVersion = "pnp-enterprise-wiki-package/v1";
        public const string DefaultPackageFileName = "enterprise-wiki-package.json";
        public const string DefaultReportFileName = "enterprise-wiki-report.md";
        public const string DefaultReceiptFileName = "enterprise-wiki-copy-receipt.json";

        private static readonly JsonSerializerOptions CanonicalOptions = CreateOptions(false);
        private static readonly JsonSerializerOptions IndentedOptions = CreateOptions(true);

        public static string ComputeSnapshotDigest(EnterpriseWikiSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            return ComputeSha256(JsonSerializer.Serialize(snapshot, CanonicalOptions));
        }

        public static string ComputePlanDigest(EnterpriseWikiMigrationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return ComputeSha256(JsonSerializer.Serialize(plan, CanonicalOptions));
        }

        public static string ComputeSha256(string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            using (var algorithm = SHA256.Create())
            {
                var bytes = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value));
                var builder = new StringBuilder(bytes.Length * 2);
                foreach (var item in bytes)
                {
                    builder.Append(item.ToString("x2"));
                }

                return builder.ToString();
            }
        }

        public static string Save(string path, EnterpriseWikiMigrationPackage package, bool overwrite = false)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("A package path or directory is required.", nameof(path));
            }

            Validate(package);
            var packagePath = ResolvePackagePath(path);
            var directory = Path.GetDirectoryName(packagePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (File.Exists(packagePath) && !overwrite)
            {
                throw new IOException($"The package file already exists: {packagePath}");
            }

            File.WriteAllText(packagePath, JsonSerializer.Serialize(package, IndentedOptions) + Environment.NewLine, new UTF8Encoding(false));
            var reportPath = Path.Combine(Path.GetDirectoryName(packagePath) ?? string.Empty, DefaultReportFileName);
            if (!File.Exists(reportPath) || overwrite)
            {
                File.WriteAllText(reportPath, BuildReport(package), new UTF8Encoding(false));
            }

            return packagePath;
        }

        public static EnterpriseWikiMigrationPackage Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("A package path or directory is required.", nameof(path));
            }

            var packagePath = ResolvePackagePath(path);
            if (!File.Exists(packagePath))
            {
                throw new FileNotFoundException("Enterprise Wiki package not found.", packagePath);
            }

            var package = JsonSerializer.Deserialize<EnterpriseWikiMigrationPackage>(File.ReadAllText(packagePath), IndentedOptions);
            Validate(package);
            return package;
        }

        public static string SaveReceipt(string path, EnterpriseWikiCopyReceipt receipt, bool overwrite = false)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("A receipt path or directory is required.", nameof(path));
            }

            if (receipt == null)
            {
                throw new ArgumentNullException(nameof(receipt));
            }

            var fullPath = Path.GetFullPath(path);
            var receiptPath = Directory.Exists(fullPath) || string.IsNullOrEmpty(Path.GetExtension(fullPath))
                ? Path.Combine(fullPath, DefaultReceiptFileName)
                : fullPath;
            var directory = Path.GetDirectoryName(receiptPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (File.Exists(receiptPath) && !overwrite)
            {
                throw new IOException($"The receipt file already exists: {receiptPath}");
            }

            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, IndentedOptions) + Environment.NewLine, new UTF8Encoding(false));
            return receiptPath;
        }

        public static void Validate(EnterpriseWikiMigrationPackage package)
        {
            if (package == null)
            {
                throw new InvalidDataException("The Enterprise Wiki package is empty.");
            }

            if (!string.Equals(package.SchemaVersion, SchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported Enterprise Wiki package schema '{package.SchemaVersion}'.");
            }

            if (package.Snapshot == null || package.Plan == null)
            {
                throw new InvalidDataException("The Enterprise Wiki package must contain both a snapshot and a migration plan.");
            }

            var snapshotDigest = ComputeSnapshotDigest(package.Snapshot);
            if (!string.Equals(snapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The Enterprise Wiki snapshot digest does not match the package payload.");
            }

            if (!string.Equals(package.Plan.SourceSnapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The migration plan does not reference the sealed snapshot in this package.");
            }

            var planDigest = ComputePlanDigest(package.Plan);
            if (!string.Equals(planDigest, package.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The Enterprise Wiki plan digest does not match the package payload.");
            }
        }

        private static string ResolvePackagePath(string path)
        {
            var fullPath = Path.GetFullPath(path);
            if (Directory.Exists(fullPath) || string.IsNullOrEmpty(Path.GetExtension(fullPath)))
            {
                return Path.Combine(fullPath, DefaultPackageFileName);
            }

            return fullPath;
        }

        private static string BuildReport(EnterpriseWikiMigrationPackage package)
        {
            var report = package.Report ?? new EnterpriseWikiCustomerReport();
            var builder = new StringBuilder();
            builder.AppendLine("# Enterprise Wiki migration report");
            builder.AppendLine();
            builder.AppendLine($"- State: `{package.State}`");
            builder.AppendLine($"- Source: `{package.Snapshot.Source.PageServerRelativeUrl}`");
            builder.AppendLine($"- Target: `{package.Plan.TargetPageServerRelativeUrl}`");
            builder.AppendLine($"- Snapshot SHA-256: `{package.SnapshotDigest}`");
            builder.AppendLine($"- Plan SHA-256: `{package.PlanDigest}`");
            builder.AppendLine();
            if (!string.IsNullOrWhiteSpace(report.Summary))
            {
                builder.AppendLine(report.Summary);
                builder.AppendLine();
            }

            AppendList(builder, "Captured ingredients", report.CapturedIngredients);
            AppendList(builder, "Blockers", report.Blockers);
            AppendList(builder, "Warnings", report.Warnings);
            return builder.ToString();
        }

        private static void AppendList(StringBuilder builder, string heading, System.Collections.Generic.IEnumerable<string> values)
        {
            var items = (values ?? Array.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
            builder.AppendLine($"## {heading}");
            builder.AppendLine();
            if (items.Length == 0)
            {
                builder.AppendLine("- None");
            }
            else
            {
                foreach (var item in items)
                {
                    builder.AppendLine($"- {item}");
                }
            }

            builder.AppendLine();
        }

        private static JsonSerializerOptions CreateOptions(bool writeIndented)
        {
            var options = new JsonSerializerOptions
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                PropertyNameCaseInsensitive = false,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = writeIndented
            };
            options.Converters.Add(new JsonStringEnumConverter());
            return options;
        }
    }
}
