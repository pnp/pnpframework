using System;
using System.IO;
using PnP.Framework.Migration.Packaging;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public static class PublishingPagePackageSerializer
    {
        public static string Serialize<T>(T value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return MigrationContractSerializer.SerializeIndented(value) + Environment.NewLine;
        }

        public static T Deserialize<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentException("Package JSON is required.", nameof(json));
            }

            var value = MigrationContractSerializer.Deserialize<T>(json);
            if (value == null)
            {
                throw new InvalidDataException($"The JSON payload did not contain a {typeof(T).Name} value.");
            }

            return value;
        }

        public static T Deserialize<T>(Stream jsonStream)
        {
            if (jsonStream == null)
            {
                throw new ArgumentNullException(nameof(jsonStream));
            }

            var value = MigrationContractSerializer.Deserialize<T>(jsonStream);
            if (value == null)
            {
                throw new InvalidDataException($"The JSON payload did not contain a {typeof(T).Name} value.");
            }

            return value;
        }

        internal static string SerializeCanonical<T>(T value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return MigrationContractSerializer.SerializeCanonical(value);
        }
    }
}
