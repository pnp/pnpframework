using System;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.PublishingPages
{
    public static class PublishingPagePackageSerializer
    {
        private static readonly JsonSerializerOptions CanonicalOptions = CreateOptions(false);

        private static readonly JsonSerializerOptions IndentedOptions = CreateOptions(true);

        public static string Serialize<T>(T value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return JsonSerializer.Serialize(value, IndentedOptions) + Environment.NewLine;
        }

        public static T Deserialize<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentException("Package JSON is required.", nameof(json));
            }

            var value = JsonSerializer.Deserialize<T>(json, IndentedOptions);
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

            return JsonSerializer.Serialize(value, CanonicalOptions);
        }

        private static JsonSerializerOptions CreateOptions(bool writeIndented)
        {
            var options = new JsonSerializerOptions
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.Never,
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
