using System;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Packaging
{
    internal static class MigrationContractSerializer
    {
        private static readonly JsonSerializerOptions CanonicalOptions = CreateOptions(false);

        private static readonly JsonSerializerOptions IndentedOptions = CreateOptions(true);

        public static string SerializeCanonical<T>(T value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return JsonSerializer.Serialize(value, CanonicalOptions);
        }

        public static string SerializeIndented<T>(T value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return JsonSerializer.Serialize(value, IndentedOptions);
        }

        public static T Deserialize<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentException("Migration contract JSON is required.", nameof(json));
            }

            return JsonSerializer.Deserialize<T>(json, IndentedOptions);
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
