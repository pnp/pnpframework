using System;
using System.IO;
using System.Text;
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

        /// <summary>
        /// Serializes a digest-sealed contract without mutating the source
        /// instance. The selected root property is emitted as JSON null in its
        /// original position, matching the historical digest representation.
        /// </summary>
        public static string SerializeCanonicalWithNullRootProperty<T>(
            T value,
            string clrPropertyName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }
            if (string.IsNullOrWhiteSpace(clrPropertyName))
            {
                throw new ArgumentException(
                    "A root contract property name is required.",
                    nameof(clrPropertyName));
            }

            var jsonPropertyName = CanonicalOptions.PropertyNamingPolicy == null
                ? clrPropertyName
                : CanonicalOptions.PropertyNamingPolicy.ConvertName(clrPropertyName);
            using (var document = JsonDocument.Parse(SerializeCanonical(value)))
            {
                if (document.RootElement.ValueKind != JsonValueKind.Object)
                {
                    throw new InvalidOperationException(
                        "A digest-sealed migration contract must serialize as a JSON object.");
                }

                using (var stream = new MemoryStream())
                {
                    using (var writer = new Utf8JsonWriter(
                        stream,
                        new JsonWriterOptions
                        {
                            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                            Indented = false
                        }))
                    {
                        var replaced = false;
                        writer.WriteStartObject();
                        foreach (var property in document.RootElement.EnumerateObject())
                        {
                            writer.WritePropertyName(property.Name);
                            if (string.Equals(
                                property.Name,
                                jsonPropertyName,
                                StringComparison.Ordinal))
                            {
                                writer.WriteNullValue();
                                replaced = true;
                            }
                            else
                            {
                                property.Value.WriteTo(writer);
                            }
                        }
                        writer.WriteEndObject();
                        writer.Flush();

                        if (!replaced)
                        {
                            throw new InvalidOperationException(
                                "Digest property '" + clrPropertyName
                                + "' is not present on the serialized migration contract.");
                        }
                    }

                    return Encoding.UTF8.GetString(stream.ToArray());
                }
            }
        }

        public static T Deserialize<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentException("Migration contract JSON is required.", nameof(json));
            }

            return JsonSerializer.Deserialize<T>(json, IndentedOptions);
        }

        public static T Deserialize<T>(Stream jsonStream)
        {
            if (jsonStream == null)
            {
                throw new ArgumentNullException(nameof(jsonStream));
            }
            if (!jsonStream.CanRead)
            {
                throw new ArgumentException("Migration contract JSON stream must be readable.", nameof(jsonStream));
            }
            if (jsonStream.CanSeek && jsonStream.Length - jsonStream.Position == 0)
            {
                throw new ArgumentException("Migration contract JSON is required.", nameof(jsonStream));
            }

            return JsonSerializer.Deserialize<T>(jsonStream, IndentedOptions);
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
