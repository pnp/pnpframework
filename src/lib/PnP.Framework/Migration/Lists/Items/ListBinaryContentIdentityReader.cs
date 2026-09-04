using System;
using System.Collections.Generic;
using System.Text.Json;

namespace PnP.Framework.Migration.Lists.Items
{
    internal static class ListBinaryContentIdentityReader
    {
        private const string MediaServiceMetadataFieldName = "MediaServiceMetadata";
        private const string MetaInfoFieldName = "MetaInfo";
        private const string MediaServiceMetadataPrefix = "vti_mediaservicemetadata:SW|";
        private const string DirectEvidenceSource = "SharePoint.ListItem.MediaServiceMetadata";
        private const string MetaInfoEvidenceSource = "SharePoint.MetaInfo.vti_mediaservicemetadata";

        public static ListBinaryContentIdentitySnapshot Read(
            IDictionary<string, object> fieldValues)
        {
            var directMetadata = FindFieldValue(fieldValues, MediaServiceMetadataFieldName);
            var directIdentity = ReadJson(directMetadata, DirectEvidenceSource);
            if (directIdentity != null)
            {
                return directIdentity;
            }

            var metaInfo = FindFieldValue(fieldValues, MetaInfoFieldName);
            if (string.IsNullOrWhiteSpace(metaInfo))
            {
                return null;
            }

            foreach (var line in metaInfo.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                if (!line.StartsWith(MediaServiceMetadataPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                var json = line.Substring(MediaServiceMetadataPrefix.Length).Trim();
                if (json.Length == 0)
                {
                    continue;
                }
                var identity = ReadJson(json, MetaInfoEvidenceSource);
                if (identity != null)
                {
                    return identity;
                }
            }
            return null;
        }

        private static string FindFieldValue(IDictionary<string, object> fieldValues, string fieldName)
        {
            if (fieldValues == null)
            {
                return null;
            }
            foreach (var value in fieldValues)
            {
                if (string.Equals(value.Key, fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    return Convert.ToString(value.Value);
                }
            }
            return null;
        }

        private static ListBinaryContentIdentitySnapshot ReadJson(string json, string evidenceSource)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }
            try
            {
                using (var document = JsonDocument.Parse(json))
                {
                    var root = document.RootElement;
                    var quickXorHash = ReadString(root, "quickxorhash");
                    var contentTag = ReadString(root, "ctag");
                    if (string.IsNullOrWhiteSpace(quickXorHash)
                        && string.IsNullOrWhiteSpace(contentTag))
                    {
                        return null;
                    }
                    return new ListBinaryContentIdentitySnapshot
                    {
                        QuickXorHash = quickXorHash,
                        ContentTag = contentTag,
                        EvidenceSource = evidenceSource
                    };
                }
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static string ReadString(JsonElement value, string name)
        {
            foreach (var property in value.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)
                    && property.Value.ValueKind == JsonValueKind.String)
                {
                    return property.Value.GetString();
                }
            }
            return null;
        }
    }
}
