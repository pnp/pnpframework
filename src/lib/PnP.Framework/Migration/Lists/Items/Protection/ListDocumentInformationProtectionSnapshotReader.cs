using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Items.Protection
{
    internal static class ListDocumentInformationProtectionSnapshotReader
    {
        public static ListDocumentInformationProtectionSnapshot Read(
            IDictionary<string, object> fieldValues)
        {
            var labelId = ReadValue(fieldValues, "_IpLabelId");
            if (string.IsNullOrWhiteSpace(labelId))
            {
                return null;
            }

            return new ListDocumentInformationProtectionSnapshot
            {
                LabelId = labelId,
                AssignmentMethod = ReadValue(fieldValues, "_IpLabelAssignmentMethod"),
                HasUserDefinedProtection = ReadValue(fieldValues, "_HasUserDefinedProtection"),
                OwnerEmail = ReadValue(fieldValues, "_IpLabelOwnerEmail"),
                LabelHash = ReadValue(fieldValues, "_IpLabelHash"),
                PromotionCtagVersion = ReadValue(fieldValues, "_IpLabelPromotionCtagVersion"),
                DecryptSkipReason = ReadMetaInfoValue(fieldValues, "vti_decryptskipreason")
            };
        }

        private static string ReadValue(IDictionary<string, object> fieldValues, string internalName)
        {
            if (fieldValues == null)
            {
                return null;
            }
            foreach (var value in fieldValues)
            {
                if (string.Equals(value.Key, internalName, StringComparison.OrdinalIgnoreCase))
                {
                    return value.Value == null ? null : Convert.ToString(value.Value);
                }
            }
            return null;
        }

        private static string ReadMetaInfoValue(
            IDictionary<string, object> fieldValues,
            string propertyName)
        {
            var metaInfo = ReadValue(fieldValues, "MetaInfo");
            if (string.IsNullOrWhiteSpace(metaInfo))
            {
                return null;
            }
            foreach (var line in metaInfo.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                var separator = line.IndexOf('|');
                var propertySeparator = line.IndexOf(':');
                if (separator <= propertySeparator || propertySeparator <= 0
                    || !string.Equals(
                        line.Substring(0, propertySeparator),
                        propertyName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                return line.Substring(separator + 1);
            }
            return null;
        }
    }
}
