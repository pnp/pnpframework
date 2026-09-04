using System;
using System.Text;

namespace PnP.Framework.Migration.Lists.Items
{
    internal static class ListBinaryPayloadClassifier
    {
        private static readonly byte[] CompoundFileHeader =
        {
            0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1
        };

        private static readonly byte[] DrmTransformMarker =
            Encoding.Unicode.GetBytes("DRMEncryptedTransform");

        private static readonly byte[] EncryptedPackageMarker =
            Encoding.Unicode.GetBytes("EncryptedPackage");

        public static ListBinaryRepresentationKind Classify(byte[] payload)
        {
            return HasPrefix(payload, CompoundFileHeader)
                && Contains(payload, DrmTransformMarker)
                && Contains(payload, EncryptedPackageMarker)
                    ? ListBinaryRepresentationKind.InformationRightsManagedEnvelope
                    : ListBinaryRepresentationKind.OrdinaryFilePayload;
        }

        private static bool HasPrefix(byte[] payload, byte[] prefix)
        {
            if (payload == null || payload.Length < prefix.Length)
            {
                return false;
            }
            for (var index = 0; index < prefix.Length; index++)
            {
                if (payload[index] != prefix[index])
                {
                    return false;
                }
            }
            return true;
        }

        private static bool Contains(byte[] payload, byte[] value)
        {
            if (payload == null || value == null || value.Length == 0 || payload.Length < value.Length)
            {
                return false;
            }
            for (var offset = 0; offset <= payload.Length - value.Length; offset++)
            {
                var matched = true;
                for (var index = 0; index < value.Length; index++)
                {
                    if (payload[offset + index] == value[index])
                    {
                        continue;
                    }
                    matched = false;
                    break;
                }
                if (matched)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
