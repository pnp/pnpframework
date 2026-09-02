using System;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomySearchIdentity
    {
        public static bool IsExact(string value, Guid storeId, Guid setId, Guid termId)
        {
            if (string.Equals(value, "UNVALIDATED", StringComparison.Ordinal))
            {
                return true;
            }
            var parts = (value ?? string.Empty).Split('|');
            Guid observedStore;
            Guid observedSet;
            Guid observedTerm;
            return parts.Length >= 3
                && TryDecompressGuid(parts[0], out observedStore)
                && observedStore == storeId
                && TryDecompressGuid(parts[1], out observedSet)
                && observedSet == setId
                && TryDecompressGuid(parts[2], out observedTerm)
                && observedTerm == termId;
        }

        public static string Rewrite(string value, Guid targetStoreId, Guid targetSetId, Guid termId)
        {
            if (string.Equals(value, "UNVALIDATED", StringComparison.Ordinal))
            {
                return value;
            }
            var parts = (value ?? string.Empty).Split('|');
            Guid observedTerm;
            if (parts.Length < 3
                || !TryDecompressGuid(parts[2], out observedTerm)
                || observedTerm != termId)
            {
                throw new InvalidOperationException($"Taxonomy hidden-list search identity is not sealed to Term '{termId:D}'.");
            }
            parts[0] = Convert.ToBase64String(targetStoreId.ToByteArray());
            parts[1] = Convert.ToBase64String(targetSetId.ToByteArray());
            return string.Join("|", parts);
        }

        private static bool TryDecompressGuid(string value, out Guid id)
        {
            try
            {
                var bytes = Convert.FromBase64String(value);
                if (bytes.Length == 16)
                {
                    id = new Guid(bytes);
                    return true;
                }
            }
            catch (FormatException)
            {
            }

            id = Guid.Empty;
            return false;
        }
    }
}
