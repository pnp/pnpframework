using PnP.Framework.Migration.Packaging;
using System;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    internal static class PublishingPageValidationCanonical
    {
        public static bool Equals<T>(T expected, T actual)
        {
            return string.Equals(
                MigrationContractSerializer.SerializeCanonical(expected),
                MigrationContractSerializer.SerializeCanonical(actual),
                StringComparison.Ordinal);
        }
    }
}
