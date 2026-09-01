using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    internal static class ContentTypeRuntimeCatalog
    {
        private static readonly HashSet<string> KnownTargetRuntimeIds = new HashSet<string>(new[]
        {
            BuiltInContentTypeId.Item,
            BuiltInContentTypeId.Document,
            BuiltInContentTypeId.Folder,
            BuiltInContentTypeId.DocumentSet,
            // SharePoint runtime children used by Document Set and Asset Library templates.
            // Keep this list exact: a prefix match would incorrectly classify custom
            // document content types as target-runtime schema.
            "0x0120D520A8",
            "0x0120D520A808",
            "0x0101009148F5A04DDD49CBA7127AADA5FB792B",
            "0x0101009148F5A04DDD49CBA7127AADA5FB792B00291D173ECE694D56B19D111489C4369D",
            "0x0101009148F5A04DDD49CBA7127AADA5FB792B006973ACD696DC4858A76371B2FB2F439A",
            "0x0101009148F5A04DDD49CBA7127AADA5FB792B00AADE34325A8B49CDA8BB4DB53328F214"
        }, StringComparer.OrdinalIgnoreCase);

        public static bool IsTargetRuntime(string contentTypeId)
        {
            return !string.IsNullOrWhiteSpace(contentTypeId) && KnownTargetRuntimeIds.Contains(contentTypeId);
        }
    }
}
