using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Fields
{
    public sealed class PageFieldValueSnapshot
    {
        public Guid Id { get; set; }

        public string InternalName { get; set; }

        public string Title { get; set; }

        public string TypeAsString { get; set; }

        public string SchemaXml { get; set; }

        public bool ReadOnly { get; set; }

        public bool Hidden { get; set; }

        public bool Required { get; set; }

        public bool HasValue { get; set; }

        public PageFieldValueKind Kind { get; set; }

        public string Value { get; set; }

        public IList<string> StringValues { get; set; } = new List<string>();

        public PageUrlValueSnapshot UrlValue { get; set; }

        public IList<PageLookupValueSnapshot> LookupValues { get; set; } = new List<PageLookupValueSnapshot>();

        public IList<PageTaxonomyValueSnapshot> TaxonomyValues { get; set; } = new List<PageTaxonomyValueSnapshot>();

        public TaxonomyFieldRelationshipBindingSnapshot TaxonomyBinding { get; set; }

        public string TaxonomyValueSetSha256 { get; set; }

        public string BinaryBase64 { get; set; }

        public string RawType { get; set; }

        public string RawValue { get; set; }

        public string RawValueJson { get; set; }

        public PageCaptureStatus CaptureStatus { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
