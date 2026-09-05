using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Markup
{
    public sealed class PageDirectiveSnapshot
    {
        public string Inherits { get; set; }

        public string MasterPageFile { get; set; }

        public string Language { get; set; }

        public string CodeBehind { get; set; }

        public string CodeFile { get; set; }

        public IList<PageDirectiveAttribute> Attributes { get; set; } = new List<PageDirectiveAttribute>();
    }
}
