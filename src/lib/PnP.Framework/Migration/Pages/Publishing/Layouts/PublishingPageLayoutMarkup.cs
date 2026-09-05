using System.Collections.Generic;
using PnP.Framework.Migration.Pages.Markup;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal sealed class PublishingPageLayoutMarkup
    {
        public PageDirectiveSnapshot PageDirective { get; set; }

        public IList<PublishingPageLayoutRegistration> Registrations { get; set; } = new List<PublishingPageLayoutRegistration>();

        public IList<PublishingPageLayoutControl> Controls { get; set; } = new List<PublishingPageLayoutControl>();

        public IList<PublishingPageLayoutZone> Zones { get; set; } = new List<PublishingPageLayoutZone>();

        public IList<PublishingPageLayoutResourceReference> ResourceReferences { get; set; } = new List<PublishingPageLayoutResourceReference>();

        public IList<string> RequiredFieldIdentifiers { get; set; } = new List<string>();
    }
}
