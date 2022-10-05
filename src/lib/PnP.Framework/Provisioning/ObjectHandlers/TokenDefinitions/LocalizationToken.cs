using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{localization:[key]}",
     Description = "Returns a value from a in the template provided resource file given the locale of the site that the template is applied to",
     Example = "{localization:MyListTitle}",
     Returns = "My List Title")]
    internal class LocalizationToken : TokenDefinition
    {
        private readonly List<ResourceEntry> _resourceEntries;
        private readonly int? _defaultLCID;

        public LocalizationToken(Web web, string key, List<ResourceEntry> resourceEntries, int? defaultLCID)
            : base(web, $"{{loc:{Regex.Escape(key)}}}", $"{{localize:{Regex.Escape(key)}}}", $"{{localization:{Regex.Escape(key)}}}", $"{{resource:{Regex.Escape(key)}}}", $"{{res:{Regex.Escape(key)}}}")
        {
            _resourceEntries = resourceEntries;
            _defaultLCID = defaultLCID;
        }

        public override string GetReplaceValue()
        {
            var entry = _resourceEntries.FirstOrDefault(r => r.LCID == this.Web.Language);
            if (entry != null)
            {
                return entry.Value;
            }
            else
            {
                // fallback to default LCID or to the first resource string
                var defaultEntry = _defaultLCID.HasValue ?
                    _resourceEntries.FirstOrDefault(r => r.LCID == _defaultLCID) :
                    _resourceEntries.First();

                return defaultEntry.Value;
            }

        }

        public List<ResourceEntry> ResourceEntries
        {
            get { return _resourceEntries; }
        }
    }
}