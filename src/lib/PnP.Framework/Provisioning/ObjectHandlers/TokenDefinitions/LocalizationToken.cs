using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{localization:[key]}",
     Description = "Returns a value from a in the template provided resource file given the locale of the site that the template is applied to",
     Example = "{localization:MyListTitle}",
     Returns = "My List Title")]
    internal class LocalizationToken : TokenDefinition
    {
        private readonly int _webLanguage;
        private readonly int? _defaultLcid;
        private readonly Dictionary<int, ResourceEntry> _entriesByLanguage;

        public IReadOnlyList<ResourceEntry> ResourceEntries { get; }

        public LocalizationToken(Web web, string key, List<ResourceEntry> resourceEntries, int? defaultLcid)
            : base(web, $"{{loc:{Regex.Escape(key)}}}", $"{{localize:{Regex.Escape(key)}}}", $"{{localization:{Regex.Escape(key)}}}", $"{{resource:{Regex.Escape(key)}}}", $"{{res:{Regex.Escape(key)}}}")
        {
            ResourceEntries = resourceEntries;
            _defaultLcid = defaultLcid;
            _webLanguage = (int)web.Language;
            _entriesByLanguage = new Dictionary<int, ResourceEntry>(capacity: resourceEntries.Count + 1);

            for (var index = 0; index < resourceEntries.Count; index++)
            {
                var entry = resourceEntries[index];
                _entriesByLanguage[entry.LCID] = entry;
            }
        }

        public override string GetReplaceValue()
        {
            if (_entriesByLanguage.TryGetValue(_webLanguage, out ResourceEntry entry)
                // Fallback to default LCID.
                || (_defaultLcid.HasValue && _entriesByLanguage.TryGetValue(_defaultLcid.Value, out entry)))
            {
                return entry.Value;
            }

            // Fallback to old logic as for me _defaultLCID has always a Value i.e. 0 or the correct LCID.
            return ResourceEntries[0].Value;
        }
    }
}