using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectiontermsetid:[termsetname]}",
        Description = "Returns the id of the given termset name located in the sitecollection termgroup",
        Example = "{sitecollectiontermsetid:MyTermset}",
        Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
    internal class SiteCollectionTermSetIdToken : TokenDefinition
    {
        private readonly string _value;

        public SiteCollectionTermSetIdToken(Web web, string termsetName, Guid id)
            : base(web, $"{{sitecollectiontermsetid:{Regex.Escape(termsetName)}}}")
        {
            _value = id.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}