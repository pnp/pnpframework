using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Xml.Linq;

namespace PnP.Framework.Provisioning.Providers.Markdown.Writers
{
    /// <summary>
    /// Class to write out the markdown for the base properties
    /// </summary>
    [TemplateSchemaWriter(WriterSequence = 1010,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateWebSettingsWriter : PnPBaseSchemaWriter<WebSettings>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            writer.WriteLine("# Site Summary");
            writer.WriteLine();
            writer.WriteLine($"Site name - To be set manually");
            writer.WriteLine();
            writer.WriteLine(GetSiteTemplateNameFromTemplateCode(template.BaseSiteTemplate));
            writer.WriteLine();
            writer.WriteLine($"Landing page - {template.WebSettings.WelcomePage}");
            writer.WriteLine();
            bool commentsEnabled = template.WebSettings.CommentsOnSitePagesDisabled;
            string commentsEnabledText = "Comments on pages enabled";
            if (commentsEnabled)
            {
                commentsEnabledText = "Comments on pages disabled";
            }
            writer.WriteLine(commentsEnabledText);
            writer.WriteLine();
            //TODO: set out search scope
            /*
            string searchScopeText = "Default search scope";
            if (template.WebSettings.SearchScope != SearchScopes.DefaultScope)
            {
                searchScopeText = $"Search scope - {template.WebSettings.SearchScope.ToString()}";
            }
            writer.WriteLine(searchScopeText);
            writer.WriteLine();
            */
        }
    }
}
