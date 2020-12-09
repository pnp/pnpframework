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
            if (template.WebSettings != null)
            {
                var webSettings = template.WebSettings;
                WriteHeader("Site Summary", 1, writer);
                WriteTextField(string.IsNullOrEmpty(webSettings.Title) ? "To be set manually" : webSettings.Title, "Site name", writer);
                writer.WriteLine(GetSiteTemplateNameFromTemplateCode(template.BaseSiteTemplate));
                writer.WriteLine();
                WriteTextField(template.WebSettings.WelcomePage, "Landing page", writer);
                bool commentsEnabled = template.WebSettings.CommentsOnSitePagesDisabled;
                string commentsEnabledText = "Comments on pages enabled";
                if (commentsEnabled)
                {
                    commentsEnabledText = "Comments on pages disabled";
                }
                writer.WriteLine(commentsEnabledText);
                writer.WriteLine();
            }
        }
    }
}
