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
    [TemplateSchemaWriter(WriterSequence = 1070, Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateFeaturesWriter : PnPBaseSchemaWriter<Feature>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            if (template.Features != null)
            {
                WriteHeader("Features", 1, writer);
                if (template.Features.SiteFeatures != null && template.Features.SiteFeatures.Count > 0)
                {
                    WriteHeader("Site Collection Features", 2, writer);
                    foreach (var feature in template.Features.SiteFeatures)
                    {
                        WriteTextField(feature.Deactivate ? "Deactivate" : "Activate", feature.Id.ToString(), writer);
                    }
                }
                if (template.Features.WebFeatures != null && template.Features.WebFeatures.Count > 0)
                {
                    WriteHeader("Site Features", 2, writer);
                    foreach (var feature in template.Features.WebFeatures)
                    {
                        WriteTextField(feature.Deactivate ? "Deactivate" : "Activate", feature.Id.ToString(), writer);
                    }
                }
                writer.WriteLine();
            }
        }
    }
}
