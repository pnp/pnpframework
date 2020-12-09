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
    [TemplateSchemaWriter(WriterSequence = 1030,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateNavigationWriter : PnPBaseSchemaWriter<Navigation>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            writer.WriteLine("# Navigation");
            writer.WriteLine();
            writer.WriteLine($"## Left navigation:");
            writer.WriteLine();
            writer.WriteLine($"Leave as default.");
            writer.WriteLine();
            writer.WriteLine($"## Top navigation:");
            writer.WriteLine();
            writer.WriteLine($"Leave as default.");
            writer.WriteLine();
        }
    }
}
