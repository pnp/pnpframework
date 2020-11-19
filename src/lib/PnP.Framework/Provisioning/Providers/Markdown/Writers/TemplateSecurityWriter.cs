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
    [TemplateSchemaWriter(WriterSequence = 1020,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateSecurityWriter : PnPBaseSchemaWriter<SiteSecurity>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            writer.WriteLine("# Permissions");
            writer.WriteLine();
            writer.WriteLine($"Permissions should be left as default, making use of the default Visitor, Member and Owner groups.");
            writer.WriteLine();
        }
    }
}
