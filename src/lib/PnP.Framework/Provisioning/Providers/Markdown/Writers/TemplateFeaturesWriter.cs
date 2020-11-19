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
    //[TemplateSchemaWriter(WriterSequence = 1070,
    //    Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateFeaturesWriter// : PnPBaseSchemaWriter<Feature>
    {
        public /*override*/ void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            writer.WriteLine("# Features");
            writer.WriteLine();
            writer.WriteLine($"Coming soon");
            writer.WriteLine();
        }
    }
}
