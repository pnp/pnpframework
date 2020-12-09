using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Xml.Linq;

namespace PnP.Framework.Provisioning.Providers.Markdown.Writers
{
    /// <summary>
    /// Class to write out the markdown for the base properties
    /// </summary>
    [TemplateSchemaWriter(WriterSequence = 1050,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateContentTypesWriter : PnPBaseSchemaWriter<ContentType>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            if (template.ContentTypes != null && template.ContentTypes.Count > 0)
            {
                WriteHeader("Content Types", 1, writer);
                writer.WriteLine();
                writer.WriteLine("| Name | Description | Group |");
                writer.WriteLine("| :------------- | :----------: | :----------: |");
                using (TextWriter groupDetailsWriter = new StringWriter())
                {

                    string currentGroup = "";

                    foreach (var ct in template.ContentTypes.OrderBy(o => o.Group))
                    {
                        if (currentGroup != ct.Group)
                        {
                            groupDetailsWriter.WriteLine("<br/>");
                            groupDetailsWriter.WriteLine();
                            WriteHeader($"Group - {ct.Group}", 3, groupDetailsWriter);
                            groupDetailsWriter.WriteLine();
                        }


                        writer.WriteLine($"|  {ct.Name} | {ct.Description}   | {ct.Group}   |");

                        groupDetailsWriter.WriteLine("<br/>");
                        groupDetailsWriter.WriteLine();
                        WriteHeader(ct.Name, 3, groupDetailsWriter);
                        groupDetailsWriter.WriteLine();
                        groupDetailsWriter.WriteLine($"**Description** - {ct.Description}");
                        groupDetailsWriter.WriteLine();
                        groupDetailsWriter.WriteLine("**Fields**:");
                        groupDetailsWriter.WriteLine();
                        //TODO: get parent content type
                        groupDetailsWriter.WriteLine("| Name   |     Required     | Hidden       |");
                        //TODO: cna we list with field is from a parent content type
                        groupDetailsWriter.WriteLine("| :------------- | :----------: | :----------: |");

                        foreach (var ctField in ct.FieldRefs)
                        {
                            groupDetailsWriter.WriteLine($"| {ctField.Name}   | {ctField.Required.ToString()}     | {ctField.Hidden.ToString()}       |");
                        }
                    }
                    writer.WriteLine(groupDetailsWriter.ToString());
                    writer.WriteLine("<br/>");
                    writer.WriteLine();
                }
            }
        }
    }
}
