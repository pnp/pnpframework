using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Markdown;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Xml;
using System.Xml.Linq;

namespace PnP.Framework.Provisioning.Providers.Markdown.Writers
{
    /// <summary>
    /// Class to write out the markdown for the base properties
    /// </summary>
    [TemplateSchemaWriter(WriterSequence = 1060,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateListInstancesWriter : PnPBaseSchemaWriter<ListInstance>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            if (template.Lists != null && template.Lists.Count > 0)
            {
                WriteHeader("Lists", 1, writer);
                WriteNewLine(writer);
                writer.WriteLine("| Name | Description |");
                writer.WriteLine("| :------------- | :----------: |");
                using (TextWriter detailsWriter = new StringWriter())
                {

                    foreach (var list in template.Lists)
                    {

                        writer.WriteLine($"|  {list.Title} | {list.Description}   |");
                        WriteNewLine(detailsWriter);
                        WriteHeader(list.Title, 2, detailsWriter);
                        WriteNewLine(detailsWriter);
                        WriteTextField(list.Description, "Description", detailsWriter);
                        WriteTextField(list.OnQuickLaunch.ToString(), "Show on quick launch?", detailsWriter);
                        WriteTextField(GetListTemplateNameFromTemplateCode(list.TemplateType.ToString()), "List template", detailsWriter);
                        WriteTextField(list.Url, "URL", detailsWriter);
                        WriteNewLine(detailsWriter);

                        WriteHeader("Versioning", 3, detailsWriter);
                        WriteTextField("TBC", "Require content approval?", detailsWriter);
                        WriteTextField(list.EnableVersioning.ToString(), "Create a version each time you edit an item in this list ?", detailsWriter);
                        WriteTextField("TBC", "Draft item security", detailsWriter);
                        WriteNewLine(detailsWriter);

                        WriteHeader("Advanced settings", 3, detailsWriter);
                        WriteTextField(list.ContentTypesEnabled.ToString(), "Allow management of content types?", detailsWriter);
                        WriteNewLine(detailsWriter);

                        WriteHeader("Content Types", 3, detailsWriter);
                        foreach (var binding in list.ContentTypeBindings)
                        {
                            WriteText($"- {binding.ContentTypeId}", detailsWriter);
                        }
                        WriteNewLine(detailsWriter);

                        WriteHeader("Views", 3, detailsWriter);
                        detailsWriter.WriteLine("| Display Name |  Default?  |   Name    |");
                        detailsWriter.WriteLine("| :------------- | :----------: | :----------: |");

                        TextWriter viewDetailsWriter = new StringWriter();

                        var xmlViewFields = from f in list.Views
                                            select XElement.Parse(f.SchemaXml).ToXmlElement();

                        foreach (var xmlField in xmlViewFields)
                        {
                            var viewDisplayName = xmlField.Attributes["DisplayName"].Value;
                            //var viewType = xmlField.Attributes["Type"].Value;
                            var viewName = xmlField.Attributes["Name"].Value;

                            detailsWriter.WriteLine($"| {viewDisplayName} | TBC | {viewName} |");

                            WriteHeader(viewDisplayName, 4, viewDetailsWriter);
                            WriteNewLine(viewDetailsWriter);

                            WriteAttributeField("Url", "View Url", viewDetailsWriter, xmlField);
                            WriteText("**Fields:**", viewDetailsWriter);

                            foreach (XmlElement fieldNode in xmlField.SelectNodes("//ViewFields//FieldRef"))
                            {
                                WriteText($"- {GetAttributeValue("Name", fieldNode)}", viewDetailsWriter);
                            }
                        }
                        detailsWriter.WriteLine(viewDetailsWriter.ToString());
                        WriteNewLine(viewDetailsWriter);

                        WriteText("NB: Currently the documentation assumes you are using Content Types so it just shows the field refs", viewDetailsWriter);

                        if (list.FieldRefs != null && list.FieldRefs.Count() > 0)
                        {
                            WriteText("**Fields:**", detailsWriter);
                        }

                        foreach (var fieldRef in list.FieldRefs)
                        {
                            var fieldDisplayName = fieldRef.DisplayName;
                            var fieldName = fieldRef.Name;

                            detailsWriter.WriteLine($"| {fieldDisplayName} | {fieldRef.Required.ToString()} | {fieldName} |");
                        }
                    }
                    writer.WriteLine(detailsWriter.ToString());
                    WriteNewLine(writer);
                }
            }

        }
    }
}
