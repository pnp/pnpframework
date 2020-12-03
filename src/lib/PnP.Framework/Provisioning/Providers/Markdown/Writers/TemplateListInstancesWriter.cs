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
            writer.WriteLine("<br/>");
            writer.WriteLine();
            writer.WriteLine("# Lists");
            writer.WriteLine();
            writer.WriteLine("| Name | Description |");
            writer.WriteLine("| :------------- | :----------: |");
            TextWriter groupDetailsWriter = new StringWriter();

            foreach (var list in template.Lists)
            {

                writer.WriteLine($"|  {list.Title} | {list.Description}   |");
                WriteNewLine(groupDetailsWriter);
                WriteHeader(list.Title, 2, groupDetailsWriter);
                WriteNewLine(groupDetailsWriter);
                WriteTextField(list.Description, "Description", groupDetailsWriter);
                WriteTextField(list.OnQuickLaunch.ToString(), "Show on quick launch?", groupDetailsWriter);
                WriteTextField(GetListTemplateNameFromTemplateCode(list.TemplateType.ToString()), "List template", groupDetailsWriter);
                WriteTextField(list.Url, "URL", groupDetailsWriter);
                WriteNewLine(groupDetailsWriter);

                WriteHeader("Versioning", 3, groupDetailsWriter);                
                WriteTextField("TBC", "Require content approval?", groupDetailsWriter);
                WriteTextField(list.EnableVersioning.ToString(), "Create a version each time you edit an item in this list ?", groupDetailsWriter);
                WriteTextField("TBC", "Draft item security", groupDetailsWriter);
                WriteNewLine(groupDetailsWriter);

                WriteHeader("Advanced settings", 3, groupDetailsWriter);
                WriteTextField(list.ContentTypesEnabled.ToString(), "Allow management of content types?", groupDetailsWriter);
                WriteNewLine(groupDetailsWriter);

                WriteHeader("Content Types", 3, groupDetailsWriter);
                foreach (var binding in list.ContentTypeBindings)
                {
                    WriteText($"- {binding.ContentTypeId}", groupDetailsWriter);
                }
                WriteNewLine(groupDetailsWriter);

                WriteHeader("Views", 3, groupDetailsWriter);
                groupDetailsWriter.WriteLine("| Display Name |  Default?  |   Name    |");
                groupDetailsWriter.WriteLine("| :------------- | :----------: | :----------: |");

                TextWriter viewDetailsWriter = new StringWriter();

                var xmlViewFields = from f in list.Views
                                select XElement.Parse(f.SchemaXml).ToXmlElement();

                foreach (var xmlField in xmlViewFields)
                {
                    var viewDisplayName = xmlField.Attributes["DisplayName"].Value;
                    //var viewType = xmlField.Attributes["Type"].Value;
                    var viewName = xmlField.Attributes["Name"].Value;

                    groupDetailsWriter.WriteLine($"| {viewDisplayName} | TBC | {viewName} |");

                    WriteHeader(viewDisplayName, 4, viewDetailsWriter);
                    WriteNewLine(viewDetailsWriter);
                   
                    WriteAttributeField("Url", "View Url", viewDetailsWriter, xmlField);
                    WriteText("**Fields:**", viewDetailsWriter);

                    foreach (XmlElement fieldNode in xmlField.SelectNodes("//ViewFields//FieldRef"))
                    {
                        WriteText($"- {GetAttributeValue("Name", fieldNode)}", viewDetailsWriter);
                    }
                }
                groupDetailsWriter.WriteLine(viewDetailsWriter.ToString());
                WriteNewLine(viewDetailsWriter);

                WriteText("NB: Currently the documentation assumes you are using Content Types so it just shows the field refs", viewDetailsWriter);

                if (list.FieldRefs != null && list.FieldRefs.Count() > 0)
                {
                    WriteText("**Fields:**", groupDetailsWriter);
                }

                foreach (var fieldRef in list.FieldRefs)
                {
                    var fieldDisplayName = fieldRef.DisplayName;
                    var fieldName = fieldRef.Name;

                    groupDetailsWriter.WriteLine($"| {fieldDisplayName} | {fieldRef.Required.ToString()} | {fieldName} |");
                }
            }
            writer.WriteLine(groupDetailsWriter.ToString());
            WriteNewLine(writer);

        }
    }
}
