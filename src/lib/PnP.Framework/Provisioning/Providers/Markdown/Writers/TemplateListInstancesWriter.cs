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

                        using (TextWriter viewDetailsWriter = new StringWriter())
                        {

                            var xmlViewFields = from f in list.Views
                                                select XElement.Parse(f.SchemaXml);

                            foreach (var xmlField in xmlViewFields)
                            {
                                var viewDisplayName = xmlField.Attribute("DisplayName").Value;
                                //var viewType = xmlField.Attributes["Type"].Value;
                                var viewName = xmlField.Attribute("Name").Value;

                                detailsWriter.WriteLine($"| {viewDisplayName} | TBC | {viewName} |");

                                WriteHeader(viewDisplayName, 4, viewDetailsWriter);
                                WriteNewLine(viewDetailsWriter);

                                WriteAttributeField("Url", "View Url", viewDetailsWriter, xmlField);
                                WriteText("**Fields:**", viewDetailsWriter);


                                foreach (var fieldNode in xmlField.Descendants("ViewFields").Descendants("FieldRef"))
                                {
                                    WriteText($"- {GetAttributeValue("Name", fieldNode)}", viewDetailsWriter);
                                }
                                // foreach (XmlElement fieldNode in xmlField.SelectNodes("//ViewFields//FieldRef"))
                                // {
                                //     WriteText($"- {GetAttributeValue("Name", fieldNode)}", viewDetailsWriter);
                                // }
                            }
                            detailsWriter.WriteLine(viewDetailsWriter.ToString());
                            WriteNewLine(viewDetailsWriter);

                            WriteText("NB: Currently the documentation assumes you are using Content Types so it just shows the field refs", viewDetailsWriter);
                        }

                        if ((list.FieldRefs != null && list.FieldRefs.Count > 0) || (list.Fields != null && list.Fields.Count >0))
                        {
                            WriteHeader("Fields", 3, detailsWriter);
                            detailsWriter.WriteLine($"| Field Name | Display Name | Site Column*/List Column | Type | Required? | Hidden? | Max Length | Read Only? | Unique Values? |");
                            detailsWriter.WriteLine($"| :---------- | :------------: | :----------------------: | :----: | :--------: | :------: | :----------: | :---------: | :-------------: |");

                            foreach (var fieldRef in list.FieldRefs)
                            {
                                var fieldDisplayName = fieldRef.DisplayName;
                                var fieldName = fieldRef.Name;
                                var fieldRequired = fieldRef.Required.ToString();

                                detailsWriter.WriteLine($"{fieldName} | {fieldDisplayName} | Site Column* |  | {fieldRequired} |  |  |  |  |");

                            }

                            foreach (var field in list.Fields)
                            {
                                if (field.SchemaXml != null)
                                {
                                    var xmlField = XElement.Parse(field.SchemaXml);
                                    var fieldName = xmlField.Attribute("Name").Value;
                                    var fieldDisplayName = xmlField.Attribute("DisplayName").Value;
                                    var fieldType = xmlField.Attribute("Type").Value;
                                    var fieldRequired = xmlField.Attribute("Required").Value;

                                    // These may or may not be set on the XML node
                                    var fieldHidden = xmlField.Attribute("Hidden") != null ? xmlField.Attribute("Hidden").Value: "";
                                    var fieldMaxLength = xmlField.Attribute("MaxLength") != null ? xmlField.Attribute("MaxLength").Value: "";
                                    var fieldReadOnly = xmlField.Attribute("ReadOnly") != null ? xmlField.Attribute("ReadOnly").Value: "";
                                    var fieldUniqueValues = xmlField.Attribute("EnforceUniqueValues") != null ? xmlField.Attribute("EnforceUniqueValues").Value: "";

                                    detailsWriter.WriteLine($"{fieldName} | {fieldDisplayName} | List Column | {fieldType} | {fieldRequired} | {fieldHidden} | {fieldMaxLength} | {fieldReadOnly} | {fieldUniqueValues} |");
                                }

                            }

                            if (list.FieldRefs != null && list.FieldRefs.Count > 0)
                            {
                                detailsWriter.WriteLine("\n \\* For site column information, refer to Fields section at site-level.", detailsWriter);
                            }
                        }
                        
                    }
                    writer.WriteLine(detailsWriter.ToString());
                    WriteNewLine(writer);
                }
            }

        }
    }
}
