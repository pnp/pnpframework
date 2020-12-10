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
    [TemplateSchemaWriter(WriterSequence = 1040,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateSiteColumnsWriter : PnPBaseSchemaWriter<Field>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            //TODO: Handle null values and add write line after each value for a new line.
            if (template.SiteFields != null && template.SiteFields.Any())
            {
                writer.WriteLine("# Site Columns");
                writer.WriteLine();
                writer.WriteLine("The table below lists the columns with their display name to help eyeball from the list of columns in Site Settings . When creating, ensure you use the Internal Name documented in the sections below.");
                writer.WriteLine();

                var xmlFields = template.SiteFields
                    .OrderBy(f =>
                    {
                        if (!string.IsNullOrEmpty(f.SchemaXml))
                        {
                            var element = XElement.Parse(f.SchemaXml);
                            if (element.Attribute("Name") != null)
                            {
                                return element.Attribute("Name").Value;
                            }
                        }
                        return null;
                    })
                    .Select(f => XElement.Parse(f.SchemaXml));

                var fieldNames = from f in template.SiteFields select GetAttributeValue("Name", XElement.Parse(f.SchemaXml));

                var maxNameLength = fieldNames.Aggregate("", (max, cur) => max.Length > cur.Length ? max : cur).Length;
                writer.WriteLine("| Name | DisplayName | Type | ID | Group");
                writer.WriteLine("| :------------- | :----------- | :----------- | :---------- | :---------- |");
                using (TextWriter groupDetailsWriter = new StringWriter())
                {
                    foreach (var xmlField in xmlFields)
                    {
                        var fieldName = GetAttributeValue("Name", xmlField);
                        var fieldDisplayName = GetAttributeValue("DisplayName", xmlField);
                        var fieldType = GetAttributeValue("Type", xmlField);
                        var fieldGroup = GetAttributeValue("Group", xmlField);
                        var fieldId = GetAttributeValue("ID", xmlField);

                        writer.WriteLine($"| {fieldName.PadRight(maxNameLength) } | {fieldDisplayName} | {fieldType} | {fieldId } | {fieldGroup} |");

                        WriteNewLine(groupDetailsWriter);
                        WriteNewLine(groupDetailsWriter);
                        WriteHeader(fieldName, 3, groupDetailsWriter);

                        groupDetailsWriter.WriteLine($"{GetSiteColumnTypeNameFromTemplateCode(GetAttributeValue("Type", xmlField))}");
                        groupDetailsWriter.WriteLine();
                        WriteAttributeField("Name", "Internal name", groupDetailsWriter, xmlField);
                        WriteAttributeField("DisplayName", "Display name", groupDetailsWriter, xmlField);
                        WriteAttributeField("StaticName", "Static name", groupDetailsWriter, xmlField);
                        WriteAttributeField("Group", "Group", groupDetailsWriter, xmlField);
                        WriteAttributeField("Description", "Description", groupDetailsWriter, xmlField);
                        WriteYesNoAttributeField("Required", "Require that this column contains information", groupDetailsWriter, xmlField);
                        WriteYesNoAttributeField("EnforceUniqueValues", "Enforce Unique Values", groupDetailsWriter, xmlField);

                        switch (fieldType.ToString().ToLower())
                        {
                            case "text":
                                TextFieldWriter(groupDetailsWriter, xmlField);
                                break;
                            case "number":
                                NumberFieldWriter(groupDetailsWriter, xmlField);
                                break;
                            case "choice":
                                ChoiceFieldWriter(groupDetailsWriter, xmlField);
                                break;
                            case "user":
                                UserFieldWriter(groupDetailsWriter, xmlField);
                                break;
                            case "datetime":
                                DateTimeFieldWriter(groupDetailsWriter, xmlField);
                                break;
                            case "note":
                                NoteFieldWriter(groupDetailsWriter, xmlField);
                                break;
                                //TODO: add calculated column example
                        }

                        WriteAttributeField("CustomFormatter", "Custom Formatting", groupDetailsWriter, xmlField);
                    }


                    writer.Write(groupDetailsWriter.ToString());
                }
                WriteNewLine(writer);
            }
        }

        private string GetFieldOrderByValue(Field f, string attribute)
        {
            if (!string.IsNullOrEmpty(f.SchemaXml))
            {
                var element = XElement.Parse(f.SchemaXml);
                if (element.Attribute(attribute) != null)
                {
                    return element.Attribute(attribute).Value;
                }
            }
            return null;
        }

        private void TextFieldWriter(TextWriter detailWriter, XElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("MaxLength", "Maximum number of characters", detailWriter, xmlField);
            var defaultNode = xmlField.Descendants("Default").FirstOrDefault();
            if (defaultNode != null)
            {
                WriteTextField(defaultNode.Value, "Default value", detailWriter);
            }
        }

        private void NumberFieldWriter(TextWriter detailWriter, XElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("Decimals", "Number of decimal places", detailWriter, xmlField);
            WriteAttributeField("Min", "Min", detailWriter, xmlField);
            WriteAttributeField("Max", "Max", detailWriter, xmlField);
            var defaultNode = xmlField.Descendants("Default").FirstOrDefault();
            if (defaultNode != null)
            {
                WriteTextField(defaultNode.Value, "Default value", detailWriter);
            }
        }
        private void ChoiceFieldWriter(TextWriter detailWriter, XElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);

            if (xmlField.Descendants("CHOICES") != null && xmlField.Descendants("CHOICES").Descendants() != null)
            {
                detailWriter.WriteLine("**Choices:**");
                detailWriter.WriteLine();
                if (xmlField.Descendants("CHOICES") != null)
                {
                    foreach (var choiceNode in xmlField.Descendants("CHOICES").Descendants())
                    {
                        detailWriter.WriteLine("- " + choiceNode.Value);
                        detailWriter.WriteLine();
                    }
                }
            }

            WriteNewLine(detailWriter);
            WriteAttributeField("Format", "Display choices using", detailWriter, xmlField);
            WriteYesNoAttributeField("FillInChoice", "Allow fill in choices?", detailWriter, xmlField);
            var defaultNode = xmlField.Descendants("Default").FirstOrDefault();
            if (defaultNode != null)
            {
                WriteTextField(defaultNode.Value, "Default value", detailWriter);
            }
        }
        private void UserFieldWriter(TextWriter detailWriter, XElement xmlField)
        {
            WriteYesNoAttributeField("Mult", "Allow multiple selections?", detailWriter, xmlField);
            WriteAttributeField("UserSelectionMode", "Allow selection of", detailWriter, xmlField);

            var userSelectionScopeText = GetAttributeValue("UserSelectionScope", xmlField);
            switch (userSelectionScopeText)
            {
                case "0":
                    userSelectionScopeText = "All Users";
                    break;
                case "1":
                    userSelectionScopeText = "SharePoint Group";
                    break;
            }
            detailWriter.WriteLine($"Choose from - {userSelectionScopeText}");
            detailWriter.WriteLine();

            var showFieldTextText = GetAttributeValue("UserSelectionScope", xmlField);
            switch (showFieldTextText)
            {
                case "Title":
                    showFieldTextText = "User";
                    break;
                case "ComplianceAssetId":
                    showFieldTextText = "Compliance Asset Id";
                    break;
                case "Name":
                    showFieldTextText = "Account";
                    break;
                case "Email":
                    showFieldTextText = "Work Email";
                    break;
                case "OtherMail":
                    showFieldTextText = "OtherMail";
                    break;
                case "UserExpiration":
                    showFieldTextText = "UserExpiration";
                    break;
                case "UserLastDeletionTime":
                    showFieldTextText = "User Last Deletion Time";
                    break;
                case "MobilePhone":
                    showFieldTextText = "Mobile phone";
                    break;
                case "SipAddress":
                    showFieldTextText = "SIP Address";
                    break;
                case "Department":
                    showFieldTextText = "Department";
                    break;
                case "JobTitle":
                    showFieldTextText = "Title";
                    break;
                case "FirstName":
                    showFieldTextText = "First name";
                    break;
                case "LastName":
                    showFieldTextText = "Last name";
                    break;
                case "WorkPhone":
                    showFieldTextText = "Work phone";
                    break;
                case "UserName":
                    showFieldTextText = "User name";
                    break;
                case "Office":
                    showFieldTextText = "Office";
                    break;
                case "ID":
                    showFieldTextText = "ID";
                    break;
                case "Modified":
                    showFieldTextText = "Modified";
                    break;
                case "Created":
                    showFieldTextText = "Created";
                    break;
                case "ImnName":
                    showFieldTextText = "Name (with prescence)";
                    break;
                case "PictureOnly_Size_36px":
                    showFieldTextText = "Picture Only (36x36)";
                    break;
                case "PictureOnly_Size_48px":
                    showFieldTextText = "Picture Only (48x48)";
                    break;
                case "PictureOnly_Size_72px":
                    showFieldTextText = "Picture Only (72x72)";
                    break;
                case "NameWithPictureAndDetails":
                    showFieldTextText = "Name (with picture and details)";
                    break;
                case "ContentType":
                    showFieldTextText = "Content Type";
                    break;
            }
            detailWriter.WriteLine($"**Show field** - {showFieldTextText}");
            detailWriter.WriteLine();
        }
        private void DateTimeFieldWriter(TextWriter detailWriter, XElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("Format", "Date and Time Format", detailWriter, xmlField);
            // DateOnly or Datetime
            WriteAttributeField("FriendlyDisplayFormat", "Display Format", detailWriter, xmlField);
            // CalType
            var defaultNode = xmlField.Descendants("Default").FirstOrDefault();
            if (defaultNode != null)
            {
                WriteTextField(defaultNode.Value, "Default value", detailWriter);
            }
        }

        private void NoteFieldWriter(TextWriter detailWriter, XElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("UnlimitedLengthInDocumentLibrary", "Allow unlimited length in document libraries", detailWriter, xmlField);
            WriteAttributeField("NumLines", "Number of lines for editing", detailWriter, xmlField);
            string richTextValue = "Plain text";
            if (GetAttributeValue("RichText", xmlField).ToUpper() == "TRUE")
            {
                switch (GetAttributeValue("RichTextMode", xmlField))
                {
                    case "Compatible":
                        richTextValue = "Rich text(Bold, italics, text alignment, hyperlinks)";
                        break;
                    case "FullHtml":
                        richTextValue = "Enhanced rich text(Rich text with pictures, tables, and hyperlinks)";
                        break;
                }
            }
            WriteTextField(richTextValue, "Specify the type of text to allow", detailWriter);

            WriteYesNoAttributeField("AppendOnly", "Append Changes to Existing Text", detailWriter, xmlField);

        }
    }
}
