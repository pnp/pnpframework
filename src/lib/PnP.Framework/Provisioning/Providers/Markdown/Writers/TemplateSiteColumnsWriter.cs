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

            writer.WriteLine("# Site Columns");
            writer.WriteLine();
            writer.WriteLine("The table below lists the columns with their display name to help eyeball from the list of columns in Site Settings . When creating, ensure you use the Internal Name documented in the sections below.");
            writer.WriteLine();

            var xmlFields = from f in template.SiteFields
                            orderby XElement.Parse(f.SchemaXml).ToXmlElement().Attributes["Group"].Value
                            orderby XElement.Parse(f.SchemaXml).ToXmlElement().Attributes["Name"].Value
                            select XElement.Parse(f.SchemaXml).ToXmlElement();


            writer.WriteLine("| Name | Type |");
            writer.WriteLine("| :------------- | :----------: |");
            TextWriter groupDetailsWriter = new StringWriter();
            string currentGroup = "";

            foreach (var xmlField in xmlFields)
            {
                var fieldDisplayName = GetAttributeValue("DisplayName", xmlField);
                var fieldType = GetAttributeValue("Type", xmlField);
                var fieldGroup = GetAttributeValue("Group", xmlField);

                if (currentGroup != fieldGroup)
                {
                    WriteNewLine(groupDetailsWriter);
                    WriteHeader(fieldGroup, 3, groupDetailsWriter);
                    currentGroup = fieldGroup;
                }

                writer.WriteLine($"|  {fieldDisplayName} | {fieldType}   |");

                WriteNewLine(groupDetailsWriter);
                WriteNewLine(groupDetailsWriter);
                WriteHeader(fieldDisplayName, 3, groupDetailsWriter);
                
                groupDetailsWriter.WriteLine($"{GetSiteColumnTypeNameFromTemplateCode(GetAttributeValue("Type", xmlField))}");
                groupDetailsWriter.WriteLine();
                WriteAttributeField("StaticName", "Internal name", groupDetailsWriter, xmlField);
                WriteAttributeField("Description", "Description", groupDetailsWriter, xmlField);
                WriteYesNoAttributeField("Required", "Require that this column contains information", groupDetailsWriter, xmlField);
                WriteYesNoAttributeField("EnforceUniqueValues", "Enforce Unique Values", groupDetailsWriter, xmlField);

                switch (fieldType.ToString().ToLower())
                {
                    case "text":
                        TextWriter(groupDetailsWriter, xmlField);
                        break;
                    case "number":
                        NumberWriter(groupDetailsWriter, xmlField);
                        break;
                    case "choice":
                        ChoiceWriter(groupDetailsWriter, xmlField);
                        break;
                    case "user":
                        UserWriter(groupDetailsWriter, xmlField);
                        break;
                    case "datetime":
                        DateTimeWriter(groupDetailsWriter, xmlField);
                        break;
                    case "note":
                        NoteWriter(groupDetailsWriter, xmlField);
                        break;
                    //TODO: add calculated column example
                }

                WriteAttributeField("CustomFormatter", "Custom Formatting", groupDetailsWriter, xmlField);
            }

            
            writer.Write(groupDetailsWriter.ToString());
            WriteNewLine(writer);

        }

        private void TextWriter(TextWriter detailWriter, XmlElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("MaxLength", "Maximum number of characters", detailWriter, xmlField);
            if (xmlField.SelectSingleNode("Default") != null)
            {
                WriteTextField(xmlField.SelectSingleNode("Default").Value, "Default value", detailWriter);
            }
        }

        private void NumberWriter(TextWriter detailWriter, XmlElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("Decimals", "Number of decimal places", detailWriter, xmlField);
            WriteAttributeField("Min", "Min", detailWriter, xmlField);
            WriteAttributeField("Max", "Max", detailWriter, xmlField);
            if (xmlField.SelectSingleNode("Default") != null) {
                WriteTextField(xmlField.SelectSingleNode("Default").Value, "Default value", detailWriter);
            }
        }
        private void ChoiceWriter(TextWriter detailWriter, XmlElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);

            detailWriter.WriteLine("Choices:");
            detailWriter.WriteLine();
            if (xmlField.SelectSingleNode("CHOICES") != null)
            {
                foreach (XmlNode choice in xmlField.SelectSingleNode("CHOICES").ChildNodes)
                {
                    detailWriter.WriteLine("- " + choice.InnerText);
                    detailWriter.WriteLine();
                }
            }
            WriteNewLine(detailWriter);
            WriteAttributeField("Display choices using", "Format", detailWriter, xmlField);
            WriteYesNoAttributeField("FillInChoice", "Allow fill in choices?", detailWriter, xmlField);

            if (xmlField.SelectSingleNode("Default") != null)
            {
                WriteTextField(xmlField.SelectSingleNode("Default").Value, "Default value", detailWriter);
            }
        }
        private void UserWriter(TextWriter detailWriter, XmlElement xmlField)
        {
            WriteYesNoAttributeField("Mult", "Allow multiple selections?", detailWriter, xmlField);
            WriteAttributeField("Allow selection of", "UserSelectionMode", detailWriter, xmlField);

            var userSelectionScopeText = GetAttributeValue("UserSelectionScope", xmlField);
            switch(userSelectionScopeText)
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
            switch(showFieldTextText)
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
        private void DateTimeWriter(TextWriter detailWriter, XmlElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("Format", "Date and Time Format", detailWriter, xmlField);
            // DateOnly or Datetime
            WriteAttributeField("FriendlyDisplayFormat", "Display Format", detailWriter, xmlField);
            // CalType
            if (xmlField.SelectSingleNode("Default") != null)
            {
                WriteTextField(xmlField.SelectSingleNode("Default").Value, "Default value", detailWriter);
            }
        }

        private void NoteWriter(TextWriter detailWriter, XmlElement xmlField)
        {
            WriteAttributeField("Indexed", "Indexed", detailWriter, xmlField);
            WriteAttributeField("UnlimitedLengthInDocumentLibrary", "Allow unlimited length in document libraries", detailWriter, xmlField);
            WriteAttributeField("NumLines", "Number of lines for editing", detailWriter, xmlField);
            string richTextValue = "Plain text";
            if (GetAttributeValue("RichText", xmlField).ToUpper()=="TRUE")
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
