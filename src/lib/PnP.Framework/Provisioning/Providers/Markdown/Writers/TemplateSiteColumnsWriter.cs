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
    internal class TemplateSiteColumnsWriter : IPnPSchemaWriter
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        protected LambdaExpression CreateSelectorLambda(Type targetType, String propertyName)
        {
            return (Expression.Lambda(
                Expression.Convert(
                    Expression.MakeMemberAccess(
                        Expression.Parameter(targetType, "i"),
                        targetType.GetProperty(propertyName,
                            System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.Public)),
                    typeof(object)),
                ParameterExpression.Parameter(targetType, "i")));
        }

        public void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            //TODO: Handle null values and add write line after each value for a new line.

            writer.WriteLine("## Site Columns");
            writer.WriteLine();
            var xmlFields = from f in template.SiteFields
                            //TODO: sort by group - orderby XElement.Parse(f.SchemaXml).ToXmlElement().Attributes["Group"]
                            select XElement.Parse(f.SchemaXml).ToXmlElement();


            writer.WriteLine("| Column | Type |");
            writer.WriteLine("| :------------- | :----------: |");
            TextWriter groupDetailsWriter = new StringWriter();

            foreach (var xmlField in xmlFields)
            {
                var fieldDisplayName = xmlField.Attributes["DisplayName"].Value;
                var fieldType = xmlField.Attributes["Type"].Value;
                var fieldGroup = xmlField.Attributes["Group"].Value;

                writer.WriteLine($"|  {fieldDisplayName} | {fieldType}   |");

                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"#### {fieldDisplayName}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"Type - {fieldType}");
                groupDetailsWriter.WriteLine($"Name - {xmlField.Attributes["Name"].Value}");
                groupDetailsWriter.WriteLine($"Static name - {xmlField.Attributes["StaticName"].Value}");
                groupDetailsWriter.WriteLine($"Required - {xmlField.Attributes["Required"].Value}");
                groupDetailsWriter.WriteLine($"ID - {xmlField.Attributes["ID"].Value}");
                groupDetailsWriter.WriteLine($"Enforce Unique Values - {xmlField.Attributes["EnforceUniqueValues"].Value}");
                
                switch (fieldType.ToString().ToLower())
                {
                    case "text":
                        TextWriter(template, writer, xmlField);
                        break;
                    case "number":
                        NumberWriter(template, writer, xmlField);
                        break;
                    case "choice":
                        ChoiceWriter(template, writer, xmlField);
                        break;
                    case "user":
                        UserWriter(template, writer, xmlField);
                        break;
                    case "datetime":
                        DateTimeWriter(template, writer, xmlField);
                        break;
                    case "note":
                        NoteWriter(template, writer, xmlField);
                        break;
                }
            }

            writer.WriteLine(groupDetailsWriter.ToString());

        }

        private void TextWriter(ProvisioningTemplate template, TextWriter writer, XmlElement xmlField)
        {
            writer.WriteLine($"Indexed - {xmlField.Attributes["Indexed"].Value}");
            writer.WriteLine($"Max length - {xmlField.Attributes["MaxLength"].Value}");
        }

        private void NumberWriter(ProvisioningTemplate template, TextWriter writer, XmlElement xmlField)
        {
            writer.WriteLine($"Indexed - {xmlField.Attributes["Indexed"].Value}");
            writer.WriteLine($"Decimals - {xmlField.Attributes["Decimals"].Value}");
        }
        private void ChoiceWriter(ProvisioningTemplate template, TextWriter writer, XmlElement xmlField)
        {
            writer.WriteLine($"Indexed - {xmlField.Attributes["Indexed"].Value}");
            writer.WriteLine($"Format - {xmlField.Attributes["Format"].Value}");
            writer.WriteLine($"Allow fill in choices? - {xmlField.Attributes["FillInChoice"].Value}");
            //TODO: show choices - innerxml?
            /*
            <CHOICES>
        <CHOICE>Area One</CHOICE>
        <CHOICE>Area Two</CHOICE>
        <CHOICE>Area Three</CHOICE>
      </CHOICES>
             */
        }
        private void UserWriter(ProvisioningTemplate template, TextWriter writer, XmlElement xmlField)
        {
            writer.WriteLine($"List - {xmlField.Attributes["List"].Value}");
            writer.WriteLine($"Field to show - {xmlField.Attributes["ShowField"].Value}");
            //TODO: show this in a nicer way
            writer.WriteLine($"User Selection Mode - {xmlField.Attributes["UserSelectionMode"].Value}");
            writer.WriteLine($"User Selection Scope - {xmlField.Attributes["UserSelectionScope"].Value}");

            /*
             <Field List="UserInfo" ShowField="ImnName" UserSelectionMode="PeopleOnly" UserSelectionScope="0"  />
             */
        }
        private void DateTimeWriter(ProvisioningTemplate template, TextWriter writer, XmlElement xmlField)
        {
            writer.WriteLine($"Indexed - {xmlField.Attributes["Indexed"].Value}");
            writer.WriteLine($"Format - {xmlField.Attributes["Format"].Value}");
            writer.WriteLine($"Show in friendly display format? - {xmlField.Attributes["FriendlyDisplayFormat"].Value}");
            writer.WriteLine($"Cal Type - {xmlField.Attributes["CalType"].Value}");
            /*
             <Field Format="DateOnly" FriendlyDisplayFormat="Disabled" CalType="0">
    </Field>
             */
        }

        private void NoteWriter(ProvisioningTemplate template, TextWriter writer, XmlElement xmlField)
        {
            writer.WriteLine($"Indexed - {xmlField.Attributes["Indexed"].Value}");
            writer.WriteLine($"Number of lines - {xmlField.Attributes["NumLines"].Value}");
            writer.WriteLine($"Rich text? - {xmlField.Attributes["RichText"].Value}");
            writer.WriteLine($"Rich text mode - {xmlField.Attributes["RichTextMode"].Value}");
            writer.WriteLine($"Isolate styles? - {xmlField.Attributes["IsolateStyles"].Value}");
            writer.WriteLine($"Sortable? - {xmlField.Attributes["Sortable"].Value}");
        }
    }
}
