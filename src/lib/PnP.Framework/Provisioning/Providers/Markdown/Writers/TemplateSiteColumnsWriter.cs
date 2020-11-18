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

            writer.WriteLine("# Site Columns");
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

                groupDetailsWriter.WriteLine("<br/><br/>");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"### {fieldDisplayName}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**Type** - {fieldType}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**Name** - {xmlField.Attributes["Name"].Value}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**Static name** - {xmlField.Attributes["StaticName"].Value}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**Required** - {xmlField.Attributes["Required"].Value}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**ID** - {xmlField.Attributes["ID"].Value}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**Enforce Unique Values** - {xmlField.Attributes["EnforceUniqueValues"].Value}");
                groupDetailsWriter.WriteLine();

                switch (fieldType.ToString().ToLower())
                {
                    case "text":
                        TextWriter(template, groupDetailsWriter, xmlField);
                        break;
                    case "number":
                        NumberWriter(template, groupDetailsWriter, xmlField);
                        break;
                    case "choice":
                        ChoiceWriter(template, groupDetailsWriter, xmlField);
                        break;
                    case "user":
                        UserWriter(template, groupDetailsWriter, xmlField);
                        break;
                    case "datetime":
                        DateTimeWriter(template, groupDetailsWriter, xmlField);
                        break;
                    case "note":
                        NoteWriter(template, groupDetailsWriter, xmlField);
                        break;
                }
            }

            writer.Write(groupDetailsWriter.ToString());
            writer.WriteLine("<br/>");
            writer.WriteLine();

        }

        private void TextWriter(ProvisioningTemplate template, TextWriter detailWriter, XmlElement xmlField)
        {
            detailWriter.WriteLine($"**Indexed** - {xmlField.Attributes["Indexed"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Max length** - {xmlField.Attributes["MaxLength"].Value}");
            detailWriter.WriteLine();
        }

        private void NumberWriter(ProvisioningTemplate template, TextWriter detailWriter, XmlElement xmlField)
        {
            detailWriter.WriteLine($"**Indexed** - {xmlField.Attributes["Indexed"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Decimals** - {xmlField.Attributes["Decimals"].Value}");
            detailWriter.WriteLine();
        }
        private void ChoiceWriter(ProvisioningTemplate template, TextWriter detailWriter, XmlElement xmlField)
        {
            detailWriter.WriteLine($"**Indexed** - {xmlField.Attributes["Indexed"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Format** - {xmlField.Attributes["Format"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Allow fill in choices?** - {xmlField.Attributes["FillInChoice"].Value}");
            detailWriter.WriteLine();
            //TODO: show choices - innerxml?
            /*
            <CHOICES>
        <CHOICE>Area One</CHOICE>
        <CHOICE>Area Two</CHOICE>
        <CHOICE>Area Three</CHOICE>
      </CHOICES>
             */
        }
        private void UserWriter(ProvisioningTemplate template, TextWriter detailWriter, XmlElement xmlField)
        {
            detailWriter.WriteLine($"**List** - {xmlField.Attributes["List"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Field to show** - {xmlField.Attributes["ShowField"].Value}");
            detailWriter.WriteLine();
            //TODO: show this in a nicer way
            detailWriter.WriteLine($"**User Selection Mode** - {xmlField.Attributes["UserSelectionMode"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**User Selection Scope** - {xmlField.Attributes["UserSelectionScope"].Value}");
            detailWriter.WriteLine();

            /*
             <Field List="UserInfo" ShowField="ImnName" UserSelectionMode="PeopleOnly" UserSelectionScope="0"  />
             */
        }
        private void DateTimeWriter(ProvisioningTemplate template, TextWriter detailWriter, XmlElement xmlField)
        {
            detailWriter.WriteLine($"**Indexed** - {xmlField.Attributes["Indexed"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Format** - {xmlField.Attributes["Format"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Show in friendly display format?** - {xmlField.Attributes["FriendlyDisplayFormat"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Cal Type** - {xmlField.Attributes["CalType"].Value}");
            detailWriter.WriteLine();
            /*
             <Field Format="DateOnly" FriendlyDisplayFormat="Disabled" CalType="0">
    </Field>
             */
        }

        private void NoteWriter(ProvisioningTemplate template, TextWriter detailWriter, XmlElement xmlField)
        {
            detailWriter.WriteLine($"**Indexed** - {xmlField.Attributes["Indexed"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Number of lines** - {xmlField.Attributes["NumLines"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Rich text?** - {xmlField.Attributes["RichText"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Rich text mode** - {xmlField.Attributes["RichTextMode"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Isolate styles?** - {xmlField.Attributes["IsolateStyles"].Value}");
            detailWriter.WriteLine();
            detailWriter.WriteLine($"**Sortable?** - {xmlField.Attributes["Sortable"].Value}");
            detailWriter.WriteLine();
        }
    }
}
