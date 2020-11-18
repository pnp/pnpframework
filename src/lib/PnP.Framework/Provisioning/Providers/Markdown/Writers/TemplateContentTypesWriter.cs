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
    internal class TemplateContentTypesWriter : IPnPSchemaWriter
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
            writer.WriteLine("# Content Types");
            writer.WriteLine();
            writer.WriteLine("| Name | Description |");
            writer.WriteLine("| :------------- | :----------: |");
            TextWriter groupDetailsWriter = new StringWriter();

            string currentGroup = "";

            foreach (var ct in template.ContentTypes.OrderBy(o => o.Group))
            {
                if (currentGroup != ct.Group)
                {
                    groupDetailsWriter.WriteLine("<br/>");
                    groupDetailsWriter.WriteLine();
                    groupDetailsWriter.Write($"### Group - {ct.Group}");
                    groupDetailsWriter.WriteLine();
                }

                
                writer.WriteLine($"|  {ct.Name} | {ct.Description}   |");

                groupDetailsWriter.WriteLine("<br/>");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"#### {ct.Name}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**Description** - {ct.Description}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"**ID** - {ct.Id}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine("**Fields**:");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine("| Name   |     Required     | Hidden       |");
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
