using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Markdown;
using PnP.Framework.Provisioning.Providers.Xml;
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
    [TemplateSchemaWriter(WriterSequence = 1060,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateListInstancesWriter : IPnPSchemaWriter
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
            writer.WriteLine("## Lists");
            writer.WriteLine();
            writer.WriteLine("| Name | Description |");
            writer.WriteLine("| :------------- | :----------: |");
            TextWriter groupDetailsWriter = new StringWriter();

            string currentGroup = "";

            foreach (var list in template.Lists)
            {

                writer.WriteLine($"|  {list.Title} | {list.Description}   |");

                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"#### {list.Title}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"Description - {list.Description}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"Template type - {list.TemplateType}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"Url - {list.Url}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine($"Enable versioning - {list.EnableVersioning}");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine("Views:");
                groupDetailsWriter.WriteLine();
                groupDetailsWriter.WriteLine("| Display Name |  Default?  |   Name    |");
                groupDetailsWriter.WriteLine("| :------------- | :----------: | :----------: |");

                var xmlFields = from f in list.Views
                                select XElement.Parse(f.SchemaXml).ToXmlElement();

                foreach (var xmlField in xmlFields)
                {
                    var viewDisplayName = xmlField.Attributes["DisplayName"].Value;
                    //var viewType = xmlField.Attributes["Type"].Value;
                    var viewName = xmlField.Attributes["Name"].Value;

                    groupDetailsWriter.WriteLine($"| {viewDisplayName} | TBC | {viewName} |");
                    groupDetailsWriter.WriteLine();
                }
            }
            writer.WriteLine(groupDetailsWriter.ToString());
            writer.WriteLine();

            /*
             * <pnp:ListInstance Title="Style Library" Description="Use the style library to store style sheets, such as CSS or XSL files. The style sheets in this gallery can be used by this site or any of its subsites." 
             * DocumentTemplate="" TemplateType="101" Url="tyle Library" EnableVersioning="true" MinorVersionLimit="0" MaxVersionLimit="500" DraftVersionVisibility="0" TemplateFeatureID="00bfea71-e717-4e80-aa17-d0c71b360101" EnableAttachments="false" DefaultDisplayFormUrl="{site}/Style Library/Forms/DispForm.aspx" DefaultEditFormUrl="{site}/Style Library/Forms/EditForm.aspx" DefaultNewFormUrl="{site}/Style Library/Forms/Upload.aspx" ImageUrl="/_layouts/15/images/itdl.png?rev=47" IrmExpire="false" IrmReject="false" IsApplicationList="false" ValidationFormula="" ValidationMessage="">
          <pnp:ContentTypeBindings>
            <pnp:ContentTypeBinding ContentTypeID="0x0101" Default="true" />
            <pnp:ContentTypeBinding ContentTypeID="0x0120" />
          </pnp:ContentTypeBindings>
          <pnp:Views>
            <View Name="{73E8B349-5F9E-458B-BC1F-7BC77E7AAE8E}" DefaultView="TRUE" MobileView="TRUE" MobileDefaultView="TRUE" Type="HTML" DisplayName="All Documents" Url="{site}/Style Library/Forms/AllItems.aspx" Level="1" BaseViewID="1" ContentTypeID="0x" ImageUrl="/_layouts/15/images/dlicon.png?rev=47">
              <Query>
                <OrderBy>
                  <FieldRef Name="FileLeafRef" />
                </OrderBy>
              </Query>
              <ViewFields>
                <FieldRef Name="DocIcon" />
                <FieldRef Name="LinkFilename" />
                <FieldRef Name="Modified" />
                <FieldRef Name="Editor" />
              </ViewFields>
              <RowLimit Paged="TRUE">30</RowLimit>
              <JSLink>clienttemplates.js</JSLink>
            </View>
          </pnp:Views>
          <pnp:FieldRefs>
            <pnp:FieldRef ID="d307dff3-340f-44a2-9f4b-fbfe1ba07459" Name="_CommentCount" DisplayName="Comment count" />
            <pnp:FieldRef ID="db8d9d6d-dc9a-4fbd-85f3-4a753bfdc58c" Name="_LikeCount" DisplayName="Like count" />
          </pnp:FieldRefs>
        </pnp:ListInstance>
            */
        }
    }
}
