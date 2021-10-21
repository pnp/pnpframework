using PnP.Framework.Provisioning.Model;
using System;
using System.IO;
using System.Linq.Expressions;
using System.Xml;
using System.Xml.Linq;

namespace PnP.Framework.Provisioning.Providers.Markdown
{
    /// <summary>
    /// Base class for every Schema Serializer
    /// </summary>
    internal abstract class PnPBaseSchemaWriter<TModelType> : IPnPSchemaWriter
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        public abstract void Writer(ProvisioningTemplate template, TextWriter writer);


        /// <summary>
        /// Protected method to create a Lambda Expression like: i => i.Property
        /// </summary>
        /// <param name="targetType">The Type of the .NET property to apply the Lambda Expression to</param>
        /// <param name="propertyName">The name of the property of the target object</param>
        /// <returns></returns>
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

        protected void WriteAttributeField(string attributeName, string fieldName, TextWriter fieldWriter, XElement xmlField)
        {
            var value = GetAttributeValue(attributeName, xmlField);
            if (!string.IsNullOrEmpty(value))
            {
                fieldWriter.WriteLine($"**{fieldName}** - {value}");
                fieldWriter.WriteLine();
            }
        }

        protected void WriteTextField(string fieldValue, string fieldName, TextWriter fieldWriter)
        {
            fieldWriter.WriteLine($"**{fieldName}** - {fieldValue}");
            fieldWriter.WriteLine();
        }

        protected void WriteText(string textValue, TextWriter fieldWriter)
        {
            fieldWriter.WriteLine($"{textValue}");
            fieldWriter.WriteLine();
        }

        protected void WriteYesNoAttributeField(string attributeName, string fieldName, TextWriter fieldWriter, XElement xmlField)
        {

            var fieldText = GetAttributeValue("fieldName", xmlField);
            if (fieldText.ToUpper() == "TRUE")
            {
                fieldText = "Yes";
            }
            else
            {
                fieldText = "No";
            }


            fieldWriter.WriteLine($"**{fieldName}** - {fieldText}");
            fieldWriter.WriteLine();
        }

        protected void WriteNewLine(TextWriter fieldWriter)
        {
            fieldWriter.WriteLine("<br/>");
            fieldWriter.WriteLine();
        }

        protected void WriteHeader(string headerText, int headerLevel, TextWriter fieldWriter)
        {
            var headerPrefix = "".PadLeft(headerLevel, '#');
            fieldWriter.WriteLine($"{headerPrefix} {headerText}");
            fieldWriter.WriteLine();
        }

        protected static string GetAttributeValue(string attributeName, XElement xmlField)
        {
            if (xmlField.Attribute(attributeName) != null)
            {

                return xmlField.Attribute(attributeName).Value;
            }
            return "";
        }
        protected string GetSiteColumnTypeNameFromTemplateCode(string templateCode)
        {
            switch (templateCode)
            {
                case "Text":
                    return $"**Type** - Single line of text";
                case "Note":
                    return $"**Type** - Multi-line text";
                default:
                    return $"**Type** - {templateCode}";
            }
        }

        protected string GetSiteTemplateNameFromTemplateCode(string templateCode)
        {
            switch (templateCode)
            {
                case "GROUP#0":
                    return $"**Site type** - Modern team site";
                case "SITEPAGEPUBLISHING#0":
                    return $"**Site type** - Modern communications site";
                default:
                    return $"**Site type** - Other ({templateCode})";
            }
            //TODO: map other template types but need to identify template names from somewhere
            /*
             <pnp:WebTemplate LanguageCode="1033" TemplateName="GLOBAL#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="STS#3" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="STS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="STS#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="STS#2" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="MPS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="MPS#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="MPS#2" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="MPS#3" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="MPS#4" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="CENTRALADMIN#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="WIKI#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="BLOG#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SGS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="TENANTADMIN#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="APP#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="APPCATALOG#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="ACCSRV#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="ACCSVC#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="ACCSVC#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="BDR#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="CONTENTCTR#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="TBH#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="DEV#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="EDISC#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="EDISC#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="EXPRESS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="FunSite#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="OFFILE#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="OFFILE#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="EHS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="EHS#2" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="EHS#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="OSRV#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="PPSMASite#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="BICenterSite#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="PWA#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="PWS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="REVIEWCTR#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="RedirectSite#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="POLICYCTR#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#2" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#3" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#4" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#5" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#6" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#7" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#8" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#9" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPERS#10" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSMSITE#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSTOC#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSTOPIC#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSNEWS#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="CMSPUBLISHING#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="BLANKINTERNET#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="BLANKINTERNET#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="BLANKINTERNET#2" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSNHOME#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSSITES#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSCOMMU#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSREPORTCENTER#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSPORTAL#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SRCHCEN#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="PROFILES#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="BLANKINTERNETCONTAINER#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SPSMSITEHOST#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="ENTERWIKI#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="PROJECTSITE#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="PRODUCTCATALOG#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="COMMUNITY#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="COMMUNITYPORTAL#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="GROUP#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="POINTPUBLISHINGHUB#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="POINTPUBLISHINGPERSONAL#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="POINTPUBLISHINGPERSONAL#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="POINTPUBLISHINGTOPIC#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SITEPAGEPUBLISHING#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="TEAMCHANNEL#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="TEAMCHANNEL#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SRCHCENTERLITE#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SRCHCENTERLITE#1" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="TenantAdminSpo#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="TestSite#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="visprus#0" />
  <pnp:WebTemplate LanguageCode="1033" TemplateName="SAPWorkflowSite#0" />
             */
        }

        protected string GetListTemplateNameFromTemplateCode(string templateCode)
        {

            switch (templateCode)
            {
                case "100": return "**List template** - Custom List";
                case "101": return "**List template** - Document Library";
                case "102": return "**List template** - Survey";
                case "103": return "**List template** - Links";
                case "104": return "**List template** - Announcements";
                case "105": return "**List template** - Contacts";
                case "106": return "**List template** - Calendar";
                case "107": return "**List template** - Tasks";
                case "108": return "**List template** - Discussion Board";
                case "109": return "**List template** - Picture Library";
                case "110": return "**List template** - DataSources";
                case "115": return "**List template** - Form Library";
                case "117": return "**List template** - No Code Workflows";
                case "118": return "**List template** - Custom Workflow Process";
                case "119": return "**List template** - Wiki Page Library";
                case "120": return "**List template** - CustomGrid";
                case "122": return "**List template** - No Code Public Workflows<14>";
                case "140": return "**List template** - Workflow History";
                case "150": return "**List template** - Project Tasks";
                case "600": return "**List template** - Public Workflows External List<15>";
                case "1100": return "**List template** - Issues Tracking";
                case "1230": return "**List template** - Draft Apps";
                case "3100": return "**List template** - Access App";
                case "10102": return "**List template** - Converted Forms";
                case "170": return "**List template** - Promoted Links";
                case "433": return "**List template** - Report Library";
                case "432": return "**List template** - Status List";
                default: return $"**List template** - Code not found ({templateCode})";
            }
        }
    }
}
