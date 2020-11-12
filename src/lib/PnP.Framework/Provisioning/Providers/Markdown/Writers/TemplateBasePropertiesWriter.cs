using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Xml.Linq;

namespace PnP.Framework.Provisioning.Providers.Markdown.Writers
{
    /// <summary>
    /// Class to write out the markdown for the base properties
    /// </summary>
    internal class TemplateBasePropertiesWriter : IPnPSchemaWriter
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
            writer.WriteLine("## Web Info");
            writer.WriteLine();
            writer.WriteLine($"Landing page - {template.WebSettings.WelcomePage}");
            writer.WriteLine();
            writer.WriteLine($"Request Access Email - {template.WebSettings.RequestAccessEmail}");
            writer.WriteLine();
            bool commentsEnabled = template.WebSettings.CommentsOnSitePagesDisabled;
            string commentsEnabledText = "Comments on pages enabled";
            if (commentsEnabled)
            {
                commentsEnabledText = "Comments on pages disabled";
            }
            writer.WriteLine(commentsEnabledText);
            writer.WriteLine();

            string searchScopeText = "Default search scope";
            if (template.WebSettings.SearchScope != SearchScopes.DefaultScope)
            {
                searchScopeText = $"Search scope - {template.WebSettings.SearchScope.ToString()}";
            }
            writer.WriteLine(searchScopeText);
            writer.WriteLine();
        }
    }
}
