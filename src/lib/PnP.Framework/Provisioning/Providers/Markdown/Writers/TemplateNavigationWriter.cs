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
    [TemplateSchemaWriter(WriterSequence = 1030,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateNavigationWriter : PnPBaseSchemaWriter<Navigation>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {
            if (template.Navigation != null)
            {
                WriteHeader("Navigation", 1, writer);
                WriteTextField(template.Navigation.AddNewPagesToNavigation.ToString(), "Add new pages to navigation", writer);
                WriteTextField(template.Navigation.EnableTreeView.ToString(), "Treeview enabled", writer);
                WriteTextField(template.Navigation.CreateFriendlyUrlsForNewPages.ToString(), "Create friendly urls for new pages", writer);


                if (template.Navigation.GlobalNavigation != null)
                {
                    WriteHeader("Global Navigation", 2, writer);
                    WriteTextField(template.Navigation.GlobalNavigation.NavigationType.ToString(), "Navigation Type", writer);
                }
            }
        }
    }
}
