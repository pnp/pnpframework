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
    [TemplateSchemaWriter(WriterSequence = 1020,
        Scope = WriterScope.ProvisioningTemplate)]
    internal class TemplateSecurityWriter : PnPBaseSchemaWriter<SiteSecurity>
    {
        public override void Writer(ProvisioningTemplate template, TextWriter writer)
        {

            using (var detailsWriter = new StringWriter())
            {
                using (var associatedGroupsWriter = new StringWriter())
                {
                    if (!string.IsNullOrEmpty(template.Security.AssociatedOwnerGroup))
                    {
                        associatedGroupsWriter.WriteLine($"|**Associated Owners Group**|{template.Security.AssociatedOwnerGroup}|");
                    }
                    if (!string.IsNullOrEmpty(template.Security.AssociatedMemberGroup))
                    {
                        associatedGroupsWriter.WriteLine($"|**Associated Members Group**|{template.Security.AssociatedMemberGroup}|");
                    }
                    if (!string.IsNullOrEmpty(template.Security.AssociatedVisitorGroup))
                    {
                        associatedGroupsWriter.WriteLine($"|**Associated Visitors Group**|{template.Security.AssociatedVisitorGroup}|");
                    }
                    var associatedGroupText = associatedGroupsWriter.ToString();
                    if (!string.IsNullOrEmpty(associatedGroupText))
                    {

                        WriteHeader("Associated Groups", 2, detailsWriter);
                        if (associatedGroupText.Contains("{groupsitetitle}"))
                        {
                            detailsWriter.WriteLine("*{groupsitetitle} will be replaced at provisioning time with the actual site title*");
                        }
                        detailsWriter.WriteLine("| Group Type | Group Name |");
                        detailsWriter.WriteLine("| --- | --- |");
                        WriteText(associatedGroupText, detailsWriter);
                    }
                }
                if (template.Security != null && template.Security.AdditionalAdministrators != null && template.Security.AdditionalAdministrators.Count > 0)
                {
                    WriteHeader("Administrators", 2, detailsWriter);
                    foreach (var admin in template.Security.AdditionalAdministrators)
                    {
                        WriteText(admin.Name, detailsWriter);
                    }
                }

                if (template.Security != null && template.Security.AdditionalAdministrators != null && template.Security.AdditionalOwners.Count > 0)
                {
                    WriteHeader("Owners", 2, detailsWriter);
                    foreach (var admin in template.Security.AdditionalOwners)
                    {
                        WriteText(admin.Name, detailsWriter);
                    }
                }

                if (template.Security != null && template.Security.AdditionalMembers != null && template.Security.AdditionalMembers.Count > 0)
                {
                    WriteHeader("Members", 2, detailsWriter);
                    foreach (var admin in template.Security.AdditionalOwners)
                    {
                        WriteText(admin.Name, detailsWriter);
                    }
                }

                if (template.Security != null && template.Security.AdditionalVisitors != null && template.Security.AdditionalVisitors.Count > 0)
                {
                    WriteHeader("Members", 2, detailsWriter);
                    foreach (var admin in template.Security.AdditionalOwners)
                    {
                        WriteText(admin.Name, detailsWriter);
                    }
                }

                var results = detailsWriter.ToString();
                if (!string.IsNullOrEmpty(results))
                {
                    writer.WriteLine("# Permissions");
                    writer.Write(results);
                }
            }
        }
    }
}
