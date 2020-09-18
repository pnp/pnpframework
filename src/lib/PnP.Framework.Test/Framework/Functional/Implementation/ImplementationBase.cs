using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System;

namespace PnP.Framework.Tests.Framework.Functional.Implementation
{
    internal abstract class ImplementationBase
    {
        #region Apply template and read the "result"
        public TestProvisioningTemplateResult TestProvisioningTemplate(ClientContext cc, string templateName, Handlers handlersToProcess = Handlers.All, ProvisioningTemplateApplyingInformation ptai = null, ProvisioningTemplateCreationInformation ptci = null)
        {
            try
            {
                // Read the template from XML and apply it
                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");
                ProvisioningTemplate sourceTemplate = provider.GetTemplate(templateName);

                if (ptai == null)
                {
                    ptai = new ProvisioningTemplateApplyingInformation
                    {
                        HandlersToProcess = handlersToProcess
                    };
                }

                if (ptai.ProgressDelegate == null)
                {
                    ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                    {
                        Console.WriteLine("Applying template - {0}/{1} - {2}", progress, total, message);
                    };
                }

                sourceTemplate.Connector = provider.Connector;

                TokenParser sourceTokenParser = new TokenParser(cc.Web, sourceTemplate);

                cc.Web.ApplyProvisioningTemplate(sourceTemplate, ptai);

                // Read the site we applied the template to 
                if (ptci == null)
                {
                    ptci = new ProvisioningTemplateCreationInformation(cc.Web)
                    {
                        HandlersToProcess = handlersToProcess
                    };
                }

                if (ptci.ProgressDelegate == null)
                {
                    ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                    {
                        Console.WriteLine("Getting template - {0}/{1} - {2}", progress, total, message);
                    };
                }

                ProvisioningTemplate targetTemplate = cc.Web.GetProvisioningTemplate(ptci);

                return new TestProvisioningTemplateResult()
                {
                    SourceTemplate = sourceTemplate,
                    SourceTokenParser = sourceTokenParser,
                    TargetTemplate = targetTemplate,
                    TargetTokenParser = new TokenParser(cc.Web, targetTemplate),
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToDetailedString(cc));
                throw;
            }
        }
        #endregion

    }
}
