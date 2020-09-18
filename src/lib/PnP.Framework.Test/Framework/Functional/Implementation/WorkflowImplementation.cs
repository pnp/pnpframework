using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Tests.Framework.Functional.Validators;
using System;

namespace PnP.Framework.Tests.Framework.Functional.Implementation
{
    internal class WorkflowImplementation : ImplementationBase
    {
        internal void Workflows(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web)
                {
                    HandlersToProcess = Handlers.Lists | Handlers.Workflows,
                    FileConnector = new FileSystemConnector(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates")
                };

                string xmlFileName = null;
                xmlFileName = "workflows_add_1605.xml";

                var result = TestProvisioningTemplate(cc, xmlFileName, Handlers.Lists | Handlers.Workflows, null, ptci);
                WorkflowValidator wv = new WorkflowValidator();
                Assert.IsTrue(wv.Validate(result.SourceTemplate.Workflows, result.TargetTemplate.Workflows, result.TargetTokenParser));
            }
        }
    }
}