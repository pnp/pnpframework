using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Test.Framework.Functional.Validators;

namespace PnP.Framework.Test.Framework.Functional.Implementation
{
    internal class PublishingImplementation : ImplementationBase
    {
        internal void SiteCollectionPublishing(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                var result = TestProvisioningTemplate(cc, "publishing_add.xml", Handlers.Publishing);
                PublishingValidator pubVal = new PublishingValidator();
                Assert.IsTrue(pubVal.Validate(result.SourceTemplate.Publishing, result.TargetTemplate.Publishing, cc));
            }
        }
    }
}