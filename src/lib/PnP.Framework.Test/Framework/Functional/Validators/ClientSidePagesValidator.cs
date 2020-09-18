using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Xml;

namespace PnP.Framework.Test.Framework.Functional.Validators
{
    [TestClass]
    public class ClientSidePagesValidator : ValidatorBase
    {
        public ClientSidePagesValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05;
        }

        public bool Validate(ClientSidePageCollection sourcePages, Microsoft.SharePoint.Client.ClientContext ctx)
        {
            return true;
        }
    }
}
