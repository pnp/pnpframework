using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Xml;
using System.IO;
using System.Linq;
using System.Text;

namespace PnP.Framework.Test.Framework.Providers
{
    [TestClass]
    public class ListInstanceDataSourceTests
    {
        /// <summary>
        /// One external list carrying BCS settings and one library carrying default column values.
        /// Both settings are StringDictionaryItem arrays on the same list type, which is what the
        /// serializers used to confuse.
        /// </summary>
        private const string Template = @"<?xml version=""1.0"" encoding=""utf-8""?>
<pnp:Provisioning xmlns:pnp=""http://schemas.dev.office.com/PnP/2022/09/ProvisioningSchema"">
  <pnp:Templates ID=""CONTAINER"">
    <pnp:ProvisioningTemplate ID=""SAMPLE"" Version=""1"">
      <pnp:Lists>
        <pnp:ListInstance Title=""Customers"" Url=""Lists/Customers"" TemplateType=""600"">
          <pnp:DataSource>
            <pnp:DataSourceItem Key=""LobSystemInstance"" Value=""ODS"" />
            <pnp:DataSourceItem Key=""EntityNamespace"" Value=""https://sample.contoso.com/ODS"" />
            <pnp:DataSourceItem Key=""Entity"" Value=""ODS.Customer"" />
            <pnp:DataSourceItem Key=""SpecificFinder"" Value=""CustomerRead Item"" />
          </pnp:DataSource>
        </pnp:ListInstance>
        <pnp:ListInstance Title=""Documents"" Url=""Shared Documents"" TemplateType=""101"">
          <pnp:DefaultColumnValues>
            <pnp:DefaultColumnValue Key=""Field1"" Value=""Custom Value 1"" />
            <pnp:DefaultColumnValue Key=""Field2"" Value=""Custom Value 2"" />
          </pnp:DefaultColumnValues>
        </pnp:ListInstance>
      </pnp:Lists>
    </pnp:ProvisioningTemplate>
  </pnp:Templates>
</pnp:Provisioning>";

        [TestMethod]
        public void DataSourceSurvivesATemplateRoundTrip()
        {
            ListInstance externalList = RoundTrip().Lists.First(l => l.Title == "Customers");

            Assert.AreEqual(4, externalList.DataSource.Count, "the external list lost its DataSource entries");
            Assert.AreEqual("ODS", externalList.DataSource["LobSystemInstance"]);
            Assert.AreEqual("ODS.Customer", externalList.DataSource["Entity"]);
        }

        [TestMethod]
        public void DefaultColumnValuesSurviveATemplateRoundTrip()
        {
            ListInstance library = RoundTrip().Lists.First(l => l.Title == "Documents");

            Assert.AreEqual(2, library.DefaultColumnValues.Count, "the library lost its DefaultColumnValues entries");
            Assert.AreEqual("Custom Value 1", library.DefaultColumnValues["Field1"]);
        }

        /// <summary>
        /// A list that declares no DataSource has to come back without one, rather than with the
        /// contents of the first other setting that happens to use the same schema item type.
        /// </summary>
        [TestMethod]
        public void DefaultColumnValuesAreNotReadAsADataSource()
        {
            ProvisioningTemplate loaded = Load();
            ListInstance library = loaded.Lists.First(l => l.Title == "Documents");

            Assert.AreEqual(0, library.DataSource.Count, "DefaultColumnValues were read as a DataSource");
        }

        private static ProvisioningTemplate Load()
        {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(Template));
            return XMLPnPSchemaFormatter.LatestFormatter.ToProvisioningTemplate(input);
        }

        private static ProvisioningTemplate RoundTrip()
        {
            var formatter = XMLPnPSchemaFormatter.LatestFormatter;
            using var saved = formatter.ToFormattedTemplate(Load());
            return formatter.ToProvisioningTemplate(saved);
        }
    }
}
