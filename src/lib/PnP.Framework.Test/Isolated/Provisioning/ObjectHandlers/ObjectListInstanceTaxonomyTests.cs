using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml.Linq;

namespace PnP.Framework.Test.Isolated.Provisioning.ObjectHanlders
{
    [TestClass]
    public class ObjectListInstanceTaxonomyTests
    {
        [DataTestMethod]
        [DataRow("TaxonomyFieldType")]
        [DataRow("TaxonomyFieldTypeMulti")]
        public void ExtractFields_ExportsBackingFieldForTaxonomyMaterializedAsField(string fieldType)
        {
            const string siteUrl = "https://test.sharepoint.com/sites/TestWeb";
            var webId = new Guid("11111111-1111-1111-1111-111111111111");
            var listId = new Guid("22222222-2222-2222-2222-222222222222");
            var fieldId = new Guid("33333333-3333-3333-3333-333333333333");
            var backingFieldId = new Guid("44444444-4444-4444-4444-444444444444");
            var fieldSchema = $"<Field ID='{fieldId:B}' Name='Metadata' DisplayName='Metadata' Type='{fieldType}' />";
            var backingSchema = $"<Field ID='{backingFieldId:B}' Name='Metadata_0' Type='Note' Hidden='TRUE' SourceID='{listId:B}' />";
            var responses = new MockEntryResponseProvider();
            responses.ResponseEntries.Add(new MockResponseEntry
            {
                Url = siteUrl,
                PropertyName = "Web",
                ReturnValue = new { Id = $"/Guid({webId})/", Url = siteUrl, ServerRelativeUrl = "/sites/TestWeb", WebTemplate = "STS", Configuration = 0 }
            });
            responses.ResponseEntries.Add(new MockResponseEntry
            {
                Url = siteUrl,
                PropertyName = "RootWeb",
                ReturnValue = new { Id = $"/Guid({webId})/" }
            });
            responses.ResponseEntries.Add(new MockResponseEntry
            {
                Url = siteUrl,
                PropertyName = "Fields",
                ReturnValue = new { _Child_Items_ = Array.Empty<object>() }
            });
            responses.ResponseEntries.Add(new MockResponseEntry
            {
                Url = siteUrl,
                Method = "GetByTitle",
                ReturnValue = new
                {
                    Id = $"/Guid({listId})/",
                    Title = "Documents",
                    Fields = new
                    {
                        // Explicitly return SP.Field even when taxonomy types are already loaded.
                        _Child_Items_ = new[] { new { _ObjectType_ = "SP.Field", _ObjectIdentity_ = "taxonomy-field",
                            Id = $"/Guid({fieldId})/", InternalName = "Metadata", Title = "Metadata", Hidden = false,
                            TypeAsString = fieldType, SchemaXml = fieldSchema } }
                    }
                }
            });
            responses.ResponseEntries.Add(new MockResponseEntry
            {
                Url = siteUrl,
                PropertyName = "taxonomy-field",
                ReturnValue = new { Id = $"/Guid({fieldId})/", TextField = $"/Guid({backingFieldId})/" }
            });
            responses.ResponseEntries.Add(new MockResponseEntry
            {
                Url = siteUrl,
                Method = "GetById",
                ReturnValue = new { Id = $"/Guid({backingFieldId})/", SchemaXml = backingSchema }
            });

            using (var context = new ClientContext(siteUrl))
            {
                context.WebRequestExecutorFactory = new MockWebRequestExecutorFactory(responses);
                var sourceList = context.Web.Lists.GetByTitle("Documents");
                context.Load(sourceList, list => list.Id, list => list.Title, list => list.Fields);
                context.ExecuteQuery();
                Assert.AreEqual(typeof(Microsoft.SharePoint.Client.Field), sourceList.Fields[0].GetType());

                var handler = new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.Export);
                var list = new ListInstance { Title = "Documents" };
                typeof(ObjectListInstance).GetMethod("ExtractFields", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(handler, new object[] { context.Web, sourceList, new List<FieldRef>(), list,
                        new List<Microsoft.SharePoint.Client.List> { sourceList },
                        new ProvisioningTemplateCreationInformation(context.Web), new ProvisioningTemplate() });

                Assert.AreEqual(2, list.Fields.Count);
                var backing = XElement.Parse(list.Fields[0].SchemaXml);
                Assert.AreEqual(backingFieldId, Guid.Parse(backing.Attribute("ID").Value));
                Assert.IsNull(backing.Attribute("SourceID"));
                var taxonomy = XElement.Parse(list.Fields[1].SchemaXml);
                Assert.AreEqual(fieldId, Guid.Parse(taxonomy.Attribute("ID").Value));
                Assert.AreEqual(fieldType, taxonomy.Attribute("Type").Value);
            }
        }
    }
}
