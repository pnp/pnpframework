using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers;
using PnP.Framework.Provisioning.Providers.Xml;
using PnP.Framework.Test.Framework.Providers.Extensibility;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Linq;

namespace PnP.Framework.Test.Framework.Providers
{
    [TestClass]
    public class XMLProvidersTests
    {
        #region Test variables

        static readonly string testContainer = "pnptest";
        static readonly string testContainerSecure = "pnptestsecure";
        static readonly string testTemplatesDocLib = "PnPTemplatesTests";

        private const string TEST_CATEGORY = "Framework Provisioning XML Providers";

        #endregion=

        #region Test initialize and cleanup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            if (!String.IsNullOrEmpty(TestCommon.DevSiteUrl))
            {
                CleanupTemplatesFromSharePointLibrary();
                UploadTemplatesToSharePointLibrary();
            }
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            if (!String.IsNullOrEmpty(TestCommon.DevSiteUrl))
            {
                CleanupTemplatesFromSharePointLibrary();
            }
        }

        private static void UploadTemplatesToSharePointLibrary()
        {
            var context = TestCommon.CreatePnPClientContext();

            var docLib = context.Web.CreateDocumentLibrary(testTemplatesDocLib);
            context.Load(docLib, d => d.RootFolder);
            context.ExecuteQueryRetry();

            var templatesToUpload = new string[] {
                "ProvisioningTemplate-2021-03-Sample-02.xml"
            };

            foreach (var tempFile in templatesToUpload)
            {
                // Create or overwrite the "myblob" blob with contents from a local file.
                using (var fileStream = System.IO.File.OpenRead(String.Format(@"{0}\..\..\..\Resources\Templates\{1}", AppDomain.CurrentDomain.BaseDirectory, tempFile)))
                {
                    docLib.RootFolder.UploadFile(tempFile, fileStream, true);
                    context.ExecuteQueryRetry();
                }
            }

        }

        private static void CleanupTemplatesFromSharePointLibrary()
        {
            var context = TestCommon.CreatePnPClientContext();

            var docLib = context.Web.GetListByTitle(testTemplatesDocLib);
            if (docLib != null)
            {
                context.Load(docLib);
                context.ExecuteQueryRetry();
                docLib.DeleteObject();
                context.ExecuteQueryRetry();
            }
        }
        #endregion

        #region XML File System tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplatesTest()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplates();

            Assert.IsTrue(result.Count > 15);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplate1Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2021-03-Sample-01.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 1);
            Assert.IsTrue(result.Files.Count == 1);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplate2Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2021-03-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.Files.Count == 5);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLFileSystemGetTemplate3Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2021-03-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        #endregion

        #region XML SharePoint tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSharePointGetTemplate1Test()
        {
            var _expectedID = "SPECIALTEAM";
            var _expectedVersion = 1.0;

            var context = TestCommon.CreatePnPClientContext();

            XMLSharePointTemplateProvider provider =
                new XMLSharePointTemplateProvider(context,
                    TestCommon.DevSiteUrl,
                    testTemplatesDocLib);

            var result = provider.GetTemplate("ProvisioningTemplate-2021-03-Sample-02.xml");

            Assert.AreEqual(_expectedID, result.Id);
            Assert.AreEqual(_expectedVersion, result.Version);
            Assert.IsTrue(result.Lists.Count == 2);
            Assert.IsTrue(result.SiteFields.Count == 4);
        }

        #endregion

        #region XInclude Tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLResolveValidXInclude()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2021-03-Valid-XInclude-01.xml");

            Assert.IsTrue(result.PropertyBagEntries.Count == 2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLResolveInvalidXInclude()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var result = provider.GetTemplate("ProvisioningTemplate-2021-03-NOT-Valid-XInclude-01.xml");

            Assert.IsTrue(result.PropertyBagEntries.Count == 0);
        }

        #endregion

        #region Provider Extensibility Tests

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLEncryptionTest()
        {
            // If you need to generate the test certificate, you can use the following PowerShell syntax:
            // New-SelfSignedCertificate -KeyUsageProperty All -KeyLength 2048 -KeyAlgorithm RSA -FriendlyName "PnPTestCertificate" -Subject "PnPTestCertificate" -NotAfter (Get-Date).AddYears(5) -CertStoreLocation "Cert:\CurrentUser\My"

            X509Certificate2 certificate = RetrieveCertificateFromStore(new X509Store(StoreLocation.CurrentUser), "PnPTestCertificate");

            if (certificate == null)
            {
                Assert.Inconclusive("Missing certificate with SN=PnPTestCertificate in CurrentUser Certificate Store, so can't test");
            }

            XMLEncryptionTemplateProviderExtension extension =
                new XMLEncryptionTemplateProviderExtension();

            extension.Initialize(certificate);

            ITemplateProviderExtension[] extensions = new ITemplateProviderExtension[] { extension };

            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var template = provider.GetTemplate("ProvisioningTemplate-2022-09-Sample-01.xml");
            template.DisplayName = "Ciphered template";

            provider.SaveAs(template, "ProvisioningTemplate-2022-09-Ciphered.xml", extensions);
            var result = provider.GetTemplate("ProvisioningTemplate-2022-09-Ciphered.xml", extensions);

            provider.Delete("ProvisioningTemplate-2022-09-Ciphered.xml");

            Assert.IsTrue(result.DisplayName == "Ciphered template");
        }

        private static X509Certificate2 RetrieveCertificateFromStore(X509Store store, String subjectName)
        {
            if (store == null)
                throw new ArgumentNullException(nameof(store));

            X509Certificate2 cert = null;

            try
            {
                store.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certs = store.Certificates.Find(X509FindType.FindBySubjectName, subjectName, false);

                if (certs.Count > 0)
                {
                    // Get the first certificate in the collection
                    cert = certs[0];
                }
            }
            finally
            {
                if (store != null)
                    store.Close();
            }

            return cert;
        }

        #endregion

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_SerializeDeserialize_201903()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201903Serializer();
            var template1 = provider.GetTemplate("ProvisioningSchema-2019-03-FullSample-01.xml", serializer);
            Assert.IsNotNull(template1);

            provider.SaveAs(template1, "ProvisioningSchema-2019-03-FullSample-01-OUT.xml", serializer);
            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningSchema-2019-03-FullSample-01-OUT.xml"));

            var template2 = provider.GetTemplate("ProvisioningSchema-2019-03-FullSample-01-OUT.xml", serializer);
            Assert.IsNotNull(template2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_SerializeDeserialize_201909()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var serializer = new XMLPnPSchemaV201909Serializer();
            var template1 = provider.GetTemplate("ProvisioningSchema-2019-09-FullSample-01.xml", serializer);
            Assert.IsNotNull(template1);

            provider.SaveAs(template1, "ProvisioningSchema-2019-09-FullSample-01-OUT.xml", serializer);
            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\ProvisioningSchema-2019-03-FullSample-01-OUT.xml"));

            var template2 = provider.GetTemplate("ProvisioningSchema-2019-09-FullSample-01-OUT.xml", serializer);
            Assert.IsNotNull(template2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_ProvisioningHierarchy_Load_201903()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-03-FullSample-01.xml");

            Assert.IsNotNull(hierarchy);
            Assert.IsNotNull(hierarchy.Templates);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_ProvisioningHierarchy_Save_201903()
        {
            XMLTemplateProvider provider =
                 new XMLFileSystemTemplateProvider(
                     String.Format(@"{0}..\..\..\Resources",
                     AppDomain.CurrentDomain.BaseDirectory),
                     "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-03-FullSample-01.xml");

            // Save the hierarchy
            var outputFile = "ProvisioningSchema-2019-03-FullSample-01-OUT.xml";
            var serializer = new XMLPnPSchemaV201903Serializer();
            provider.SaveAs(hierarchy, outputFile, serializer);

            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{outputFile}"));

            var hierarchy2 = provider.GetHierarchy(outputFile);
            Assert.IsNotNull(hierarchy2);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_ProvisioningHierarchy_Load_201909()
        {
            XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(
                    String.Format(@"{0}..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-09-FullSample-01.xml");

            Assert.IsNotNull(hierarchy);
            Assert.IsNotNull(hierarchy.Templates);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_ProvisioningHierarchy_Save_201909()
        {
            XMLTemplateProvider provider =
                 new XMLFileSystemTemplateProvider(
                     String.Format(@"{0}..\..\..\Resources",
                     AppDomain.CurrentDomain.BaseDirectory),
                     "Templates");

            var hierarchy = provider.GetHierarchy("ProvisioningSchema-2019-09-FullSample-01.xml");

            // Save the hierarchy
            var outputFile = "ProvisioningSchema-2019-09-FullSample-01-OUT.xml";
            provider.SaveAs(hierarchy, outputFile);

            Assert.IsTrue(System.IO.File.Exists($"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{outputFile}"));

            var hierarchy2 = provider.GetHierarchy(outputFile);
            Assert.IsNotNull(hierarchy2);
        }
    }
}
