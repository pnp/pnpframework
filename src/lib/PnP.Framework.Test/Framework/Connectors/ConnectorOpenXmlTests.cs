using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.IO;
using System.Linq;

namespace PnP.Framework.Test.Framework.Connectors
{
    [TestClass]
    public class ConnectorOpenXmlTests
    {
        private const string packageFileName = "TestTemplate.pnp";
        private const string packageFileNameBackwardsCompatibility = "TestTemplateOriginal.pnp";
        private const string packageFileNameBackwardsCompatibility2 = "TestTemplateOriginal2.pnp";

        #region Test initialize and cleanup

        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            OpenXMLSaveTemplateInternal();
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            // File system setup
            if (File.Exists(String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory)
                    + @"\Templates\TestTemplate.pnp"))
            {
                System.IO.File.Delete(String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory)
                    + @"\Templates\TestTemplate.pnp");
            }
        }

        #endregion

        #region OpenXML Connector tests

        /// <summary>
        /// Create a PNP OpenXML package file and add a sample template to it
        /// </summary>
        [TestMethod]
        public void OpenXMLSaveTemplate()
        {
            Boolean checkFileExistence = File.Exists(String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory)
                    + @"\Templates\TestTemplate.pnp");
            Assert.IsTrue(checkFileExistence);
        }

        private static void OpenXMLSaveTemplateInternal()
        {
            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector,
                "OfficeDevPnP Automated Test");

            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\Templates\ProvisioningSchema-2022-09-FullSample-01.xml", "", openXMLConnector);
            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagelogo.png", "Images", openXMLConnector);
            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagebg.jpg", "Images", openXMLConnector);

            if (openXMLConnector is ICommitableFileConnector)
            {
                ((ICommitableFileConnector)openXMLConnector).Commit();
            }
        }

        [TestMethod]
        public void OpenXMLLoadTemplateOriginal()
        {
            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var openXMLConnector = new OpenXMLConnector(packageFileNameBackwardsCompatibility, fileSystemConnector);
            var templateFile = openXMLConnector.GetFileStream("ProvisioningSchema-2015-12-FullSample-02.xml");

            XMLPnPSchemaV201903Serializer formatter = new XMLPnPSchemaV201903Serializer();
            var checkTemplate = formatter.IsValid(templateFile);

            Assert.IsTrue(checkTemplate);

            var image1 = openXMLConnector.GetFileStream("garagelogo.png", "Images");
            Assert.IsNotNull(image1);

            var image2 = openXMLConnector.GetFileStream("garagebg.jpg", "Images");
            Assert.IsNotNull(image2);
        }

        [TestMethod]
        public void OpenXMLDeleteFileFromTemplate()
        {
            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector);

            openXMLConnector.DeleteFile("garagelogo.png", "Images");

            var image1 = openXMLConnector.GetFileStream("garagelogo.png", "Images");
            Assert.IsNull(image1);
        }

        private static void SaveFileInPackage(String filePath, String container, FileConnectorBase connector)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                String fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
                connector.SaveFileStream(fileName, container, fs);
            }
        }

        [TestMethod]
        public void OpenXMLGetFileFromTemplate()
        {
            var fileSystemConnector = new FileSystemConnector(String.Format(@"{0}\..\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");
            var openXMLConnector = new OpenXMLConnector(packageFileName, fileSystemConnector);
            var file = openXMLConnector.GetFile("garagebg.jpg");
            Assert.IsNull(file);
            file = openXMLConnector.GetFile("garagebg.jpg", "Images");
            Assert.IsNotNull(file);
            Stream stream = openXMLConnector.GetFileStream("garagebg.jpg");
            Assert.IsNull(stream);
            stream = openXMLConnector.GetFileStream("garagebg.jpg", "Images");
            Assert.IsNotNull(stream.Length > 0);
        }

        [TestMethod]
        public void OpenXMLGetFilesFromFolder()
        {
            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            var openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector,
                "OfficeDevPnP Automated Test");

            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagelogo.png", "Images", openXMLConnector);
            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagelogo.png", "Images\\Test", openXMLConnector);
            SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagebg.jpg", "Images/Test", openXMLConnector);

            if (openXMLConnector is ICommitableFileConnector)
            {
                ((ICommitableFileConnector)openXMLConnector).Commit();
            }

            openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector,
                "OfficeDevPnP Automated Test");


            var folders = openXMLConnector.GetFolders();
            Assert.IsTrue(folders.Count > 0);
            Assert.IsTrue(folders.Exists(s => string.Equals(s, "Images", StringComparison.OrdinalIgnoreCase)));
            var files = openXMLConnector.GetFiles("Images");
            Assert.IsTrue(files.Count > 0);

            files = openXMLConnector.GetFiles("Images\\Test");
            Assert.IsTrue(files.Count == 2);

            files = openXMLConnector.GetFiles("Images/Test");
            Assert.IsTrue(files.Count == 2);

        }

        /// <summary>
        /// Mixing backfard and forward slashes in file pathes causes incorrect folder comprasion that leads to files duplication under the same path
        /// </summary>
        [TestMethod]
        public void OpenXMLFileDuplicationTest()
        {
            var fileSystemConnector = new FileSystemConnector(
                    String.Format(@"{0}\..\..\..\Resources",
                    AppDomain.CurrentDomain.BaseDirectory),
                    "Templates");

            int retries = 3;
            while (retries-- > 0)
            {

                var openXMLConnector = new OpenXMLConnector(packageFileName,
                fileSystemConnector,
                "OfficeDevPnP Automated Test");

                SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagelogo.png", "Images\\OpenXMLFileDuplicationTest", openXMLConnector);
                SaveFileInPackage(fileSystemConnector.Parameters[FileConnectorBase.CONNECTIONSTRING] + @"\garagebg.jpg", "Images/OpenXMLFileDuplicationTest", openXMLConnector);

                var files = openXMLConnector.GetFiles("Images\\OpenXMLFileDuplicationTest");
                Assert.IsTrue(files.Count == 2);

                files = openXMLConnector.GetFiles("Images/OpenXMLFileDuplicationTest");
                Assert.IsTrue(files.Count == 2);

                if (openXMLConnector is ICommitableFileConnector)
                {
                    ((ICommitableFileConnector)openXMLConnector).Commit();
                }
            }
        }

        /// <summary>
        /// Tests that the template can be loaded successfully from the XMLOpenXMLTemplateProvider given the template filename
        /// </summary>
        [TestMethod]
        public void OpenXMLFileLoadTemplateTest()
        {
            var fileSystemConnector = new FileSystemConnector(String.Format(@"{0}\..\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");
            var openXMLConnector = new OpenXMLConnector(packageFileName, fileSystemConnector);
            var templateFileName = openXMLConnector.GetFiles().FirstOrDefault(f => f.EndsWith(".xml"));

            var templateProvider = new XMLOpenXMLTemplateProvider(openXMLConnector);
            var template = templateProvider.GetTemplate(templateFileName);

            Assert.IsNotNull(template);
        }

        /// <summary>
        /// Save a template using XMLOpenXMLTemplateProvider and ensure it is saved correctly
        /// </summary>
        [TestMethod]
        public void XMLOpenXMLTemplateProvider_SaveAs()
        {
            string packageName = Guid.NewGuid().ToString() + ".pnp";
            string templateName = "a" + Guid.NewGuid().ToString();

            var fileSystemConnector = new FileSystemConnector(String.Format(@"{0}\..\..\..\Resources", AppDomain.CurrentDomain.BaseDirectory), "Templates");

            try
            {
                var openXMLConnector = new OpenXMLConnector(packageName, fileSystemConnector);

                var templateProvider = new XMLOpenXMLTemplateProvider(openXMLConnector);
                var template = new Provisioning.Model.ProvisioningTemplate()
                {
                    Description = "Test"
                };

                //Add the template to the package
                templateProvider.SaveAs(template, templateName);

                //Re-open it and check that it has been saved with the correct template name
                openXMLConnector = new OpenXMLConnector(packageName, fileSystemConnector);
                template = templateProvider.GetTemplate(templateName);

                Assert.AreEqual("Test", template?.Description);
            }
            finally
            {
                fileSystemConnector.DeleteFile(packageName);
            }
        }
        #endregion
    }
}
