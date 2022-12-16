﻿using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using File = Microsoft.SharePoint.Client.File;

namespace PnP.Framework.Provisioning.Connectors
{

    /// <summary>
    /// Connector for files in SharePoint
    /// </summary>
    public class SharePointConnector : FileConnectorBase
    {
        #region public variables
        public const string CLIENTCONTEXT = "ClientContext";
        #endregion

        #region Constructors
        /// <summary>
        /// Base constructor
        /// </summary>
        public SharePointConnector()
            : base()
        {

        }

        /// <summary>
        /// SharePointConnector constructor. Allows to directly set root folder and sub folder
        /// </summary>
        /// <param name="clientContext">Client context for SharePoint connection</param>
        /// <param name="connectionString">Site collection URL (e.g. https://yourtenant.sharepoint.com/sites/dev) </param>
        /// <param name="container">Library + folder that holds the files (mydocs/myfolder)</param>
        public SharePointConnector(ClientRuntimeContext clientContext, string connectionString, string container)
            : base()
        {
            if (clientContext == null)
            {
                throw new ArgumentNullException(nameof(clientContext));
            }

            if (String.IsNullOrEmpty(connectionString))
            {
                throw new ArgumentException(nameof(connectionString));
            }

            if (String.IsNullOrEmpty(container))
            {
                throw new ArgumentException(nameof(container));
            }
            container = container.Replace('\\', '/');

            this.AddParameter(CLIENTCONTEXT, clientContext);
            this.AddParameterAsString(CONNECTIONSTRING, connectionString);
            this.AddParameterAsString(CONTAINER, container);
        }

        #endregion

        #region Base class overrides
        /// <summary>
        /// Get the files available in the default container
        /// </summary>
        /// <returns>List of files</returns>
        public override List<string> GetFiles()
        {
            return GetFiles(GetContainer());
        }

        /// <summary>
        /// Get the files available in the specified container
        /// </summary>
        /// <param name="container">Name of the container to get the files from</param>
        /// <returns>List of files</returns>
        public override List<string> GetFiles(string container)
        {
            if (String.IsNullOrEmpty(container))
            {
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');

            List<string> result = new List<string>();

            using (ClientContext cc = GetClientContext().Clone(GetConnectionString()))
            {
                List list = cc.Web.GetListByUrl(GetDocumentLibrary(container));
                string folders = GetUrlFolders(container);

                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = @"<View Scope='FilesOnly'><Query></Query></View>"
                };

                if (folders.Length > 0)
                {
                    camlQuery.FolderServerRelativeUrl = $"{list.RootFolder.ServerRelativeUrl}{folders}";
                }

                ListItemCollection listItems = list.GetItems(camlQuery);
                cc.Load(listItems);
                cc.ExecuteQueryRetry();

                foreach (var listItem in listItems)
                {
                    result.Add(listItem.FieldValues["FileLeafRef"].ToString());
                }
            }

            return result;
        }

        /// <summary>
        /// Get the folders of the default container
        /// </summary>
        /// <returns>List of folders</returns>
        public override List<string> GetFolders()
        {
            return GetFolders(GetContainer());
        }

        /// <summary>
        /// Get the folders of a specified container
        /// </summary>
        /// <param name="container">Name of the container to get the folders from</param>
        /// <returns>List of folders</returns>
        public override List<string> GetFolders(string container)
        {
            if (String.IsNullOrEmpty(container))
            {
                throw new ArgumentException("container");
            }
            container = container.Replace('\\', '/');

            List<string> result = new List<string>();

            using (ClientContext cc = GetClientContext().Clone(GetConnectionString()))
            {
                List list = cc.Web.GetListByUrl(GetDocumentLibrary(container));
                string folders = GetUrlFolders(container);

                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = @"<View><Query><Where><Eq><FieldRef Name='ContentType' /><Value Type='Text'>Folder</Value></Eq></Where></Query></View>"
                };

                if (folders.Length > 0)
                {
                    camlQuery.FolderServerRelativeUrl = $"{list.RootFolder.ServerRelativeUrl}{folders}";
                }

                ListItemCollection listItems = list.GetItems(camlQuery);
                cc.Load(listItems);
                cc.ExecuteQueryRetry();

                foreach (var listItem in listItems)
                {
                    result.Add(listItem.FieldValues["FileLeafRef"].ToString());
                }
            }

            return result;
        }

        /// <summary>
        /// Gets a file as string from the default container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <returns>String containing the file contents</returns>
        public override string GetFile(string fileName)
        {
            return GetFile(fileName, GetContainer());
        }

        /// <summary>
        /// Gets a file as string from the specified container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <param name="container">Name of the container to get the file from</param>
        /// <returns>String containing the file contents</returns>
        public override string GetFile(string fileName, string container)
        {
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (container != null)
            {
                container = container.Replace('\\', '/');
            }

            string result = null;
            MemoryStream stream = null;
            try
            {
                stream = GetFileFromStorage(fileName, container);

                if (stream == null)
                {
                    return null;
                }

                result = Encoding.UTF8.GetString(stream.ToArray());
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }

            return result;
        }

        /// <summary>
        /// Gets a file as stream from the default container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <returns>String containing the file contents</returns>
        public override Stream GetFileStream(string fileName)
        {
            return GetFileStream(fileName, GetContainer());
        }

        /// <summary>
        /// Gets a file as stream from the specified container
        /// </summary>
        /// <param name="fileName">Name of the file to get</param>
        /// <param name="container">Name of the container to get the file from</param>
        /// <returns>String containing the file contents</returns>
        public override Stream GetFileStream(string fileName, string container)
        {
            Log.Debug("SharePointConnector", "GetFileStream('{0}','{1}')", fileName, container);
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (container != null)
            {
                container = container.Replace('\\', '/');
            }

            return GetFileFromStorage(fileName, container);
        }

        /// <summary>
        /// Saves a stream to the default container with the given name. If the file exists it will be overwritten
        /// </summary>
        /// <param name="fileName">Name of the file to save</param>
        /// <param name="stream">Stream containing the file contents</param>
        public override void SaveFileStream(string fileName, Stream stream)
        {
            SaveFileStream(fileName, GetContainer(), stream);
        }

        /// <summary>
        /// Saves a stream to the specified container with the given name. If the file exists it will be overwritten
        /// </summary>
        /// <param name="fileName">Name of the file to save</param>
        /// <param name="container">Name of the container to save the file to</param>
        /// <param name="stream">Stream containing the file contents</param>
        public override void SaveFileStream(string fileName, string container, Stream stream)
        {
            if (container != null)
            {
                container = container.Replace('\\', '/');
            }

            try
            {
                using (ClientContext cc = GetClientContext().Clone(GetConnectionString()))
                {
                    Folder spFolder;

                    if (!string.IsNullOrEmpty(container))
                    {
                        List list = cc.Web.GetListByUrl(GetDocumentLibrary(container));

                        string folders = GetUrlFolders(container);

                        if (folders.Length == 0)
                        {
                            spFolder = list.RootFolder;
                        }
                        else
                        {
                            spFolder = list.RootFolder;
                            string[] parts = container.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                            for (int i = 1; i < parts.Length; i++)
                            {
                                var prevFolder = spFolder;
                                spFolder = spFolder.ResolveSubFolder(parts[i]);

                                if (spFolder == null)
                                {
                                    spFolder = prevFolder.CreateFolder(parts[i]);
                                }
                            }
                        }
                    }
                    else
                    {
                        spFolder = cc.Web.RootFolder;
                    }

                    spFolder.UploadFile(fileName, stream, true);
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_SharePoint_FileSaved, fileName, GetConnectionString(), container);
                }
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_SharePoint_FileSaveFailed, fileName, GetConnectionString(), container, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Deletes a file from the default container
        /// </summary>
        /// <param name="fileName">Name of the file to delete</param>
        public override void DeleteFile(string fileName)
        {
            DeleteFile(fileName, GetContainer());
        }

        /// <summary>
        /// Deletes a file from the specified container
        /// </summary>
        /// <param name="fileName">Name of the file to delete</param>
        /// <param name="container">Name of the container to delete the file from</param>
        public override void DeleteFile(string fileName, string container)
        {
            if (container != null)
            {
                container = container.Replace('\\', '/');
            }

            try
            {
                using (ClientContext cc = GetClientContext().Clone(GetConnectionString()))
                {
                    string fileServerRelativeUrl = GetFileServerRelativeUrl(cc, fileName, container);
                    File file = cc.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileServerRelativeUrl));
                    cc.Load(file);
                    cc.ExecuteQueryRetry();

                    if (file != null)
                    {
                        file.DeleteObject();
                        cc.ExecuteQueryRetry();
                        Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_SharePoint_FileDeleted, fileName, GetConnectionString(), container);
                    }
                    else
                    {
                        Log.Warning(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_SharePoint_FileDeleteNotFound, fileName, GetConnectionString(), container);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_SharePoint_FileDeleteFailed, fileName, GetConnectionString(), container, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns a filename without a path
        /// </summary>
        /// <param name="fileName">Name of the file</param>
        /// <returns>Returns a filename without a path</returns>
        public override string GetFilenamePart(string fileName)
        {
            return Path.GetFileName(fileName);
        }

        #endregion

        #region Private Methods
        private string GetFileServerRelativeUrl(ClientContext cc, string fileName, string container)
        {
            Folder spFolder;
            if (!string.IsNullOrEmpty(container))
            {
                List list = cc.Web.GetListByUrl(GetDocumentLibrary(container));
                string folders = GetUrlFolders(container);

                if (folders.Length == 0)
                {
                    spFolder = list.RootFolder;
                }
                else
                {
                    spFolder = list.RootFolder;
                    string[] parts = container.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);

                    int startFrom = 1;
                    if (parts[0].Equals("_catalogs", StringComparison.InvariantCultureIgnoreCase))
                    {
                        startFrom = 2;
                    }

                    for (int i = startFrom; i < parts.Length; i++)
                    {
                        spFolder = spFolder.ResolveSubFolder(parts[i]);
                    }
                }
            }
            else
            {
                spFolder = cc.Web.RootFolder;
            }

            spFolder.EnsureProperties(f => f.ServerRelativeUrl);

            return UrlUtility.Combine(spFolder.ServerRelativeUrl, fileName);
        }

        private MemoryStream GetFileFromStorage(string fileName, string container)
        {

            try
            {
                using (ClientContext cc = GetClientContext().Clone(GetConnectionString()))
                {
                    string fileServerRelativeUrl = GetFileServerRelativeUrl(cc, fileName, container);
                    var file = cc.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileServerRelativeUrl));
                    cc.Load(file);
                    cc.ExecuteQueryRetry();
                    if (file.Exists)
                    {
                        MemoryStream stream = new MemoryStream();
                        var streamResult = file.OpenBinaryStream();
                        cc.ExecuteQueryRetry();

                        streamResult.Value.CopyTo(stream);
                        streamResult.Value.Dispose();

                        Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_SharePoint_FileRetrieved, fileName, GetConnectionString(), container);

                        // Set the stream position to the beginning
                        stream.Position = 0;
                        return stream;
                    }

                    throw new Exception("File not found");
                }
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_SharePoint_FileNotFound, fileName, GetConnectionString(), container, ex.Message);
                return null;
            }
        }

        private string GetDocumentLibrary(string container)
        {
            string[] parts = container.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length > 1)
            {
                if (parts[0].Equals("_catalogs", StringComparison.InvariantCultureIgnoreCase))
                {
                    return $"_catalogs/{parts[1]}";
                }
            }

            return parts[0];
        }

        private string GetUrlFolders(string container)
        {
            string[] parts = container.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length > 1)
            {
                int startFrom = 1;
                if (parts[0].Equals("_catalogs", StringComparison.InvariantCultureIgnoreCase))
                {
                    startFrom = 2;
                }

                string folder = "";
                for (int i = startFrom; i < parts.Length; i++)
                {
                    folder = folder + "/" + parts[i];
                }

                return folder;
            }
            else
            {
                return "";
            }
        }

        private ClientRuntimeContext GetClientContext()
        {
            if (this.Parameters.ContainsKey(CLIENTCONTEXT))
            {
                return this.Parameters[CLIENTCONTEXT] as ClientRuntimeContext;
            }
            else
            {
                throw new Exception("No clientcontext specified");
            }
        }

        #endregion

        internal override string GetContainer()
        {
            if (this.Parameters.ContainsKey(CONTAINER))
            {
                return this.Parameters[CONTAINER].ToString();
            }
            return string.Empty;
        }

    }
}
