﻿using PnP.Framework.Diagnostics;
using PnP.Framework.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PnP.Framework.Provisioning.Connectors
{
    /// <summary>
    /// Connector for files in file system
    /// </summary>
    public class FileSystemConnector : FileConnectorBase
    {
        private readonly bool useFileStreams;

        #region Constructors
        /// <summary>
        /// Base constructor
        /// </summary>
        public FileSystemConnector()
            : base()
        {

        }

        /// <summary>
        /// FileSystemConnector constructor. Allows to directly set root folder and sub folder
        /// </summary>
        /// <param name="connectionString">Root folder (e.g. c:\temp or .\resources or . or .\resources\templates)</param>
        /// <param name="container">Sub folder (e.g. templates or resources\templates or blank</param>
        /// <param name="useFileStreams">Use FileStreams instead of MemoryStreams</param>
        public FileSystemConnector(string connectionString, string container, bool useFileStreams = false)
            : base()
        {
            if (String.IsNullOrEmpty(connectionString))
            {
                throw new ArgumentException("connectionString");
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }
            container = container.Replace('/', '\\');

            this.useFileStreams = useFileStreams;
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
                container = "";
            }
            container = container.Replace('/', '\\');

            List<string> result = new List<string>();

            string path = ConstructPath("", container);

            foreach (string file in Directory.EnumerateFiles(path, "*.*"))
            {
                result.Add(Path.GetFileName(file));
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
                container = "";
            }
            container = container.Replace('/', '\\');

            List<string> result = new List<string>();

            string path = ConstructPath("", container);

            foreach (string folder in Directory.EnumerateDirectories(path))
            {
                result.Add(folder.Substring(path.Length + 1));
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
        /// Returns a filename without a path
        /// </summary>
        /// <param name="fileName">Name of the file</param>
        /// <returns>Returns filename without path</returns>
        public override string GetFilenamePart(string fileName)
        {
            return Path.GetFileName(fileName);
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

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }
            container = container.Replace('/', '\\');

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
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }
            container = container.Replace('/', '\\');

            if (useFileStreams)
            {
                try
                {
                    string filePath = ConstructPath(fileName, container);
                    FileStream fileStream = File.OpenRead(filePath);
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileRetrieved, fileName, container);
                    fileStream.Position = 0;
                    return fileStream;
                }
                catch (Exception ex)
                {
                    if (ex is FileNotFoundException || ex is DirectoryNotFoundException)
                    {
                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileNotFound, fileName, container, ex.Message);
                        return null;
                    }

                    throw;
                }
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
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException(nameof(fileName));
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }
            container = container.Replace('/', '\\');

            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            try
            {
                string filePath = ConstructPath(fileName, container);

                // Ensure the target path exists
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                using (var fileStream = File.Create(filePath))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }

                Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileSaved, fileName, container);
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileSaveFailed, fileName, container, ex.Message);
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
            if (String.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("fileName");
            }

            if (String.IsNullOrEmpty(container))
            {
                container = "";
            }
            container = container.Replace('/', '\\');

            try
            {
                string filePath = ConstructPath(fileName, container);

                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileDeleted, fileName, container);
                }
                else
                {
                    Log.Warning(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileDeleteNotFound, fileName, container);
                }
            }
            catch (Exception ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileDeleteFailed, fileName, container, ex.Message);
                throw;
            }
        }
        #endregion

        #region Private methods
        private MemoryStream GetFileFromStorage(string fileName, string container)
        {
            try
            {
                string filePath = ConstructPath(fileName, container);

                MemoryStream stream;
                using (FileStream fileStream = File.OpenRead(filePath))
                {
                    stream = fileStream.ToMemoryStream();
                }

                Log.Info(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileRetrieved, fileName, container);
                stream.Position = 0;
                return stream;
            }
            catch (Exception ex)
            {
                if (ex is FileNotFoundException || ex is DirectoryNotFoundException)
                {
                    Log.Error(Constants.LOGGING_SOURCE, CoreResources.Provisioning_Connectors_FileSystem_FileNotFound, fileName, container, ex.Message);
                    return null;
                }

                throw;
            }
        }

        private string ConstructPath(string fileName, string container)
        {
            string filePath = "";

            if (container.IndexOf(@"\") > 0)
            {
                string[] parts = container.Split(new string[] { @"\" }, StringSplitOptions.RemoveEmptyEntries);
                filePath = Path.Combine(GetConnectionString(), parts[0]);

                if (parts.Length > 1)
                {
                    for (int i = 1; i < parts.Length; i++)
                    {
                        filePath = Path.Combine(filePath, parts[i]);
                    }
                }

                if (!String.IsNullOrEmpty(fileName))
                {
                    filePath = Path.Combine(filePath, fileName);
                }
            }
            else
            {
                if (!String.IsNullOrEmpty(fileName))
                {
                    filePath = Path.Combine(GetConnectionString(), container, fileName);
                }
                else
                {
                    filePath = Path.Combine(GetConnectionString(), container);
                }
            }

            return filePath;
        }

        #endregion
    }
}
