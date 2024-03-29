﻿using PnP.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Web;

namespace PnP.Framework.Provisioning.ObjectHandlers.Utilities
{
    public static class FileUtilities
    {

        public static Stream GetFileStream(ProvisioningTemplate template, Model.File file)
        {
            return GetFileStream(template, file.Src);
        }

        public static Stream GetFileStream(ProvisioningTemplate template, string fileName)
        {
            // TODO: See if we can use ConnectorFileHelper instead

            var container = String.Empty;
            if (fileName.Contains(@"\") || fileName.Contains(@"/"))
            {
                var tempFileName = fileName.Replace(@"/", @"\");
                container = fileName.Substring(0, tempFileName.LastIndexOf(@"\"));
                fileName = fileName.Substring(tempFileName.LastIndexOf(@"\") + 1);
            }

            // add the default provided container (if any)
            if (!String.IsNullOrEmpty(container))
            {
                if (!String.IsNullOrEmpty(template.Connector.GetContainer()))
                {
                    if (container.StartsWith("/"))
                    {
                        container = container.TrimStart("/".ToCharArray());
                    }
                    container = $@"{template.Connector.GetContainer()}\{container}";
                }
            }
            else
            {
                container = template.Connector.GetContainer();
            }

            var stream = template.Connector.GetFileStream(fileName, container);
            if (stream == null)
            {
                //Decode the URL and try again
                fileName = Uri.UnescapeDataString(fileName);
                container = Uri.UnescapeDataString(container);
                stream = template.Connector.GetFileStream(fileName, container);
            }

            return stream;
        }

        public static List<Model.File> GetDirectoryFiles(this Model.Directory directory,
        Dictionary<String, Dictionary<String, String>> metadataProperties = null)
        {
            var result = new List<Model.File>();

            // If the connector has a container specified we need to take that in account to find the files we need
            string folderToGrabFilesFrom = directory.Src;
            if (!String.IsNullOrEmpty(directory.ParentTemplate.Connector.GetContainer()))
            {
                folderToGrabFilesFrom = directory.ParentTemplate.Connector.GetContainer() + @"\" + directory.Src;
            }

            var files = directory.ParentTemplate.Connector.GetFiles(folderToGrabFilesFrom);

            if (!String.IsNullOrEmpty(directory.IncludedExtensions) && directory.IncludedExtensions != "*.*")
            {
                var includedExtensions = directory.IncludedExtensions.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                files = files.Where(f => includedExtensions.Contains($"*{Path.GetExtension(f).ToLower()}")).ToList();
            }

            if (!String.IsNullOrEmpty(directory.ExcludedExtensions))
            {
                var excludedExtensions = directory.ExcludedExtensions.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                files = files.Where(f => !excludedExtensions.Contains($"*{Path.GetExtension(f).ToLower()}")).ToList();
            }

            result.AddRange(from file in files
                            select new Model.File(
                                directory.Src + @"\" + file,
                                directory.Folder,
                                directory.Overwrite,
                                null, // No WebPartPages are supported with this technique
                                metadataProperties != null && metadataProperties.ContainsKey(directory.Src + @"\" + file) ?
                                    metadataProperties[directory.Src + @"\" + file] : null,
                                directory.Security,
                                directory.Level
                                ));

            if (directory.Recursive)
            {
                var subFolders = directory.ParentTemplate.Connector.GetFolders(folderToGrabFilesFrom);
                var parentFolder = directory;
                foreach (var folder in subFolders)
                {
                    directory.Src = parentFolder.Src + @"\" + folder;
                    directory.Folder = parentFolder.Folder + @"\" + folder;

                    result.AddRange(directory.GetDirectoryFiles(metadataProperties));

                    //Remove the subfolder path(added above) as the second subfolder should come under its parent folder and not under its sibling
                    parentFolder.Src = parentFolder.Src.Substring(0, parentFolder.Src.LastIndexOf(@"\"));
                    parentFolder.Folder = parentFolder.Folder.Substring(0, parentFolder.Folder.LastIndexOf(@"\"));
                }
            }

            return (result);
        }

        public static Dictionary<string, Dictionary<string, string>> GetMetadataProperties(this Model.Directory directory)
        {
            Dictionary<string, Dictionary<string, string>> result = null;

            if (!string.IsNullOrEmpty(directory.MetadataMappingFile))
            {
                var metadataPropertiesStream = directory.ParentTemplate.Connector.GetFileStream(directory.MetadataMappingFile);
                if (metadataPropertiesStream != null)
                {
                    using (var sr = new StreamReader(metadataPropertiesStream))
                    {
                        var metadataPropertiesString = sr.ReadToEnd();
                        result = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(metadataPropertiesString);
                    }
                }
            }

            return (result);
        }
    }
}

