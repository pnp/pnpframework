using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Utilities;
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
                var tempFileName = fileName.Replace('/', Path.DirectorySeparatorChar)
                                           .Replace('\\', Path.DirectorySeparatorChar);
                container = fileName.Substring(0, tempFileName.LastIndexOf(Path.DirectorySeparatorChar));
                fileName = fileName.Substring(tempFileName.LastIndexOf(Path.DirectorySeparatorChar) + 1);
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
                    container = Path.Combine(template.Connector.GetContainer(), container);
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
                folderToGrabFilesFrom = Path.Combine(directory.ParentTemplate.Connector.GetContainer(), directory.Src);
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
                            let filePath = Path.Combine(directory.Src, file)
                            select new Model.File(
                                filePath,
                                directory.Folder,
                                directory.Overwrite,
                                null, // No WebPartPages are supported with this technique
                                metadataProperties != null && metadataProperties.ContainsKey(filePath) ?
                                    metadataProperties[filePath] : null,
                                directory.Security,
                                directory.Level
                                ));

            if (directory.Recursive)
            {
                var subFolders = directory.ParentTemplate.Connector.GetFolders(folderToGrabFilesFrom);
                var parentFolder = directory;
                foreach (var folder in subFolders)
                {
                    directory.Src = Path.Combine(parentFolder.Src, folder);
                    directory.Folder = Path.Combine(parentFolder.Folder, folder);

                    result.AddRange(directory.GetDirectoryFiles(metadataProperties));

                    //Remove the subfolder path(added above) as the second subfolder should come under its parent folder and not under its sibling
                    parentFolder.Src = parentFolder.Src.Substring(0, parentFolder.Src.LastIndexOf(Path.DirectorySeparatorChar));
                    parentFolder.Folder = parentFolder.Folder.Substring(0, parentFolder.Folder.LastIndexOf(Path.DirectorySeparatorChar));
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

        internal static bool PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, ObjectHandlerBase objectHandlerBase, string serverRelativeUrl)
        {
            var success = false;
            if (creationInfo.PersistBrandingFiles)
            {
                if (creationInfo.FileConnector != null)
                {
                    if (UrlUtility.IsIisVirtualDirectory(serverRelativeUrl))
                    {
                        scope.LogWarning("File is not located in the content database. Not retrieving {0}", serverRelativeUrl);
                        return success;
                    }

                    try
                    {
                        var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
                        string fileName = string.Empty;
                        if (serverRelativeUrl.IndexOf("/") > -1)
                        {
                            fileName = serverRelativeUrl.Substring(serverRelativeUrl.LastIndexOf("/") + 1);
                        }
                        else
                        {
                            fileName = serverRelativeUrl;
                        }
                        web.Context.Load(file);
                        web.Context.ExecuteQueryRetry();
                        ClientResult<Stream> stream = file.OpenBinaryStream();
                        web.Context.ExecuteQueryRetry();

                        file.EnsureProperty(f => f.ServerRelativePath);
                        var baseUri = new Uri(web.Url);
                        var fullUri = new Uri(baseUri, file.ServerRelativePath.DecodedUrl);
                        var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));

                        // Configure the filename to use 
                        fileName = Uri.UnescapeDataString(fullUri.Segments[fullUri.Segments.Length - 1]);

                        // Build up a site relative container URL...might end up empty as well
                        String container = Uri.UnescapeDataString(folderPath.Replace(web.ServerRelativeUrl, "")).Trim('/').Replace("/", "\\");

                        using (Stream memStream = new MemoryStream())
                        {
                            CopyStream(stream.Value, memStream);
                            memStream.Position = 0;
                            if (!string.IsNullOrEmpty(container))
                            {
                                creationInfo.FileConnector.SaveFileStream(fileName, container, memStream);
                            }
                            else
                            {
                                creationInfo.FileConnector.SaveFileStream(fileName, memStream);
                            }
                        }
                        success = true;
                    }
                    catch (ServerException ex1)
                    {
                        // If we are referring a file from a location outside of the current web or at a location where we cannot retrieve the file an exception is thrown. We swallow this exception.
                        if (ex1.ServerErrorCode != -2147024809)
                        {
                            throw;
                        }
                        else
                        {
                            scope.LogWarning("File is not necessarily located in the current web. Not retrieving {0}", serverRelativeUrl);
                        }
                    }
                }
                else
                {
                    objectHandlerBase.WriteMessage("No connector present to persist footer logo.", ProvisioningMessageType.Error);
                    scope.LogError("No connector present to persist footer logo.");
                }
            }
            else
            {
                success = true;
            }
            return success;
        }

        private static void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }
    }
}

