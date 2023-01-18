using PnP.Framework.Provisioning.Connectors.OpenXML.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.IO.Packaging;

namespace PnP.Framework.Provisioning.Connectors.OpenXML
{
    /// <summary>
    /// Extension class for PnP OpenXML package files
    /// </summary>
    public static class PnPPackageExtensions
    {
        /// <summary>
        /// Packs template as a memory stream
        /// </summary>
        /// <param name="pnpInfo">PnPInfo object</param>
        /// <returns>Returns MemoryStream</returns>
        public static MemoryStream PackTemplateAsStream(this PnPInfo pnpInfo)
        {
            MemoryStream stream = new MemoryStream();
            using (PnPPackage package = PnPPackage.Open(stream, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                SavePnPPackage(pnpInfo, package);
            }
            stream.Position = 0;
            return stream;
        }

        public static void PackTemplateToStream(this PnPInfo pnpInfo, Stream stream)
        {
            using (PnPPackage package = PnPPackage.Open(stream, FileMode.Create, FileAccess.Write))
            {
                SavePnPPackage(pnpInfo, package);
            }
        }

        /// <summary>
        /// Packs template as a stream array
        /// </summary>
        /// <param name="pnpInfo">PnPInfo object</param>
        /// <returns>Returns stream as an array</returns>
        public static Byte[] PackTemplate(this PnPInfo pnpInfo)
        {
            using (MemoryStream stream = PackTemplateAsStream(pnpInfo))
            {
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Unpacks template into PnP OpenXML package info object based on memory stream
        /// </summary>
        /// <param name="stream">MemoryStream</param>
        /// <returns>Returns site template</returns>
        public static PnPInfo UnpackTemplate(this MemoryStream stream)
        {
            PnPInfo siteTemplate;
            using (PnPPackage package = PnPPackage.Open(stream, FileMode.Open,
                stream.CanWrite ? FileAccess.ReadWrite : FileAccess.Read))
            {
                siteTemplate = LoadPnPPackage(package);
            }
            return siteTemplate;
        }

        /// <summary>
        /// Unpacks template into PnP OpenXML package info object
        /// </summary>
        /// <param name="packageBytes">Package Byte</param>
        /// <returns>Returns site template</returns>
        public static PnPInfo UnpackTemplate(this Byte[] packageBytes)
        {
            using (MemoryStream stream = new MemoryStream(packageBytes))
            {
                return UnpackTemplate(stream);
            }
        }

        /// <summary>
        /// Return filename as Internal filename
        /// </summary>
        /// <param name="filename">Name of the file</param>
        /// <returns>Returns filename as Internal filename</returns>
        public static string AsInternalFilename(this string filename)
        {
            return Guid.NewGuid() + Path.GetExtension(filename);
        }

        #region Private Methods for handling templates

        private static PnPInfo LoadPnPPackage(PnPPackage package)
        {
            PnPInfo pnpInfo = new PnPInfo
            {
                Manifest = package.Manifest,
                Properties = package.Properties,
                FilesMap = package.FilesMap,

                Files = new List<PnPFileInfo>()
            };

            foreach (KeyValuePair<String, PnPPackageFileItem> file in package.Files)
            {
                pnpInfo.Files.Add(
                    new PnPFileInfo
                    {
                        InternalName = file.Key,
                        OriginalName = package.FilesMap != null ?
                            (String.IsNullOrEmpty(file.Value.Folder) ?
                            package.FilesMap.Map[file.Key] :
                            package.FilesMap.Map[file.Key].Replace(file.Value.Folder + '/', "")) :
                            file.Key,
                        Folder = file.Value.Folder,
                        Content = file.Value.Content,
                    });
            }
            return pnpInfo;
        }

        private static void SavePnPPackage(PnPInfo pnpInfo, PnPPackage package)
        {
            Debug.Assert(pnpInfo.Files.TrueForAll(f => !string.IsNullOrWhiteSpace(f.InternalName)), "All files need an InternalFileName");
            if (!pnpInfo.UseFileStreams)
            {
                package.Manifest = pnpInfo.Manifest;
                package.Properties = pnpInfo.Properties;
                package.FilesMap = new PnPFilesMap(pnpInfo.Files.ToDictionary(f => f.InternalName, f => Path.Combine(f.Folder, f.OriginalName).Replace('\\', '/').TrimStart('/')));
                package.ClearFiles();
                if (pnpInfo.Files != null)
                {
                    foreach (PnPFileInfo file in pnpInfo.Files)
                    {
                        package.AddFile(file.InternalName, file.Content);
                    }
                }
            }
            else
            {
                // Package with Create mode does not allow reads. Prepare and write the parts along with their relations in one go.
                // This is a workaround for(Memory leak with Append mode) https://github.com/dotnet/runtime/issues/1544
                var uriPath = new Uri(PnPPackage.U_PROVISIONINGTEMPLATE_MANIFEST, UriKind.Relative);
                PackagePart manifest = package.Package.CreatePart(uriPath, PnPPackage.CT_PROVISIONINGTEMPLATE_MANIFEST, PnPPackage.PACKAGE_COMPRESSION_LEVEL);
                PnPPackage.SetXamlSerializedPackagePartValue(pnpInfo.Manifest, manifest);
                package.Package.CreateRelationship(uriPath, TargetMode.Internal, PnPPackage.R_PROVISIONINGTEMPLATE_MANIFEST);

                uriPath = new Uri(PnPPackage.U_PROVISIONINGTEMPLATE_PROPERTIES, UriKind.Relative);
                PackagePart properties = package.Package.CreatePart(uriPath, PnPPackage.CT_PROVISIONINGTEMPLATE_PROPERTIES, PnPPackage.PACKAGE_COMPRESSION_LEVEL);
                manifest.CreateRelationship(uriPath, TargetMode.Internal, PnPPackage.R_PROVISIONINGTEMPLATE_PROPERTIES);

                uriPath = new Uri(PnPPackage.U_FILES_ORIGIN, UriKind.Relative);
                PackagePart filesOrigin = package.Package.CreatePart(uriPath, PnPPackage.CT_ORIGIN, PnPPackage.PACKAGE_COMPRESSION_LEVEL);
                manifest.CreateRelationship(uriPath, TargetMode.Internal, PnPPackage.R_PROVISIONINGTEMPLATE_FILES_ORIGIN);

                uriPath = new Uri(PnPPackage.U_PROVISIONINGTEMPLATE_FILES_MAP, UriKind.Relative);
                PackagePart filesMap = package.Package.CreatePart(uriPath, PnPPackage.CT_PROVISIONINGTEMPLATE_FILES_MAP, PnPPackage.PACKAGE_COMPRESSION_LEVEL);
                PnPPackage.SetXamlSerializedPackagePartValue(new PnPFilesMap(pnpInfo.Files.ToDictionary(f => f.InternalName, f => Path.Combine(f.Folder, f.OriginalName).Replace('\\', '/').TrimStart('/'))), filesMap);
                manifest.CreateRelationship(uriPath, TargetMode.Internal, PnPPackage.R_PROVISIONINGTEMPLATE_FILES_MAP);

                if (pnpInfo.Files != null)
                {
                    foreach (PnPFileInfo file in pnpInfo.Files)
                    {

#if NET6_0_OR_GREATER
                        // Set the file stream options to delete the files automatically once closed.
                        var fileStreamOptions = new FileStreamOptions { Mode = FileMode.Open, Access = FileAccess.Read, Options = FileOptions.DeleteOnClose, Share = FileShare.Delete };
                        using (FileStream fs = File.Open(Path.Combine(pnpInfo.PnPFilesPath, file.InternalName).Replace('\\', '/').TrimStart('/'), fileStreamOptions))
#else
                        using (FileStream fs = File.OpenRead(Path.Combine(pnpInfo.PnPFilesPath, file.InternalName).Replace('\\', '/').TrimStart('/')))
#endif
                        {
                            package.AddFilePart(file.InternalName, fs);
                            filesOrigin.CreateRelationship(new Uri(PnPPackage.U_DIR_FILES + file.InternalName.TrimStart('/'), UriKind.Relative), TargetMode.Internal, PnPPackage.R_PROVISIONINGTEMPLATE_FILE);
                        }
                    }
                }
            }
        }
        #endregion
    }
}
