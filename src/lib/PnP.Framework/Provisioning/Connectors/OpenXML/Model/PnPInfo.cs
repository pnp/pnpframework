using System;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Connectors.OpenXML.Model
{
    /// <summary>
    /// Global container of the PnP OpenXML file
    /// </summary>
    [Serializable]
    public class PnPInfo
    {
        /// <summary>
        /// The Manifest of the PnP OpenXML file
        /// </summary>
        public PnPManifest Manifest { get; set; } = new PnPManifest();

        /// <summary>
        /// Custom properties of the PnP OpenXML file
        /// </summary>
        public PnPProperties Properties { get; set; } = new PnPProperties();

        /// <summary>
        /// Files contained in the PnP OpenXML file
        /// </summary>
        public List<PnPFileInfo> Files { get; set; } = new List<PnPFileInfo>();

        /// <summary>
        /// Defines the mapping between original file names and OpenXML file names
        /// </summary>
        public PnPFilesMap FilesMap { get; set; }

        /// <summary>
        /// Specifies whether the file streams should be used for file contenets instead of the MemoryStream.
        /// </summary>
        public bool UseFileStreams { get; set; } = false;

        /// <summary>
        /// Path to be used for saving file contenets instead of the MemoryStream.
        /// </summary>
        public string PnPFilesPath { get; set; }
    }
}
