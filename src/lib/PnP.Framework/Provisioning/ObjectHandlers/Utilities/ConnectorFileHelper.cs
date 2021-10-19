using PnP.Framework.Provisioning.Connectors;
using System;
using System.IO;
using System.Net;

namespace PnP.Framework.Provisioning.ObjectHandlers.Utilities
{
    public static class ConnectorFileHelper
    {
        public static byte[] GetFileBytes(FileConnectorBase connector, string fileName)
        {
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
                if (!String.IsNullOrEmpty(connector.GetContainer()))
                {
                    if (container.StartsWith("/"))
                    {
                        container = container.TrimStart("/".ToCharArray());
                    }

                    container = $@"{connector.GetContainer()}\{container}";
                }
            }
            else
            {
                container = connector.GetContainer();
            }

            var stream = connector.GetFileStream(fileName, container);
            if (stream == null)
            {
                //Decode the URL and try again
                fileName = Uri.UnescapeDataString(fileName);
                stream = connector.GetFileStream(fileName, container);
            }

            if (stream == null)
                throw new ArgumentException($"The specified filename '{fileName}' cannot be found");

            byte[] returnData;

            using (var memStream = new MemoryStream())
            {
                stream.CopyTo(memStream);
                memStream.Position = 0;
                returnData = memStream.ToArray();
            }
            if (stream != null)
            {
                stream.Dispose();
            }
            return returnData;
        }
    }
}
