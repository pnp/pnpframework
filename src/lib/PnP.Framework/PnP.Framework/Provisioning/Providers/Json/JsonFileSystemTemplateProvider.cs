using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PnP.Framework.Provisioning.Connectors;

namespace PnP.Framework.Provisioning.Providers.Json
{
    public class JsonFileSystemTemplateProvider : JsonTemplateProvider
    {
        public JsonFileSystemTemplateProvider() : base()
        {

        }

        public JsonFileSystemTemplateProvider(string connectionString, string container) :
            base(new FileSystemConnector(connectionString, container))
        {
        }
    }
}
