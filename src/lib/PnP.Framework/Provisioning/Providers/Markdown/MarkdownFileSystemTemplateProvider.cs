using PnP.Framework.Provisioning.Connectors;

namespace PnP.Framework.Provisioning.Providers.Markdown
{
    public class MarkdownFileSystemTemplateProvider : MarkdownTemplateProvider
    {
        public MarkdownFileSystemTemplateProvider() : base()
        {

        }

        public MarkdownFileSystemTemplateProvider(string connectionString, string container) :
            base(new FileSystemConnector(connectionString, container))
        {
        }
    }
}
