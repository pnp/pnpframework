using Microsoft.SharePoint.Client;

namespace PnP.Framework.Migration.PublishingPages
{
    internal static class PublishingPageFileProbe
    {
        public static bool Exists(ClientContext context, string serverRelativeUrl)
        {
            var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            context.Load(file, value => value.Exists);
            try
            {
                context.ExecuteQueryRetry();
                return file.Exists;
            }
            catch (ServerException exception) when (exception.ServerErrorTypeName == "System.IO.FileNotFoundException")
            {
                return false;
            }
        }
    }
}
