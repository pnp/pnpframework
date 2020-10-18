using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;

namespace PnP.Framework.Modernization.Utilities
{
    /// <summary>
    /// Handles the "intercepting" of auth cookies that might have been added on the clientcontext object
    /// </summary>
    internal class CookieManager
    {
        private CookieContainer authCookiesContainer = null;

        internal CookieContainer GetCookies(ClientContext cc)
        {
            EventHandler<WebRequestEventArgs> cookieInterceptorHandler = CollectCookiesHandler();
            try
            {
                // Hookup a custom handler, assumes the original handler placing the cookies is ran first
                cc.ExecutingWebRequest += cookieInterceptorHandler;
                //// Trigger the handler to fire by loading something
                cc.ExecuteQueryRetry();
            }
            catch (Exception)
            {
                // Eating the exception
            }
            finally
            {
                // Disconnect the handler as we don't need it anymore
                cc.ExecutingWebRequest -= cookieInterceptorHandler;
            }

            if (this.authCookiesContainer != null && this.authCookiesContainer.Count > 0)
            {
                return this.authCookiesContainer;
            }

            return null;
        }
        

        private EventHandler<WebRequestEventArgs> CollectCookiesHandler()
        {
            return (s, e) =>
            {
                if (authCookiesContainer == null || (authCookiesContainer != null && authCookiesContainer.Count == 0))
                {
                    this.authCookiesContainer = CopyContainer(e.WebRequestExecutor.WebRequest.CookieContainer);
                }
            };
        }

        private CookieContainer CopyContainer(CookieContainer container)
        {
            if (container == null)
            {
                return null;
            }

            using (MemoryStream stream = new MemoryStream())
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(stream, container);
                stream.Seek(0, SeekOrigin.Begin);
                return (CookieContainer)formatter.Deserialize(stream);
            }
        }

    }
}
