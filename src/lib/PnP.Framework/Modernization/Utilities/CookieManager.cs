﻿using Microsoft.SharePoint.Client;
using System;
using System.Net;

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
                    this.authCookiesContainer = CopyContainer(e.WebRequestExecutor.WebRequest.CookieContainer, e.WebRequestExecutor.WebRequest.RequestUri);
                }
            };
        }

        private CookieContainer CopyContainer(CookieContainer container, Uri uri)
        {
            if (container == null)
            {
                return null;
            }

            var newContainer = new CookieContainer();
            newContainer.Add(container.GetCookies(uri));

            var adminUri = AuthenticationManager.GetTenantAdministrationUri(uri.ToString());
            newContainer.Add(container.GetCookies(adminUri));

            return newContainer;
        }

    }
}
