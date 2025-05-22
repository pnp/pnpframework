using Microsoft.Identity.Client.Extensibility;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities.OAuth
{
    internal class DefaultBrowserUi : ICustomWebUi

    {
        private Action<string, int> _openBrowserAction = null;
        private string _successMessageHtml = string.Empty;
        private string _failureMessageHtml = string.Empty;
        private bool _fullHtml = false;
        public DefaultBrowserUi(Action<string, int> openBrowserAction, string successMessageHtml, string failureMessageHtml, bool fullHtml = false)
        {
            _openBrowserAction = openBrowserAction;
            _successMessageHtml = successMessageHtml;
            _failureMessageHtml = failureMessageHtml;
            _fullHtml = fullHtml;
        }

        private const string SuccessMessageHtml = "You successfully authenticated. Feel free to close this browser/tab.";
        private const string FailureMessageHtml = "You did not succesfully authenticate. Feel free to close this browser/tab.";
        private const string CloseWindowSuccessHtml = @"<html><head><title>Authentication Complete</title><style>body{{font-family: sans-serif;margin: 0}}.title{{font-size: 1.2em;background-color: darkgreen;padding: 4;color: white}}.message{{font-size: 1.0em;margin: 10}}</style></head><body><div class=""title"">Authentication complete</div><div class=""message"">{0}</div></body></html>";
        private const string CloseWindowFailureHtml = @"<html><head><title>Authentication Failed</title><style>body{{font-family: sans-serif;margin: 0}}.title{{font-size: 1.2em;background-color: darkred;padding: 4;color: white}}.message{{font-size: 1.0em;margin: 10;}}</style></head><body><div class=""title"">Authentication Failed</div><div class=""message"">{2}</br></br></br></br>Error details: error {0} error_description:{1}</div></body></html>";

        public async Task<Uri> AcquireAuthorizationCodeAsync(
            Uri authorizationUri,
            Uri redirectUri,
            CancellationToken cancellationToken)
        {
            if (!redirectUri.IsLoopback)
            {
                throw new ArgumentException("Only loopback redirect uri is supported with this WebUI. Configure http://localhost or http://localhost:port during app registration. ");
            }

            Uri result = await InterceptAuthorizationUriAsync(
                authorizationUri,
                redirectUri,
                cancellationToken)
                .ConfigureAwait(true);

            return result;
        }

        public static string FindFreeLocalhostRedirectUri()
        {
            TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
            try
            {
                listener.Start();
                int port = ((IPEndPoint)listener.LocalEndpoint).Port;
                return "http://localhost:" + port;
            }
            finally
            {
                listener?.Stop();
            }
        }

        private void OpenBrowser(string url, int port)
        {
            _openBrowserAction?.Invoke(url, port);
        }

        private async Task<Uri> InterceptAuthorizationUriAsync(
            Uri authorizationUri,
            Uri redirectUri,
            CancellationToken cancellationToken)
        {
            OpenBrowser(authorizationUri.ToString(), redirectUri.Port);
            using (var listener = new SingleMessageTcpListener(redirectUri.Port))
            {
                Uri authCodeUri = null;
                await listener.ListenToSingleRequestAndRespondAsync(
                    (uri) =>
                    {
                        //Trace.WriteLine("Intercepted an auth code url: " + uri.ToString());
                        authCodeUri = uri;

                        return GetMessageToShowInBrowserAfterAuth(uri);
                    },
                    cancellationToken)
                .ConfigureAwait(false);

                return authCodeUri;
            }
        }

        private string GetMessageToShowInBrowserAfterAuth(Uri uri)
        {
#if !NETFRAMEWORK
            // Parse the uri to understand if an error was returned. This is done just to show the user a nice error message in the browser.
            var authCodeQueryKeyValue = System.Web.HttpUtility.ParseQueryString(uri.Query);

            string errorString = authCodeQueryKeyValue.Get("error");
#else
            Dictionary<string, string> dicQueryString = uri.Query.Split('&').ToDictionary(c => c.Split('=')[0], c => Uri.UnescapeDataString(c.Split('=')[1]));
            var errorString = dicQueryString.ContainsKey("error") ? dicQueryString["error"] : null;

#endif
            if (!string.IsNullOrEmpty(errorString))
            {
                if (!_fullHtml)
                {
#if !NETFRAMEWORK
                    string errorDescription = authCodeQueryKeyValue.Get("error_description");
#else
                string errorDescription = dicQueryString.ContainsKey("error_description") ? dicQueryString["error_description"] : null;
#endif
                    return string.Format(
                        CultureInfo.InvariantCulture,
                        CloseWindowFailureHtml,
                        errorString,
                        errorDescription,
                        string.IsNullOrEmpty(_failureMessageHtml) ? FailureMessageHtml : _failureMessageHtml);
                }
                else
                {
                    return string.Format(_failureMessageHtml, errorString);
                }
            }

            if (!_fullHtml)
            {
                return string.Format(CloseWindowSuccessHtml, string.IsNullOrEmpty(_successMessageHtml) ? SuccessMessageHtml : _successMessageHtml);
            }
            else
            {
                return _successMessageHtml;
            }
        }
    }
}
