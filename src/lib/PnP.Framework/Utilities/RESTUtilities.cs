using Microsoft.SharePoint.Client;
using PnP.Framework.Http;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities
{
    internal static class RESTUtilities
    {

        /// <summary>
        /// Executes a REST Get request. 
        /// </summary>
        /// <param name="web">The current web to execute the request against</param>
        /// <param name="endpoint">The full endpoint url, exluding the URL of the web, e.g. /_api/web/lists</param>
        /// <param name="cultureLanguageName">If specified will be set as the Accept-Language header</param>
        /// <returns></returns>
        internal static async Task<string> ExecuteGetAsync(this Web web, string endpoint, string cultureLanguageName = null)
        {
            string returnObject = null;
            web.EnsureProperty(w => w.Url);

#pragma warning disable CA2000 // Dispose objects before losing scope
            var httpClient = PnPHttpClient.Instance.GetHttpClient(context: web.Context as ClientContext);
#pragma warning restore CA2000 // Dispose objects before losing scope

            var requestUrl = $"{web.Url}{endpoint}";
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl))
            {
                request.Headers.Add("accept", "application/json;odata=nometadata");

                await PnPHttpClient.AuthenticateRequestAsync(request, web.Context as ClientContext).ConfigureAwait(false);

                if (!string.IsNullOrWhiteSpace(cultureLanguageName))
                {
                    request.Headers.Add("Accept-Language", cultureLanguageName);
                }

                // Perform actual post operation
                HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                if (response.IsSuccessStatusCode)
                {
                    // If value empty, URL is taken
                    var responseString = await response.Content.ReadAsStringAsync();
                    if (responseString != null)
                    {
                        try
                        {

                            returnObject = responseString;

                        }
                        catch { }
                    }
                }
                else
                {
                    // Something went wrong...
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
            }

            return await Task.Run(() => returnObject);
        }

        internal static async Task<string> ExecutePostAsync(this Web web, string endpoint, string payload, string cultureLanguageName = null)
        {
            string returnObject = null;

            web.EnsureProperty(w => w.Url);

#pragma warning disable CA2000 // Dispose objects before losing scope
            var httpClient = PnPHttpClient.Instance.GetHttpClient(context: web.Context as ClientContext);
#pragma warning restore CA2000 // Dispose objects before losing scope

            var requestUrl = $"{web.Url}{endpoint}";

            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl))
            {
                request.Headers.Add("accept", "application/json;odata=nometadata");

                await PnPHttpClient.AuthenticateRequestAsync(request, web.Context as ClientContext).ConfigureAwait(false);

                if (!string.IsNullOrWhiteSpace(cultureLanguageName))
                {
                    request.Headers.Add("Accept-Language", cultureLanguageName);
                }

                if (!string.IsNullOrEmpty(payload))
                {
                    var requestBody = new StringContent(payload);
                    MediaTypeHeaderValue sharePointJsonMediaType = MediaTypeHeaderValue.Parse("application/json;odata=nometadata;charset=utf-8");
                    requestBody.Headers.ContentType = sharePointJsonMediaType;
                    request.Content = requestBody;
                }

                // Perform actual post operation
                HttpResponseMessage response = await httpClient.SendAsync(request, new System.Threading.CancellationToken());

                if (response.IsSuccessStatusCode)
                {
                    // If value empty, URL is taken
                    var responseString = await response.Content.ReadAsStringAsync();
                    if (responseString != null)
                    {
                        try
                        {

                            returnObject = responseString;

                        }
                        catch { }
                    }
                }
                else
                {
                    // Something went wrong...
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
            }

            return await Task.Run(() => returnObject);
        }
    }
}
