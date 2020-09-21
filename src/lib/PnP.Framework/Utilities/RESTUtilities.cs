using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
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
        /// <returns></returns>
        internal static async Task<string> ExecuteGetAsync(this Web web, string endpoint)
        {
            string returnObject = null;
            var accessToken = web.Context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                web.EnsureProperty(w => w.Url);

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    var requestUrl = $"{web.Url}{endpoint}";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=nometadata");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (web.Context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }

                    var requestDigest = await (web.Context as ClientContext).GetRequestDigestAsync().ConfigureAwait(false);
                    request.Headers.Add("X-RequestDigest", requestDigest);

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
            }
            return await Task.Run(() => returnObject);
        }

        internal static async Task<string> ExecutePostAsync(this Web web, string endpoint, string payload)
        {
            string returnObject = null;
            var accessToken = web.Context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                web.EnsureProperty(w => w.Url);

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    var requestUrl = $"{web.Url}{endpoint}";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=nometadata");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                    }
                    else
                    {
                        if (web.Context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                    }
                    var requestDigest = await (web.Context as ClientContext).GetRequestDigestAsync().ConfigureAwait(false);
                    request.Headers.Add("X-RequestDigest", requestDigest);

                    if (!string.IsNullOrEmpty(payload))
                    {
                        ////var jsonBody = JsonConvert.SerializeObject(postObject);
                        var requestBody = new StringContent(payload);
                        MediaTypeHeaderValue sharePointJsonMediaType;
                        MediaTypeHeaderValue.TryParse("application/json;odata=nometadata;charset=utf-8", out sharePointJsonMediaType);
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
            }
            return await Task.Run(() => returnObject);
        }
    }
}
