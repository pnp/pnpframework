using Microsoft.SharePoint.Client;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities
{
    internal static class RESTUtilities
    {
#pragma warning disable CS0169
        private static ConcurrentDictionary<string, (string requestDigest, DateTime expiresOn)> requestDigestInfos = new ConcurrentDictionary<string, (string requestDigest, DateTime expiresOn)>();
#pragma warning restore CS0169

        internal static void SetAuthenticationCookies(this HttpClientHandler handler, ClientContext context)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var cookieString = CookieReader.GetCookie(context.Url)?.Replace("; ", ",")?.Replace(";", ",");
                if (cookieString == null)
                {
                    return;
                }
                var authCookiesContainer = new System.Net.CookieContainer();
                // Get FedAuth and rtFa cookies issued by ADFS when accessing claims aware applications.
                // - or get the EdgeAccessCookie issued by the Web Application Proxy (WAP) when accessing non-claims aware applications (Kerberos).
                IEnumerable<string> authCookies = null;
                if (Regex.IsMatch(cookieString, "FedAuth", RegexOptions.IgnoreCase))
                {
                    authCookies = cookieString.Split(',').Where(c => c.StartsWith("FedAuth", StringComparison.InvariantCultureIgnoreCase) || c.StartsWith("rtFa", StringComparison.InvariantCultureIgnoreCase));
                }
                else if (Regex.IsMatch(cookieString, "EdgeAccessCookie", RegexOptions.IgnoreCase))
                {
                    authCookies = cookieString.Split(',').Where(c => c.StartsWith("EdgeAccessCookie", StringComparison.InvariantCultureIgnoreCase));
                }
                if (authCookies != null)
                {
                    var siteUri = new Uri(context.Url);
                    var extension = siteUri.Host.Substring(siteUri.Host.LastIndexOf('.') + 1);
                    var cookieCollection = new CookieCollection();
                    foreach (var cookie in authCookies)
                    {
                        var cookieName = cookie.Substring(0, cookie.IndexOf("=")); // cannot use split as there might '=' in the value
                        var cookieValue = cookie.Substring(cookieName.Length + 1);
                        cookieCollection.Add(new Cookie(cookieName, cookieValue));
                    }
                    authCookiesContainer.Add(new Uri($"{siteUri.Scheme}://{siteUri.Host}"), cookieCollection);
                    var adminSiteUri = new Uri(siteUri.Scheme + "://" + siteUri.Authority.Replace($".sharepoint.{extension}", $"-admin.sharepoint.{extension}"));
                    authCookiesContainer.Add(adminSiteUri, cookieCollection);
                }
                handler.CookieContainer = authCookiesContainer;
            }
        }

        internal static async Task<string> GetRequestDigestWithCookieAuthAsync(this HttpClient httpClient, CookieContainer cookieContainer, ClientContext context)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var hostUrl = context.Url;
                if (requestDigestInfos.TryGetValue(hostUrl, out (string digestToken, DateTime expiresOn) requestDigestInfo))
                {
                    // We only have to add a request digest when running in dotnet core
                    if (DateTime.Now > requestDigestInfo.expiresOn)
                    {
                        requestDigestInfo = await GetRequestDigestInfoAsync(hostUrl, cookieContainer);
                        requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
                    }
                }
                else
                {
                    // admin url maybe?
                    requestDigestInfo = await GetRequestDigestInfoAsync(hostUrl, cookieContainer);
                    requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
                }
                return requestDigestInfo.digestToken;
            }
            return null;
        }

        private static async Task<(string digestToken, DateTime expiresOn)> GetRequestDigestInfoAsync(string siteUrl, CookieContainer cookieContainer)
        {
            using (var handler = new HttpClientHandler())
            {
                handler.CookieContainer = cookieContainer;
                using (var httpClient = new HttpClient(handler))
                {
                    string responseString = string.Empty;

                    string requestUrl = string.Format("{0}/_api/contextinfo", siteUrl.TrimEnd('/'));
                    using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl))
                    {
                        request.Headers.Add("accept", "application/json;odata=nometadata");
                        HttpResponseMessage response = await httpClient.SendAsync(request);

                        if (response.IsSuccessStatusCode)
                        {
                            responseString = await response.Content.ReadAsStringAsync();
                        }
                        else
                        {
                            var errorSb = new System.Text.StringBuilder();

                            errorSb.AppendLine(await response.Content.ReadAsStringAsync());
                            if (response.Headers.Contains("SPRequestGuid"))
                            {
                                var values = response.Headers.GetValues("SPRequestGuid");
                                if (values != null)
                                {
                                    var spRequestGuid = values.FirstOrDefault();
                                    errorSb.AppendLine($"ServerErrorTraceCorrelationId: {spRequestGuid}");
                                }
                            }

                            throw new Exception(errorSb.ToString());
                        }

                        var contextInformation = JsonSerializer.Deserialize<JsonElement>(responseString);

                        string formDigestValue = contextInformation.GetProperty("FormDigestValue").GetString();
                        int expiresIn = contextInformation.GetProperty("FormDigestTimeoutSeconds").GetInt32();
                        return (formDigestValue, DateTime.Now.AddSeconds(expiresIn - 30));
                    }
                }
            }
        }

        /// <summary>
        /// Executes a REST Get request. 
        /// </summary>
        /// <param name="web">The current web to execute the request against</param>
        /// <param name="endpoint">The full endpoint url, exluding the URL of the web, e.g. /_api/web/lists</param>
        /// <returns></returns>
        internal static async Task<string> ExecuteGetAsync(this Web web, string endpoint, string cultureLanguageName = null)
        {
            string returnObject = null;
            var accessToken = web.Context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                web.EnsureProperty(w => w.Url);

                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(web.Context as ClientContext);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    var requestUrl = $"{web.Url}{endpoint}";
                    using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl))
                    {
                        request.Headers.Add("accept", "application/json;odata=nometadata");
                        if (!string.IsNullOrEmpty(accessToken))
                        {
                            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                            var requestDigest = await (web.Context as ClientContext).GetRequestDigestAsync().ConfigureAwait(false);
                            request.Headers.Add("X-RequestDigest", requestDigest);
                        }
                        else
                        {
                            if (web.Context.Credentials is NetworkCredential networkCredential)
                            {
                                handler.Credentials = networkCredential;
                            }
                            request.Headers.Add("X-RequestDigest", await httpClient.GetRequestDigestWithCookieAuthAsync(handler.CookieContainer, web.Context as ClientContext));
                        }
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
                }
            }
            return await Task.Run(() => returnObject);
        }

        internal static async Task<string> ExecutePostAsync(this Web web, string endpoint, string payload, string cultureLanguageName = null)
        {
            string returnObject = null;
            var accessToken = web.Context.GetAccessToken();

            using (var handler = new HttpClientHandler())
            {
                web.EnsureProperty(w => w.Url);

                if (string.IsNullOrEmpty(accessToken))
                {
                    handler.SetAuthenticationCookies(web.Context as ClientContext);
                }

                using (var httpClient = new PnPHttpProvider(handler))
                {
                    var requestUrl = $"{web.Url}{endpoint}";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
                    request.Headers.Add("accept", "application/json;odata=nometadata");
                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                        var requestDigest = await (web.Context as ClientContext).GetRequestDigestAsync().ConfigureAwait(false);
                        request.Headers.Add("X-RequestDigest", requestDigest);
                    }
                    else
                    {
                        if (web.Context.Credentials is NetworkCredential networkCredential)
                        {
                            handler.Credentials = networkCredential;
                        }
                        request.Headers.Add("X-RequestDigest", await httpClient.GetRequestDigestWithCookieAuthAsync(handler.CookieContainer, web.Context as ClientContext).ConfigureAwait(false));
                    }
                    if(!string.IsNullOrWhiteSpace(cultureLanguageName))
                    {
                        request.Headers.Add("Accept-Language",cultureLanguageName);
                    }

                    if (!string.IsNullOrEmpty(payload))
                    {
                        ////var jsonBody = JsonConvert.SerializeObject(postObject);
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
            }
            return await Task.Run(() => returnObject);
        }
    }
}
