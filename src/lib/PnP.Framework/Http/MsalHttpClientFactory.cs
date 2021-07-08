using Microsoft.Identity.Client;

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;

namespace PnP.Framework.Http
{
    internal class MsalHttpClientFactory : IMsalHttpClientFactory
    {
        public HttpClient GetHttpClient()
        {
            return PnPHttpClient.Instance.GetHttpClient();
        }
    }
}
