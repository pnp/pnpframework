using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Utilities.PnPSdk;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("PnPFramework.Test")]
namespace PnP.Framework.Utilities.UnitTests
{
    public class MockLegacyAuthenticationProvider : ILegacyAuthenticationProvider
    {
        public bool RequiresCookieAuthentication { get; set; }

        public Task AuthenticateRequestAsync(Uri resource, HttpRequestMessage request)
        {
            return Task.CompletedTask;
        }

        public Task<string> GetAccessTokenAsync(Uri resource, string[] scopes)
        {
            return GetAccessTokenAsync(resource);
        }

        public Task<string> GetAccessTokenAsync(Uri resource)
        {
            return Task.FromResult("Mock_Access_Token");
        }

        public string GetCookieHeader(Uri targetUrl)
        {
            return "MockCookieHeader";
        }

        public string GetRequestDigest()
        {
            return "MockRequestDigest";
        }
    }
    internal class MockLegacyAuthenticationProviderFactory : ILegacyAuthenticationProviderFactory
    {
        public ILegacyAuthenticationProvider GetAuthenticationProvider(ClientContext context)
        {
            return new MockLegacyAuthenticationProvider();
        }
    }
}
