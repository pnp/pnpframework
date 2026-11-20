using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Http;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Test.Isolated.Http
{
    [TestClass]
    public class HttpClientWebRequestExecutorFactoryTests
    {
        private const string SiteUrl = "https://test.sharepoint.com/sites/TestWeb";
        private const string UserAgent = "ISV|Contoso|App/1.0";

        [TestMethod]
        public async Task ExecuteAsync_WithUserAgent_SendsUserAgentOnce()
        {
            var handler = new CapturingHandler();
            using var context = new ClientContext(SiteUrl);
            var factory = new HttpClientWebRequestExecutorFactory(new HttpClient(handler));
            using var executor = factory.CreateWebRequestExecutor(context, SiteUrl + "/_vti_bin/client.svc/ProcessQuery");
            executor.WebRequest.UserAgent = UserAgent;
            executor.RequestHeaders["X-RequestDigest"] = "digest";

            await executor.ExecuteAsync();

            CollectionAssert.AreEqual(new[] { UserAgent }, handler.UserAgent.ToArray());
            CollectionAssert.AreEqual(new[] { "digest" }, handler.RequestDigest.ToArray());
        }

        private sealed class CapturingHandler : HttpMessageHandler
        {
            public IEnumerable<string> UserAgent { get; private set; } = Enumerable.Empty<string>();

            public IEnumerable<string> RequestDigest { get; private set; } = Enumerable.Empty<string>();

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                UserAgent = request.Headers.UserAgent.Select(p => p.ToString()).ToArray();
                RequestDigest = request.Headers.TryGetValues("X-RequestDigest", out var digest) ? digest.ToArray() : Enumerable.Empty<string>();
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK));
            }
        }
    }
}
