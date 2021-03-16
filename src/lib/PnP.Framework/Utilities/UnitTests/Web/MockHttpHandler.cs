using Newtonsoft.Json;
using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class MockHttpHandler : DelegatingHandler
    {
        public List<MockHttpResponse> Responses { get; set; } = new List<MockHttpResponse>();
        /// <summary>
        /// This stores the number of calls to the same endpoint to alter between different calls.
        /// Not the best solution but sometimes we have to check if another action caused change in the same endpoint
        /// For example - retracting solution.
        /// </summary>
        private Dictionary<string, int> CallToTheAPIIndex { get; set; } = new Dictionary<string, int>();
        public string MockFilePath { get; set; }
        public MockHttpHandler(string mockFilePath)
        {
            MockFilePath = mockFilePath;
            if (File.Exists(MockFilePath))
            {
                Responses = JsonConvert.DeserializeObject<List<MockHttpResponse>>(File.ReadAllText(MockFilePath));
            }
        }
        public MockHttpHandler()
        {
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            string url = request.RequestUri.AbsoluteUri;
            if (!CallToTheAPIIndex.ContainsKey(url))
            {
                CallToTheAPIIndex.Add(url, 0);
            }
            else
            {
                CallToTheAPIIndex[url]++;
            }
            MockHttpResponse response = Responses.Where(resp => resp.Url == url).ToList()[CallToTheAPIIndex[url]];
            if (response != null)
            {
#pragma warning disable CA2000 // Dispose objects before losing scope
                HttpResponseMessage result = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
#pragma warning restore CA2000 // Dispose objects before losing scope
                result.Content = new StringContent(response.Content);
                return Task.FromResult(result);
            }
            return Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.NotFound));
        }
    }
}
