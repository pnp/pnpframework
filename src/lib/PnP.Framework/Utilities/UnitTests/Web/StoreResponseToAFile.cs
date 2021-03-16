using Newtonsoft.Json;
using PnP.Framework.Utilities.UnitTests.Model;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class StoreResponseToAFile : DelegatingHandler
    {
        public string MockFilePath { get; set; }
        public StoreResponseToAFile(string mockFilePath)
        {
            MockFilePath = mockFilePath;
            if (!Directory.Exists(MockFilePath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(MockFilePath));
            }
        }

        protected async override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            HttpResponseMessage response = await base.SendAsync(request, cancellationToken);
            MockHttpResponse mockResponse = new MockHttpResponse()
            {
                Url = request.RequestUri.AbsoluteUri,
                Body = request.Content != null ? request.Content.ReadAsStringAsync().Result : "",
                Content = response.Content != null ? response.Content.ReadAsStringAsync().Result : ""
            };
            var responses = JsonConvert.DeserializeObject<List<MockHttpResponse>>(File.Exists(MockFilePath) ? File.ReadAllText(MockFilePath): "[]");
            responses.Add(mockResponse);
            File.WriteAllText(MockFilePath, JsonConvert.SerializeObject(responses, MockResponseEntry.SerializerSettings));
            return response;
        }
    }
}
