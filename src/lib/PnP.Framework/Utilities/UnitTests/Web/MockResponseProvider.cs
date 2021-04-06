using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class MockResponseProvider : IMockResponseProvider
    {
        public List<MockResponse> Responses { get; private set; } = new List<MockResponse>();
        private int ResponseIndex { get; set; } = -1;
        public MockResponseProvider(List<MockResponse> responses = null)
        {
            if (responses != null)
            {
                Responses.AddRange(responses);
            }
        }
        public string GetResponse(string url, string verb, string body)
        {
            ResponseIndex++;
            MockResponse response = Responses.FirstOrDefault(resp => resp.Verb == verb && resp.Body == body);
            if (response != null)
                return response.Response;
            if(Responses.Count > ResponseIndex)
                response = Responses[ResponseIndex];
            if (response != null)
                return response.Response;
            return "[\r{\r\"SchemaVersion\":\"15.0.0.0\",\"LibraryVersion\":\"16.0.21103.12003\",\"ErrorInfo\":null,\"TraceCorrelationId\":\"8a8fb49f-608d-2000-ac45-0453dea45810\"\r}]";
        }
    }
}
