using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class MockResponseProvider : IMockResponseProvider
    {
        public List<MockResponse> Responses { get; private set; } = new List<MockResponse>();
        public MockResponseProvider(List<MockResponse> responses = null)
        {
            if (responses != null)
            {
                Responses.AddRange(responses);
            }
        }
        public string GetResponse(string url, string verb, string body)
        {
            MockResponse response = Responses.FirstOrDefault(resp =>resp.Verb == verb && resp.Body == body);
            if (response != null)
                return response.Response;
            return "{}";
        }
    }
}
