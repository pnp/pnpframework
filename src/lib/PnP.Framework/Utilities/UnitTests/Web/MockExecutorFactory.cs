using Microsoft.SharePoint.Client;
using PnP.Framework.Utilities.UnitTests.Helpers;
using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class MockExecutorFactory : WebRequestExecutorFactory
    {
        public bool RunAsIntegrationTest { get; set; }
        IMockResponseProvider ResponseProvider { get; }
        public event EventHandler<RequestExecutedArgs> OnRequestExecuted;
        public IMockDataRepository<MockResponse> MockDataRepository { get; set; }
        public List<MockResponse> IntegrationResponses { get; private set; } = new List<MockResponse>();
        public MockExecutorFactory(IMockResponseProvider responseProvider,
            IMockDataRepository<MockResponse> repo,
            bool runAsIntegrationTests = false)
        {
            ResponseProvider = responseProvider;
            RunAsIntegrationTest = runAsIntegrationTests;
            MockDataRepository = repo;
        }
        public override WebRequestExecutor CreateWebRequestExecutor(ClientRuntimeContext context, string requestUrl)
        {
            if (RunAsIntegrationTest)
            {
                ComposedWebRequestExecutor executor = new ComposedWebRequestExecutor(new SPWebRequestExecutor(context, requestUrl));
                executor.OnRequestExecuted += OnRequestExecuted;
                executor.OnRequestExecuted += delegate (object sender, RequestExecutedArgs e)
                {
                    IntegrationResponses.Add(new MockResponse()
                    {
                        Url = e.CalledUrl,
                        Body = e.RequestBody,
                        Response = e.ResponseBody,
                        Verb = "POST"
                    });
                };
                return executor;
            }
            return new MockWebRequestExecutor(requestUrl, ResponseProvider);
        }

        public void SaveMockData()
        {
            if (RunAsIntegrationTest && MockDataRepository != null)
            {
                MockDataRepository.SaveMockData(IntegrationResponses);
            }
        }
    }
}
