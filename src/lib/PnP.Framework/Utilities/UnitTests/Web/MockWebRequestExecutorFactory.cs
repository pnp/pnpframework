using Microsoft.SharePoint.Client;
using PnP.Framework.Utilities.UnitTests.Helpers;
using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class MockWebRequestExecutorFactory : WebRequestExecutorFactory
    {
        public bool RunAsIntegrationTest { get; set; }
        IMockResponseProvider ResponseProvider { get; }
        public event EventHandler<RequestExecutedArgs> OnRequestExecuted;
        public IMockDataRepository MockDataRepository { get; set; }
        public MockDataCreator MockCreator { get; set; }
        public MockWebRequestExecutorFactory(IMockResponseProvider responseProvider,
            bool runAsIntegrationTests = false,
            IMockDataRepository repo = null)
        {
            ResponseProvider = responseProvider;
            RunAsIntegrationTest = runAsIntegrationTests;
            MockDataRepository = repo;
            if (RunAsIntegrationTest)
            {
                MockCreator = new MockDataCreator(MockDataRepository);
            }
            else if (MockDataRepository != null)
            {
                ResponseProvider = new MockEntryResponseProvider()
                {
                    ResponseEntries = MockDataRepository.LoadMockData()
                };
            }
        }
        public override WebRequestExecutor CreateWebRequestExecutor(ClientRuntimeContext context, string requestUrl)
        {
            if (RunAsIntegrationTest)
            {
                ComposedWebRequestExecutor executor = new ComposedWebRequestExecutor(new SPWebRequestExecutor(context, requestUrl));
                executor.OnRequestExecuted += OnRequestExecuted;
                if (MockDataRepository != null)
                {
                    executor.OnRequestExecuted += delegate (object sender, RequestExecutedArgs e)
                    {
                        MockCreator.AddToMockResponse(e);
                    };
                }
                return executor;
            }
            return new MockWebRequestExecutor(requestUrl, ResponseProvider);
        }

        public void SaveMockData()
        {
            if (RunAsIntegrationTest && MockDataRepository != null)
            {
                MockDataRepository.SaveMockData(MockCreator.Responses);
            }
        }
    }
}
