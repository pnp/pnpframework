using Microsoft.SharePoint.Client;
using PnP.Framework.Http;
using PnP.Framework.Utilities.UnitTests.Helpers;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests
{
    public class UnitTestClientContext : ClientContext
    {
        public bool RunInIntegration { get; set; }
        public string MockFilePath { get; protected set; }
        public UnitTestClientContext(string url, bool runInIntegration, string mockFilePath) : base(url)
        {
            RunInIntegration = runInIntegration;
            MockFilePath = mockFilePath;
            WebRequestExecutorFactory = BuildExecutorFactory(runInIntegration, mockFilePath);
        }

        public static MockExecutorFactory BuildExecutorFactory(bool runInIntegration, string mockFilePath)
        {
            FileMockResponseRepository repo = new FileMockResponseRepository(mockFilePath);
            MockResponseProvider responseProvider = new MockResponseProvider(repo.LoadMockData());
            return new MockExecutorFactory(responseProvider, repo, runInIntegration);
        }

        protected override void Dispose(bool disposing)
        {
            if (RunInIntegration && WebRequestExecutorFactory is MockExecutorFactory)
            {
                (WebRequestExecutorFactory as MockExecutorFactory).SaveMockData();
            }
            base.Dispose(disposing);
        }

        public static UnitTestClientContext GetUnitTestContext(ClientContext context, bool runInIntegration, string mockFilePath)
        {
            ClientContext result = new UnitTestClientContext(context.Url, runInIntegration, mockFilePath);
            result = context.Clone(result, new Uri(context.Url));
            result.WebRequestExecutorFactory = BuildExecutorFactory(runInIntegration, mockFilePath);

            return result as UnitTestClientContext;
        }
    }
}
