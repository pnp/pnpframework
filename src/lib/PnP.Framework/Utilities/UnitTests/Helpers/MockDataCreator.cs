using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Helpers
{
    public class MockDataCreator
    {
        protected IMockDataRepository Repo { get; }
        public List<MockResponseEntry<object>> Responses { get; } = new List<MockResponseEntry<object>>();
        public MockDataCreator(IMockDataRepository repo)
        {
            Repo = repo;
        }

        public virtual void AddToMockResponse(RequestExecutedArgs requestExecutedArgs)
        {
            if (requestExecutedArgs.RequestBody.Contains("GetUpdatedFormDigestInformation "))
            {
                return;
            }
            CSOMRequest request = MockEntryResponseProvider.GetRequest(requestExecutedArgs.RequestBody);
            List<ActionObjectPath<object>> requestedOperations = MockEntryResponseProvider.GetActionObjectPathsFromRequest<object>(request);
            foreach (ActionObjectPath<object> requestedOperation in requestedOperations)
            {
                MockResponseEntry mockResponseEntry = requestedOperation.GenerateMock(requestExecutedArgs.ResponseBody, requestExecutedArgs.CalledUrl);
                if (mockResponseEntry != null && mockResponseEntry.SerializedReturnValue != "{\"IsNull\":false}")
                {
                    Responses.Add(mockResponseEntry);
                }
            }
        }

        public void Save()
        {
            Repo.SaveMockData(Responses);
        }
    }
}
