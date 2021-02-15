using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Model
{
    public class ActionObjectPath<T>
    {
        public string GetResponse(List<MockResponseEntry<T>> mockResponseEntries, CSOMRequest request)
        {
            return Action.GetResponse(ObjectPath, mockResponseEntries, request);
        }
        public BaseAction Action { get; set; }
        public Identity ObjectPath { get; set; }

        public MockResponseEntry GenerateMock(string responseBody, string calledUrl)
        {
            List<object> possibleResponses = JsonConvert.DeserializeObject<List<object>>(responseBody);
            object responseId = possibleResponses.FirstOrDefault(FindResponse);
            object associatedResponse = null;
            if (responseId != null)
            {
                associatedResponse = possibleResponses[possibleResponses.IndexOf(responseId) + 1];
            }
            if (ObjectPath != null)
            {
                return ObjectPath.CreateMockResponse(associatedResponse, Action, calledUrl);
            }
            else
            {

                return null;
            }
        }

        private bool FindResponse(object response)
        {
            if (response is Int64)
            {
                return (Int64)response == Action.Id;
            }
            return false;
        }
    }
    public class ActionObjectPath : ActionObjectPath<object>
    {

    }
}
