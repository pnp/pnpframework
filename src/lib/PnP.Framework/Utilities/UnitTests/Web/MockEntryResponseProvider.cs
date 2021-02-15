using Newtonsoft.Json;
using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class MockEntryResponseProvider<T> : IMockResponseProvider
    {
        public List<MockResponseEntry<T>> ResponseEntries { get; set; } = new List<MockResponseEntry<T>>();
        protected List<MockResponseEntry<T>> CurrentUrlResponses { get; set; }
        public ResponseHeader ResponseHeader { get; set; }
        public MockEntryResponseProvider(ResponseHeader responseHeader)
        {
            ResponseHeader = responseHeader;
        }
        public MockEntryResponseProvider() : this(new ResponseHeader())
        {

        }
        public string GetResponse(string url, string verb, string body)
        {
            CurrentUrlResponses = ResponseEntries.Where(entry =>
                entry.Url == url || entry.Url + "/_vti_bin/client.svc/ProcessQuery" == url)
                .ToList();

            StringBuilder responseBuilder = new StringBuilder();
            responseBuilder.Append('[');
            responseBuilder.Append(JsonConvert.SerializeObject(ResponseHeader));
            CSOMRequest request = GetRequest(body);
            List<ActionObjectPath<T>> actionsInRequest = GetActionObjectPathsFromRequest<T>(request);

            foreach (ActionObjectPath<T> action in actionsInRequest)
            {
                int id = action.Action.Id;
                responseBuilder.Append($",{id}, {action.GetResponse(CurrentUrlResponses, request)}");
            }

            responseBuilder.Append(']');
            return responseBuilder.ToString();
        }

        public static CSOMRequest GetRequest(string body)
        {
            CSOMRequest request;
            XmlSerializer serializer = new XmlSerializer(typeof(CSOMRequest));
            using (var reader = new StringReader(body))
            {
                request = (CSOMRequest)serializer.Deserialize(reader);
            }

            return request;
        }

        public static List<ActionObjectPath<T>> GetActionObjectPathsFromRequest<T>(CSOMRequest request)
        {
            List<ActionObjectPath<T>> result = new List<ActionObjectPath<T>>();
            foreach (BaseAction action in request.Actions)
            {
                Identity associatedIdentity = request.ObjectPaths.FirstOrDefault(path => path.Id == action.ObjectPathId);

                result.Add(new ActionObjectPath<T>()
                {
                    Action = action,
                    ObjectPath = associatedIdentity
                });
            }
            return result;
        }
    }
    public class MockEntryResponseProvider : MockEntryResponseProvider<object>
    {

    }
}
