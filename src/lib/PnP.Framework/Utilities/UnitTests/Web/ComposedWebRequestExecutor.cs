using Microsoft.SharePoint.Client;
using PnP.Framework.Utilities.UnitTests.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class ComposedWebRequestExecutor : WebRequestExecutor
    {
        public WebRequestExecutor Executor { get; }
        public override string RequestContentType
        {
            get
            {
                return Executor.RequestContentType;
            }
            set
            {
                Executor.RequestContentType = value;
            }
        }

        public override WebHeaderCollection RequestHeaders
        {
            get
            {
                return Executor.RequestHeaders;
            }
        }

        public override string RequestMethod
        {
            get
            {
                return Executor.RequestMethod;
            }
            set
            {
                Executor.RequestMethod = value;
            }
        }
        public override bool RequestKeepAlive
        {
            get
            {
                return Executor.RequestKeepAlive;
            }
            set
            {
                Executor.RequestKeepAlive = value;
            }
        }

        public override HttpStatusCode StatusCode
        {
            get
            {
                return Executor.StatusCode;
            }
        }

        public override string ResponseContentType
        {
            get
            {
                return Executor.ResponseContentType;
            }
        }

        public override WebHeaderCollection ResponseHeaders
        {
            get
            {
                return Executor.ResponseHeaders;
            }
        }
        private Stream requestStream { get; set; }
        public event EventHandler<RequestExecutedArgs> OnRequestExecuted;

        public ComposedWebRequestExecutor(WebRequestExecutor executor)
        {
            Executor = executor;
        }

        public override Stream GetRequestStream()
        {
            return Executor.GetRequestStream();
        }
        public override HttpWebRequest WebRequest
        {
            get
            {
                return Executor.WebRequest;
            }
        }
        private string GetRequestBody()
        {
            string result = "";
            Stream requestStream = this.GetRequestStream();
            requestStream.Position = 0;
            using (StreamReader reader = new StreamReader(requestStream))
            {
                result = reader.ReadToEnd();
            }
            return result;
        }
        public override void Execute()
        {
            string requestBody = GetRequestBody();
            Executor.Execute();
            if (OnRequestExecuted != null)
            {
                Stream responseStream = GetResponseStream();
                responseStream.Position = 0;
                using (MemoryStream tempStream = new MemoryStream())
                {
                    responseStream.CopyTo(tempStream);
                    tempStream.Position = 0;
                    using (StreamReader reader = new StreamReader(tempStream))
                    {
                        string responseString = reader.ReadToEnd();
                        OnRequestExecuted.Invoke(this, new RequestExecutedArgs()
                        {
                            CalledUrl = Executor.WebRequest.RequestUri.ToString(),
                            RequestBody = requestBody,
                            ResponseBody = responseString
                        });
                    }
                }
                responseStream.Position = 0;
            }
        }

        public override Stream GetResponseStream()
        {
            return Executor.GetResponseStream();
        }
    }
}
