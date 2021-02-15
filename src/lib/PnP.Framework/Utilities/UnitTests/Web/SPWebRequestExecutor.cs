using Microsoft.SharePoint.Client;
using PnP.Framework.Utilities.UnitTests.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    internal class SPWebRequestExecutor : WebRequestExecutor
    {
        private HttpWebRequest webRequest;
        private HttpWebResponse webResponse;
        private ClientRuntimeContext baseContext;

        public SPWebRequestExecutor(ClientRuntimeContext context, string requestUrl)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (string.IsNullOrEmpty(requestUrl))
                throw new ArgumentNullException(nameof(requestUrl));
            this.baseContext = context;
            this.webRequest = (HttpWebRequest)System.Net.WebRequest.Create(requestUrl);
            this.webRequest.Timeout = context.RequestTimeout;
            this.webRequest.Method = "POST";
            this.webRequest.Pipelined = false;
        }

        public override HttpWebRequest WebRequest
        {
            get
            {
                return this.webRequest;
            }
        }

        public override string RequestContentType
        {
            get
            {
                return this.webRequest.ContentType;
            }
            set
            {
                this.webRequest.ContentType = value;
            }
        }

        public override string RequestMethod
        {
            get
            {
                return this.webRequest.Method;
            }
            set
            {
                this.webRequest.Method = value;
            }
        }

        public override bool RequestKeepAlive
        {
            get
            {
                return this.webRequest.KeepAlive;
            }
            set
            {
                this.webRequest.KeepAlive = value;
            }
        }
        ComposedStream requestStream;

        public override WebHeaderCollection RequestHeaders
        {
            get
            {
                return this.webRequest.Headers;
            }
        }

        public override Stream GetRequestStream()
        {
            if (requestStream == null)
            {
                requestStream = new ComposedStream(new MemoryStream());
            }
            else if (!requestStream.BaseStream.CanWrite)
            {
                requestStream.Dispose();
                requestStream = new ComposedStream(new MemoryStream());
            }
            return requestStream;
        }

        public override void Execute()
        {
            requestStream.BaseStream.Position = 0;
            Stream webRequestStream = webRequest.GetRequestStream();
            using (StreamWriter writer = new StreamWriter(webRequestStream))
            {
                using(StreamReader reader = new StreamReader(requestStream.BaseStream))
                {
                    string requestString = reader.ReadToEnd();
                    writer.Write(requestString);
                }
            }
            webRequestStream.Close();
            this.webResponse = (HttpWebResponse)webRequest.GetResponse();
            if (responseStream != null)
            {
                responseStream.Dispose();
            }
            responseStream = new MemoryStream();
            this.webResponse.GetResponseStream().CopyTo(responseStream);
        }

        public override async Task ExecuteAsync()
        {
            this.webRequest.GetRequestStream().Close();
            WebResponse resp = (WebResponse)await this.webRequest.GetResponseAsync();
            this.webResponse = (HttpWebResponse)resp;
        }

        public override HttpStatusCode StatusCode
        {
            get
            {
                if (this.webResponse == null)
                    throw new InvalidOperationException();
                return this.webResponse.StatusCode;
            }
        }

        public override string ResponseContentType
        {
            get
            {
                if (this.webResponse == null)
                    throw new InvalidOperationException();
                return this.webResponse.ContentType;
            }
        }

        public override WebHeaderCollection ResponseHeaders
        {
            get
            {
                if (this.webResponse == null)
                    throw new InvalidOperationException();
                return this.webResponse.Headers;
            }
        }
        private MemoryStream responseStream;
        public override Stream GetResponseStream()
        {
            if (this.webResponse == null)
                throw new InvalidOperationException();
            return responseStream;
        }

        public override void Dispose()
        {
            if (this.webResponse == null)
                return;
            this.webResponse.Close();
        }
    }
}
