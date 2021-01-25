using Microsoft.SharePoint.Client;
using PnP.Framework.Utilities.UnitTests.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Web
{
    public class MockWebRequestExecutor : WebRequestExecutor
    {
        public override string RequestContentType { get; set; }
        /// <summary>
        /// Required by ExecuteQueryWithRetry
        /// </summary>
        public override HttpWebRequest WebRequest { get; }
        public override string RequestMethod { get; set; } = "POST";
        public override bool RequestKeepAlive { get; set; }
        protected HttpStatusCode statusCode { get; set; }
        public override HttpStatusCode StatusCode
        {
            get
            {
                return statusCode;
            }
        }
        protected string responseContentType = "text/xml";
        public override string ResponseContentType
        {
            get
            {
                return responseContentType;
            }
        }
        protected Stream ResponseStream { get; set; }
        public string RequestUrl { get; protected set; }
        private ComposedStream requestStream { get; set; }
        protected IMockResponseProvider ResponseProvider { get; }
        public MockWebRequestExecutor(string requestUrl, IMockResponseProvider responseProvider) : base()
        {
            this.RequestUrl = requestUrl;
            ResponseProvider = responseProvider;
            WebRequest = System.Net.WebRequest.CreateHttp(requestUrl);
        }

        private WebHeaderCollection _responseHeaders { get; set; } = new WebHeaderCollection();
        private string responseString;
        public override WebHeaderCollection ResponseHeaders
        {
            get
            {
                return _responseHeaders;
            }
        }
        private WebHeaderCollection _requestHeaders { get; set; } = new WebHeaderCollection();

        public override WebHeaderCollection RequestHeaders
        {
            get
            {
                return _requestHeaders;
            }
        }
        public override void Execute()
        {
            try
            {
                string requestBody = GetRequestBody();
                statusCode = HttpStatusCode.OK;
                if (requestBody.Contains("GetUpdatedFormDigestInformation"))
                {
                    responseString = DigestResponse;
                    responseContentType = "text/xml";
                }
                else
                {
                    responseString = ResponseProvider.GetResponse(RequestUrl, RequestMethod.ToString(), requestBody);
                    //remove duplicated escape characters (possible due to anonymous guid serializtion)
                    responseString = responseString.Replace("\\\\/", "\\/");
                    //however add \ for Guids and Dates the expectedformat is Id: "\/Guid(...)\/"
                    responseString = responseString.Replace("\"/Guid(", "\"\\/Guid(");
                    responseString = responseString.Replace("\"/Date(", "\"\\/Date(");
                    responseString = responseString.Replace(")/\"", ")\\/\"");
                    responseContentType = "application/json";
                }
            }
            catch (Exception ex)
            {
                statusCode = HttpStatusCode.InternalServerError;
                responseString = ex.Message;
            }
            ResponseStream = new MemoryStream(Encoding.UTF8.GetBytes(responseString));
        }

        private string GetRequestBody()
        {
            string result = "";
            this.requestStream.BaseStream.Position = 0;
            using (StreamReader reader = new StreamReader(this.requestStream.BaseStream))
            {
                result = reader.ReadToEnd();
            }
            requestStream.BaseStream.Flush();
            requestStream.BaseStream.Close();

            return result;
        }
        /// <summary>
        /// Called by CSOM to prepare new request
        /// </summary>
        /// <returns></returns>
        public override Stream GetRequestStream()
        {
            if (requestStream == null)
            {
                requestStream = new ComposedStream(new MemoryStream());
            }
            else if (!requestStream.BaseStream.CanWrite)
            {
                requestStream = new ComposedStream(new MemoryStream());
            }
            return requestStream;
        }
        /// <summary>
        /// Called by CSOM to read the response
        /// </summary>
        /// <returns></returns>
        public override Stream GetResponseStream()
        {
            return ResponseStream;
        }

        string DigestResponse = "<?xml version=\"1.0\" encoding=\"utf-8\"?><soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"><soap:Body><GetUpdatedFormDigestInformationResponse xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\"><GetUpdatedFormDigestInformationResult><DigestValue>mock_digest_value</DigestValue><TimeoutSeconds>1800</TimeoutSeconds><WebFullUrl>http://sp.local/sites/test_site</WebFullUrl><LibraryVersion>16.0.10355.20000</LibraryVersion><SupportedSchemaVersions>14.0.0.0,15.0.0.0</SupportedSchemaVersions></GetUpdatedFormDigestInformationResult></GetUpdatedFormDigestInformationResponse></soap:Body></soap:Envelope>";


    }
}
