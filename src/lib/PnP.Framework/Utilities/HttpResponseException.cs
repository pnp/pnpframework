using System;

namespace PnP.Framework.Utilities
{
    public class HttpResponseException : ApplicationException
    {
        public HttpResponseException(string message, Exception exception, int statusCode) :base(message, exception)
        {
            StatusCode = statusCode;
        }

        public int StatusCode { get; }
    }
}
