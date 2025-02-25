using System.Net.Http.Headers;

namespace PnP.Framework.Utilities
{
    public class HttpResult<T>
    {
        public HttpResult(T Result, HttpResponseHeaders ResponseHeaders)
        {
            this.Result = Result;
            this.ResponseHeaders = ResponseHeaders;
        }

        public T Result { get; }
        public HttpResponseHeaders ResponseHeaders { get; }
    }
}