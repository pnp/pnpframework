using System.Net.Http.Headers;

namespace PnP.Framework.Utilities
{
    public record HttpResult<T>(T Result, HttpResponseHeaders ResponseHeaders);
}
