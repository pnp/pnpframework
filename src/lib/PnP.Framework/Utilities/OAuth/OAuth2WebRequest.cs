using System.Threading.Tasks;

namespace SharePointPnP.IdentityModel.Extensions.S2S.Protocols.OAuth2
{
    public class OAuth2WebRequest : System.Net.WebRequest
    {
        private static readonly System.TimeSpan DefaultTimeout = System.TimeSpan.FromMinutes(10.0);

        private readonly System.Net.WebRequest _innerRequest;

        private readonly OAuth2AccessTokenRequest _request;

        public OAuth2WebRequest(string requestUriString, OAuth2AccessTokenRequest request)
        {
            this._innerRequest = System.Net.WebRequest.Create(requestUriString);
            this._request = request;
        }

        public override System.Net.WebResponse GetResponse()
            => GetResponseAsync().GetAwaiter().GetResult();

        public override async Task<System.Net.WebResponse> GetResponseAsync()
        {
            string text = this._request.ToString();
            this._innerRequest.AuthenticationLevel = System.Net.Security.AuthenticationLevel.None;
            this._innerRequest.ContentLength = text.Length;
            this._innerRequest.ContentType = "application/x-www-form-urlencoded";
            this._innerRequest.Method = "POST";
            this._innerRequest.Timeout = (int)OAuth2WebRequest.DefaultTimeout.TotalMilliseconds;
            using (var streamWriter = new System.IO.StreamWriter(await this._innerRequest.GetRequestStreamAsync(), System.Text.Encoding.ASCII))
                await streamWriter.WriteAsync(text);

            return await this._innerRequest.GetResponseAsync();
        }
    }
}
