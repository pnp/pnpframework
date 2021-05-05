using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace PnP.Framework.Http
{
	/// <summary>
	/// Implementation of SharePoint WebRequestExecutorFactory that utilizes HttpClient
	/// </summary>
	/// <example>
	/// var authManager = new PnP.Framework.AuthenticationManager(clientId, certificate, tenantId);
	/// var clientContext = authManager.GetContext(siteUrl);
	/// clientContext.WebRequestExecutorFactory = new HttpClientWebRequestExecutorFactory(PnPHttpClient.Instance.GetHttpClient());
	/// clientContext.Load(clientContext.Web, w => w.Title);
	/// await clientContext.ExecuteQueryRetryAsync();
	/// </example>
	public class HttpClientWebRequestExecutorFactory : WebRequestExecutorFactory
    {
        private readonly HttpClient _httpClient;

        /// <summary>
        /// Creates a WebRequestExecutorFactory that utilizes the specified HttpClient
        /// </summary>
        /// <param name="httpClient">HttpClient to use when creating new web requests</param>
        public HttpClientWebRequestExecutorFactory(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        /// <summary>
        /// Creates a WebRequestExecutor that utilizes HttpClient
        /// </summary>
        /// <param name="context">A SharePoint ClientContext</param>
        /// <param name="requestUrl">The url to create the request for</param>
        /// <returns>A WebRequestExecutor object created for the passed site URL</returns>
        public override WebRequestExecutor CreateWebRequestExecutor(ClientRuntimeContext context, string requestUrl)
        {
            return new HttpClientSPWebRequestExecutor(_httpClient, context, requestUrl);
        }
    }

	/// <summary>
	/// Implementation of SharePoint WebRequestExecutor that utilizes HttpClient
	/// </summary>
	internal class HttpClientSPWebRequestExecutor : WebRequestExecutor
	{
		private readonly HttpWebRequest _webRequest;
		private readonly HttpRequestMessage _request;
		private readonly HttpClient _httpClient;
		private HttpResponseMessage _response;
		private string _requestContentType;
		private RequestStream _requestStream;

		/// <summary>
		/// Creates a WebRequestExecutorFactory that utilizes the specified HttpClient
		/// </summary>
		/// <param name="httpClient">HttpClient to use when creating new web requests</param>
		/// <param name="context">A SharePoint ClientContext</param>
        /// <param name="requestUrl">The url to create the request for</param>
		public HttpClientSPWebRequestExecutor(HttpClient httpClient, ClientRuntimeContext context, string requestUrl)
		{
			if (string.IsNullOrEmpty(requestUrl))
				throw new ArgumentNullException(nameof(requestUrl));

			_httpClient = httpClient;
			_request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
			_webRequest = (HttpWebRequest)System.Net.WebRequest.Create(requestUrl);

			_webRequest.Timeout = context.RequestTimeout;
			_webRequest.Method = "POST";
			_webRequest.Pipelined = false;
		}

        private class RequestStream : Stream
        {
            public override bool CanRead { get; } = true;
            public override bool CanSeek { get; } = true;
            public override bool CanWrite { get; } = true;
            public override long Length => BaseStream.Length;
            public Stream BaseStream { get; }
            public RequestStream(Stream baseStream)
            {
                BaseStream = baseStream;
            }

            public override long Position
            {
                get => BaseStream.Position;
                set => BaseStream.Position = value;
            }

            public override void Flush()
            {
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                return BaseStream.Read(buffer, offset, count);
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                return BaseStream.Seek(offset, origin);
            }

            public override void SetLength(long value)
            {
                BaseStream.SetLength(value);
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                BaseStream.Write(buffer, offset, count);
            }
        }

		private async Task ExecuteImplementation()
		{
			foreach (string webRequestHeaderKey in _webRequest.Headers.Keys)
			{
				_request.Headers.Add(webRequestHeaderKey, _webRequest.Headers[webRequestHeaderKey]);
			}
			if (_webRequest.UserAgent != null)
			{
				_request.Headers.UserAgent.ParseAdd(_webRequest.UserAgent);
			}
			_requestStream.Seek(0, SeekOrigin.Begin);
			_request.Content = new StreamContent(_requestStream);
			if (MediaTypeHeaderValue.TryParse(_requestContentType, out var parsedValue))
			{
				_request.Content.Headers.ContentType = parsedValue;
			}

			_response = await _httpClient.SendAsync(_request);
		}

		public override HttpWebRequest WebRequest => _webRequest;

		public override string RequestContentType
		{
			get => _requestContentType;
			set => _requestContentType = value;
		}

		public override string RequestMethod
		{
			get => _request.Method.ToString();
			set => _request.Method = new HttpMethod(value);
		}

		public override bool RequestKeepAlive
		{
			get => !_request.Headers.ConnectionClose.GetValueOrDefault();
			set => _request.Headers.ConnectionClose = !value;
		}

		public override WebHeaderCollection RequestHeaders => _webRequest.Headers;

		public override Stream GetRequestStream()
		{
			if (_requestStream == null)
			{
				_requestStream = new RequestStream(new MemoryStream());
			}
			else if (!_requestStream.BaseStream.CanWrite)
			{
				_requestStream.Dispose();
				_requestStream = new RequestStream(new MemoryStream());
			}
			return _requestStream;
		}

		public override void Execute()
		{
			Task.Run(ExecuteImplementation).GetAwaiter().GetResult();
		}

		public override Task ExecuteAsync()
		{
			return ExecuteImplementation();
		}

		public override HttpStatusCode StatusCode
		{
			get
			{
				if (_response == null)
					throw new InvalidOperationException();
				return _response.StatusCode;
			}
		}

		public override string ResponseContentType
		{
			get
			{
				if (_response == null)
					throw new InvalidOperationException();
				_response.Content.Headers.TryGetValues("Content-Type", out var contentType);
				return contentType.FirstOrDefault();
			}
		}

		public override WebHeaderCollection ResponseHeaders
		{
			get
			{
				if (_response == null)
					throw new InvalidOperationException();
				var whc = new WebHeaderCollection();
				foreach (var header in _response.Headers)
				{
					foreach (var value in header.Value)
					{
						whc.Add(header.Key, value);
					}
				}
				return whc;
			}
		}

		public override Stream GetResponseStream()
		{
			if (_response == null)
				throw new InvalidOperationException();
			return _response.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
		}

		public override void Dispose()
		{
			_request.Dispose();
			_requestStream.Dispose();
			base.Dispose();
		}
	}
}
