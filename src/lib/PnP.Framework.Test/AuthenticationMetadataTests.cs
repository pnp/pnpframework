using Microsoft.Identity.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Test
{
    [TestClass]
    public class AuthenticationMetadataTests
    {
        private const string TenantId = "11111111-1111-1111-1111-111111111111";
        private const string HomeTenantId = "22222222-2222-2222-2222-222222222222";
        private const string ClientId = "33333333-3333-3333-3333-333333333333";
        private const string Username = "alice@partner.example";
        private const string SiteUrl = "https://example.sharepoint.com/sites/test";
        private static readonly string[] Scopes = { "https://example.sharepoint.com/.default" };

        [TestMethod]
        public void ConfiguredProviderIsAvailableWithoutAcquiringAToken()
        {
            using var http = new TokenEndpoint();
            using var password = CreatePassword();
            using var manager = CreateCredentialsManager(password, http);

            Assert.AreEqual(AuthenticationMethod.UsernamePassword, manager.Authentication.Method);
            Assert.AreEqual(AuthenticationIdentityType.User, manager.Authentication.IdentityType);
            Assert.IsNull(manager.Authentication.AccountId);
            Assert.IsNull(manager.Authentication.Username);
            Assert.IsNull(manager.Authentication.AuthenticationTenantId);
            Assert.AreEqual(0, http.Calls);

            using var unconfigured = new AuthenticationManager();
            Assert.AreEqual(AuthenticationMethod.Unknown, unconfigured.Authentication.Method);
            Assert.AreEqual(AuthenticationIdentityType.Unknown, unconfigured.Authentication.IdentityType);

            using var suppliedToken = AuthenticationManager.CreateWithAccessToken(password);
            Assert.AreEqual(AuthenticationMethod.AccessToken, suppliedToken.Authentication.Method);
            Assert.AreEqual(AuthenticationIdentityType.Unknown, suppliedToken.Authentication.IdentityType);
            Assert.IsNull(suppliedToken.Authentication.AccountId);
        }

        [TestMethod]
        public async Task SuccessfulAcquisitionRecordsGuestIdentityAndSilentContextKeepsConfiguredProvider()
        {
            using var http = new TokenEndpoint();
            using var password = CreatePassword();
            var application = CreatePublicApplication(http);
            using var credentials = AuthenticationManager.CreateWithCredentials(ClientId, Username, password);
            SetApplication(credentials, "publicClientApplication", application);
            var beforeAcquisition = credentials.Authentication;

            await credentials.GetAccessTokenAsync(Scopes);
            var snapshot = credentials.Authentication;
            AssertGuestIdentity(snapshot);
            Assert.AreEqual(AuthenticationMethod.UsernamePassword, snapshot.Method);
            Assert.IsNull(beforeAcquisition.AccountId, "Previously returned snapshots must remain unchanged.");
            int callsAfterAcquisition = http.Calls;

            await credentials.GetAccessTokenAsync(Scopes);
            AssertGuestIdentity(credentials.Authentication);
            Assert.AreEqual(callsAfterAcquisition, http.Calls, "A silent cache hit must not request another token.");

            using var browser = AuthenticationManager.CreateWithInteractiveLogin(ClientId, redirectUrl: "http://localhost", tenantId: TenantId);
            SetApplication(browser, "publicClientApplication", application);
            using var context = await browser.GetContextAsync(SiteUrl);
            Assert.AreEqual(AuthenticationMethod.Browser, browser.Authentication.Method);
            AssertGuestIdentity(browser.Authentication);
            Assert.AreEqual(callsAfterAcquisition, http.Calls, "GetContext must use the seeded account without interaction.");
        }

        [TestMethod]
        public async Task OnBehalfOfProviderIsReportedBeforeAndAfterAcquisition()
        {
            using var http = new TokenEndpoint();
            using var manager = AuthenticationManager.CreateWithOnBehalfOf(ClientId, "offline-secret",
                new UserAssertion("offline-assertion"), TenantId);
            SetApplication(manager, "confidentialClientApplication", CreateConfidentialApplication(http));

            Assert.AreEqual(AuthenticationMethod.OnBehalfOf, manager.Authentication.Method);
            Assert.AreEqual(AuthenticationIdentityType.User, manager.Authentication.IdentityType);
            Assert.IsNull(manager.Authentication.AccountId);
            Assert.AreEqual(0, http.Calls);

            await manager.GetAccessTokenAsync(Scopes);
            Assert.AreEqual(AuthenticationMethod.OnBehalfOf, manager.Authentication.Method);
            AssertGuestIdentity(manager.Authentication);
        }

        [TestMethod]
        public async Task ApplicationAcquisitionDoesNotInventAUserIdentity()
        {
            using var http = new TokenEndpoint { Delegated = false };
            using var rsa = RSA.Create(2048);
            var request = new CertificateRequest("CN=offline-fixture", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            using var certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
            using var manager = AuthenticationManager.CreateWithCertificate(ClientId, certificate, TenantId);
            var application = ConfidentialClientApplicationBuilder.Create(ClientId).WithCertificate(certificate)
                .WithAuthority($"https://login.microsoftonline.com/{TenantId}")
                .WithInstanceDiscovery(false).WithHttpClientFactory(http).Build();
            SetApplication(manager, "confidentialClientApplication", application);

            using var context = await manager.GetContextAsync(SiteUrl);
            var snapshot = manager.Authentication;
            Assert.AreEqual(AuthenticationMethod.Certificate, snapshot.Method);
            Assert.AreEqual(AuthenticationIdentityType.Application, snapshot.IdentityType);
            Assert.AreEqual(TenantId, snapshot.AuthenticationTenantId);
            Assert.IsNull(snapshot.AccountId);
            Assert.IsNull(snapshot.Username);
            Assert.IsNull(snapshot.HomeTenantId);
        }

        [TestMethod]
        public async Task FailedAcquisitionPreservesLastSuccessfulSnapshot()
        {
            using var http = new TokenEndpoint();
            using var password = CreatePassword();
            using var manager = CreateCredentialsManager(password, http);
            await manager.GetAccessTokenAsync(Scopes);
            var snapshot = manager.Authentication;

            http.FailTokenRequests = true;
            int callsBeforeFailure = http.Calls;
            var differentResource = new[] { "https://graph.microsoft.com/.default" };
            await Assert.ThrowsExceptionAsync<MsalUiRequiredException>(() => manager.GetAccessTokenAsync(differentResource));
            Assert.IsTrue(http.Calls > callsBeforeFailure, "The failed acquisition must reach the fake token endpoint.");
            Assert.AreSame(snapshot, manager.Authentication);
            AssertGuestIdentity(manager.Authentication);
        }

        private static SecureString CreatePassword()
        {
            var password = new SecureString();
            password.AppendChar('x');
            return password;
        }

        private static IPublicClientApplication CreatePublicApplication(TokenEndpoint http)
        {
            return PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority($"https://login.microsoftonline.com/{TenantId}")
                .WithInstanceDiscovery(false).WithHttpClientFactory(http).Build();
        }

        private static IConfidentialClientApplication CreateConfidentialApplication(TokenEndpoint http)
        {
            return ConfidentialClientApplicationBuilder.Create(ClientId).WithClientSecret("offline-secret")
                .WithAuthority($"https://login.microsoftonline.com/{TenantId}")
                .WithInstanceDiscovery(false).WithHttpClientFactory(http).Build();
        }

        private static AuthenticationManager CreateCredentialsManager(SecureString password, TokenEndpoint http)
        {
            var manager = AuthenticationManager.CreateWithCredentials(ClientId, Username, password);
            SetApplication(manager, "publicClientApplication", CreatePublicApplication(http));
            return manager;
        }

        private static void SetApplication(AuthenticationManager manager, string fieldName, IClientApplicationBase application)
        {
            // Inject a fully functional MSAL client with an in-memory HTTP transport, so these tests
            // exercise acquisition and its cache without contacting a tenant or opening a browser.
            var field = typeof(AuthenticationManager).GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.IsNotNull(field);
            field.SetValue(manager, application);
        }

        private static void AssertGuestIdentity(AuthenticationInfo snapshot)
        {
            Assert.AreEqual(AuthenticationIdentityType.User, snapshot.IdentityType);
            Assert.AreEqual("home-user." + HomeTenantId, snapshot.AccountId);
            Assert.AreEqual(Username, snapshot.Username);
            Assert.AreEqual(HomeTenantId, snapshot.HomeTenantId);
            Assert.AreEqual(TenantId, snapshot.AuthenticationTenantId);
        }

        private sealed class TokenEndpoint : HttpMessageHandler, IMsalHttpClientFactory
        {
            private readonly HttpClient client;

            public TokenEndpoint()
            {
                client = new HttpClient(this, disposeHandler: false);
            }

            public int Calls { get; private set; }
            public bool Delegated { get; set; } = true;
            public bool FailTokenRequests { get; set; }

            public HttpClient GetHttpClient() => client;

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                Calls++;
                object response;
                var status = HttpStatusCode.OK;
                string path = request.RequestUri.AbsolutePath;
                if (path.Contains(".well-known"))
                {
                    var root = $"https://login.microsoftonline.com/{TenantId}/oauth2/v2.0";
                    response = new
                    {
                        authorization_endpoint = root + "/authorize",
                        token_endpoint = root + "/token",
                        issuer = $"https://login.microsoftonline.com/{TenantId}/v2.0",
                        jwks_uri = "https://login.microsoftonline.com/common/discovery/v2.0/keys"
                    };
                }
                else if (path.Contains("/userrealm/"))
                {
                    response = new { account_type = "Managed", cloud_instance_name = "microsoftonline.com" };
                }
                else if (path.EndsWith("/token", StringComparison.Ordinal))
                {
                    if (FailTokenRequests)
                    {
                        status = HttpStatusCode.BadRequest;
                        response = new { error = "invalid_grant", error_description = "Offline fixture rejected acquisition." };
                    }
                    else
                    {
                        var fields = new Dictionary<string, object>
                        {
                            ["token_type"] = "Bearer",
                            ["expires_in"] = 3600,
                            ["access_token"] = "offline-access-token",
                            ["scope"] = Scopes[0]
                        };
                        if (Delegated)
                        {
                            fields["refresh_token"] = "offline-refresh-token";
                            fields["client_info"] = Encode(new { uid = "home-user", utid = HomeTenantId });
                            fields["id_token"] = Encode(new { alg = "none" }) + "." + Encode(new
                            {
                                aud = ClientId,
                                iss = $"https://login.microsoftonline.com/{TenantId}/v2.0",
                                sub = "guest-subject",
                                oid = "guest-object",
                                tid = TenantId,
                                preferred_username = Username,
                                exp = DateTimeOffset.UtcNow.AddHours(1).ToUnixTimeSeconds()
                            }) + ".";
                        }
                        response = fields;
                    }
                }
                else
                {
                    throw new InvalidOperationException("Unexpected offline HTTP request: " + request.RequestUri);
                }

                return Task.FromResult(new HttpResponseMessage(status)
                {
                    Content = new StringContent(JsonSerializer.Serialize(response), Encoding.UTF8, "application/json")
                });
            }

            private static string Encode(object value)
            {
                return Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value)))
                    .TrimEnd('=').Replace('+', '-').Replace('/', '_');
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    client.Dispose();
                }
                base.Dispose(disposing);
            }
        }
    }
}
