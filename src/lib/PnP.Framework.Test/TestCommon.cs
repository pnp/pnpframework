using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using PnP.Framework.Utilities;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;

namespace PnP.Framework.Test
{
    static class TestCommon
    {
        private static Configuration configuration;

        public static string AppSetting(string key)
        {
            try
            {
                return configuration.AppSettings.Settings[key].Value;
            }
            catch
            {
                return null;
            }
        }

        #region Constructor
        static TestCommon()
        {
            // Load configuration in a way that's compatible with a .Net Core test project as well
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap
            {
                ExeConfigFilename = @"..\..\..\App.config" //Path to your config file
            };
            configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            // Read configuration data
            TenantUrl = AppSetting("SPOTenantUrl");
            DevSiteUrl = AppSetting("SPODevSiteUrl");
            O365AccountDomain = AppSetting("O365AccountDomain");
            DefaultSiteOwner = AppSetting("DefaultSiteOwner");


            if (string.IsNullOrEmpty(DefaultSiteOwner))
            {
                DefaultSiteOwner = AppSetting("SPOUserName");
            }

            if (string.IsNullOrEmpty(TenantUrl))
            {
                throw new ConfigurationErrorsException("Tenant site Url in App.config are not set up.");
            }

            if (string.IsNullOrEmpty(DevSiteUrl))
            {
                throw new ConfigurationErrorsException("Dev site url in App.config are not set up.");
            }



            // Trim trailing slashes
            TenantUrl = TenantUrl.TrimEnd(new[] { '/' });
            DevSiteUrl = DevSiteUrl.TrimEnd(new[] { '/' });

            if (!string.IsNullOrEmpty(AppSetting("SPOCredentialManagerLabel")))
            {
                var tempCred = CredentialManager.GetCredential(AppSetting("SPOCredentialManagerLabel"));

                UserName = tempCred.UserName;
                Password = tempCred.SecurePassword;
            }
            else
            {
                if (!String.IsNullOrEmpty(AppSetting("AppId")) &&
                         !String.IsNullOrEmpty(AppSetting("AppSecret")))
                {
                    AppId = AppSetting("AppId");
                    AppSecret = AppSetting("AppSecret");
                }
                else
                {
                    throw new ConfigurationErrorsException("Tenant credentials in App.config are not set up.");
                }
            }
        }
        #endregion

        #region Properties
        public static string TenantUrl { get; set; }
        public static string DevSiteUrl { get; set; }
        static string UserName { get; set; }
        static SecureString Password { get; set; }
        public static string AppId { get; set; }
        static string AppSecret { get; set; }

        public static string O365AccountDomain { get; set; }

        /// <summary>
        /// Specifies the SiteOwner if needed (AppOnly Context, ...).
        /// </summary>
        public static string DefaultSiteOwner { get; set; }

        public static string TestWebhookUrl
        {
            get
            {
                return AppSetting("WebHookTestUrl");
            }
        }

        public static String AzureADCertPfxPassword
        {
            get
            {
                return AppSetting("AzureADCertPfxPassword");
            }
        }
        public static String AzureADClientId
        {
            get
            {
                return AppSetting("AzureADClientId");
            }
        }

        public static String AzureADCertificateFilePath
        {
            get
            {
                return AppSetting("AzureADCertificateFilePath");
            }
        }

        public static String NoScriptSite
        {
            get
            {
                return AppSetting("NoScriptSite");
            }
        }
        public static String ScriptSite
        {
            get
            {
                return AppSetting("ScriptSite");
            }
        }
        #endregion

        #region Methods
        public static ClientContext CreateClientContext()
        {
            return CreateContext(DevSiteUrl);
        }

        public static ClientContext CreateClientContext(string url)
        {
            return CreateContext(url);
        }

        public static ClientContext CreateTenantClientContext()
        {
            return CreateContext(TenantUrl);
        }

        public static PnPClientContext CreatePnPClientContext(int retryCount = 10, int delay = 500)
        {
            PnPClientContext context;
            AuthenticationManager am = new AuthenticationManager();
            if (!String.IsNullOrEmpty(AppId) && !String.IsNullOrEmpty(AppSecret))
            {
                ClientContext clientContext;

                if (new Uri(DevSiteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    //clientContext = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, Core.Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(DevSiteUrl)), AppId, AppSecret, acsHostUrl: "windows-ppe.net", globalEndPointPrefix: "login");
                    clientContext = am.GetACSAppOnlyContext(DevSiteUrl, AppId, AppSecret, AzureEnvironment.PPE);
                }
                else if (new Uri(DevSiteUrl).DnsSafeHost.Contains("sharepoint.de"))
                {
                    clientContext = am.GetACSAppOnlyContext(DevSiteUrl, AppId, AppSecret, AzureEnvironment.Germany);
                }
                else
                {
                    clientContext = am.GetACSAppOnlyContext(DevSiteUrl, AppId, AppSecret);
                }
                context = PnPClientContext.ConvertFrom(clientContext, retryCount, delay);
            }
            else
            {
                using (var authMgr = new AuthenticationManager(UserName, Password))
                {
                    ClientContext clientContext = authMgr.GetContextAsync(DevSiteUrl).GetAwaiter().GetResult();
                    context = PnPClientContext.ConvertFrom(clientContext, retryCount, delay);
                }
            }

            context.RequestTimeout = 1000 * 60 * 15;
            return context;
        }


        public static bool AppOnlyTesting()
        {
            if (!String.IsNullOrEmpty(AppSetting("AppId")) &&
                !String.IsNullOrEmpty(AppSetting("AppSecret")) &&
                String.IsNullOrEmpty(AppSetting("SPOCredentialManagerLabel")) &&
                String.IsNullOrEmpty(AppSetting("SPOUserName")) &&
                String.IsNullOrEmpty(AppSetting("SPOPassword")))
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        private static ClientContext CreateContext(string contextUrl)
        {
            
            ClientContext context = null;
            if (!String.IsNullOrEmpty(AppId) && !String.IsNullOrEmpty(AppSecret))
            {
                using (AuthenticationManager am = new AuthenticationManager())
                {
                    context = am.GetACSAppOnlyContext(contextUrl, AppId, AppSecret);
                }
            }
            else
            {
                using (AuthenticationManager am = new AuthenticationManager(UserName, Password))
                {
                    context = am.GetContext(contextUrl);
                }
            }

            context.RequestTimeout = 1000 * 60 * 15;
            return context;
        }

        #endregion


        public static string AcquireTokenAsync(string resource, string scope = null)
        {
            var tenantId = TenantExtensions.GetTenantIdByUrl(TestCommon.AppSetting("SPOTenantUrl"));
            if (tenantId == null) return null;

            var clientId = TestCommon.AppSetting("AzureADClientId");

            if (string.IsNullOrEmpty(clientId) || Password == null || string.IsNullOrEmpty(UserName))
            {
                return null;
            }

            var username = UserName;
            var password = EncryptionUtility.ToInsecureString(Password);

            string body;
            string response;
            if (scope == null) // use v1 endpoint
            {
                body = $"grant_type=password&client_id={clientId}&username={username}&password={password}&resource={resource}";

                // TODO: If your app is a public client, then the client_secret or client_assertion cannot be included. If the app is a confidential client, then it must be included.
                // https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth-ropc
                //body = $"grant_type=password&client_id={clientId}&client_secret={clientSecret}&username={username}&password={password}&resource={resource}";

                response = HttpHelper.MakePostRequestForString($"https://login.microsoftonline.com/{tenantId}/oauth2/token", body, "application/x-www-form-urlencoded");
            }
            else // use v2 endpoint
            {
                body = $"grant_type=password&client_id={clientId}&username={username}&password={password}&scope={scope}";

                // TODO: If your app is a public client, then the client_secret or client_assertion cannot be included. If the app is a confidential client, then it must be included.
                // https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth-ropc
                //body = $"grant_type=password&client_id={clientId}&client_secret={clientSecret}&username={username}&password={password}&scope={scope}";

                response = HttpHelper.MakePostRequestForString($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", body, "application/x-www-form-urlencoded");
            }

            var json = JToken.Parse(response);
            return json["access_token"].ToString();
        }

        private static Assembly _newtonsoftAssembly;
        private static string _assemblyName;

        public static void FixAssemblyResolving(string assemblyName)
        {
            _assemblyName = assemblyName;
            _newtonsoftAssembly = Assembly.LoadFrom(Path.Combine(AssemblyDirectory, $"{assemblyName}.dll"));
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        private static string AssemblyDirectory
        {
            get
            {
                var codeBase = Assembly.GetExecutingAssembly().CodeBase;
                var uri = new UriBuilder(codeBase);
                var path = Uri.UnescapeDataString(uri.Path);

                return Path.GetDirectoryName(path);
            }
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return args.Name.StartsWith(_assemblyName) ? _newtonsoftAssembly : AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(assembly => assembly.FullName == args.Name);
        }

        public static void DeleteFile(ClientContext ctx, string serverRelativeFileUrl)
        {
            var file = ctx.Web.GetFileByServerRelativeUrl(serverRelativeFileUrl);
            ctx.Load(file, f => f.Exists);
            ctx.ExecuteQueryRetry();

            if (file.Exists)
            {
                file.DeleteObject();
                ctx.ExecuteQueryRetry();
            }
        }
    }
}
