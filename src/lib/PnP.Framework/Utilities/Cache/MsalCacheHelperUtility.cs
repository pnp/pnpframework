using Microsoft.Identity.Client.Extensions.Msal;
using System;
using System.Collections.Generic;
using System.IO;

namespace PnP.Framework.Utilities.Cache
{
    public class MsalCacheHelperUtility
    {

        private static MsalCacheHelper MsalCacheHelper;
        private static readonly object ObjectLock = new();

        private static class Config
        {
            // Cache settings
            public const string CacheFileName = "m365pnpmsal.cache";
            public readonly static string CacheDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ".M365PnPAuthService");

            public const string KeyChainServiceName = "M365.PnP.Framework";
            public const string KeyChainAccountName = "M365PnPAuthCache";

            public const string LinuxKeyRingSchema = "com.m365.pnp.auth.tokencache";
            public const string LinuxKeyRingCollection = MsalCacheHelper.LinuxKeyRingDefaultCollection;
            public const string LinuxKeyRingLabel = "MSAL token cache for M365 PnP Framework.";
            public static readonly KeyValuePair<string, string> LinuxKeyRingAttr1 = new KeyValuePair<string, string>("Version", "1");
            public static readonly KeyValuePair<string, string> LinuxKeyRingAttr2 = new KeyValuePair<string, string>("Product", "M365PnPAuth");
        }

        public static MsalCacheHelper CreateCacheHelper()
        {
            if (MsalCacheHelper == null)
            {
                lock (ObjectLock)
                {
                    if (MsalCacheHelper == null)
                    {
                        StorageCreationProperties storageProperties;

                        try
                        {
                            storageProperties = new StorageCreationPropertiesBuilder(
                                Config.CacheFileName,
                                Config.CacheDir)
                            .WithLinuxKeyring(
                                Config.LinuxKeyRingSchema,
                                Config.LinuxKeyRingCollection,
                                Config.LinuxKeyRingLabel,
                                Config.LinuxKeyRingAttr1,
                                Config.LinuxKeyRingAttr2)
                            .WithMacKeyChain(
                                Config.KeyChainServiceName,
                                Config.KeyChainAccountName)
                            .Build();

                            var cacheHelper = MsalCacheHelper.CreateAsync(storageProperties).ConfigureAwait(false).GetAwaiter().GetResult();

                            cacheHelper.VerifyPersistence();
                            MsalCacheHelper = cacheHelper;

                        }
                        catch (MsalCachePersistenceException)
                        {
                            // do not use the same file name so as not to overwrite the encrypted version
                            storageProperties = new StorageCreationPropertiesBuilder(
                                    Config.CacheFileName + ".plaintext",
                                    Config.CacheDir)
                                .WithUnprotectedFile()
                                .Build();

                            var cacheHelper = MsalCacheHelper.CreateAsync(storageProperties).ConfigureAwait(false).GetAwaiter().GetResult();
                            cacheHelper.VerifyPersistence();

                            MsalCacheHelper = cacheHelper;
                        }
                        catch
                        {
                            MsalCacheHelper = null;
                        }
                    }
                }
            }
            return MsalCacheHelper;
        }
    }
}
