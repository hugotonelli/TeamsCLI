using Integrative.Encryption;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace TeamsCLI.Authentication
{
    class TokenCacheHelper
    {
        public static void EnableSerialization(ITokenCache tokenCache)
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
        }

        /// <summary>
        /// Path to the token cache. Note that this could be something different for instance for MSIX applications:
        /// private static readonly string CacheFilePath = $"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\{AppName}\msalcache.bin";
        /// </summary>
        public static readonly string CacheFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location + ".msalcache.bin";

        private static readonly object FileLock = new object();

        private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                args.TokenCache.DeserializeMsalV3(File.Exists(CacheFilePath)
                    ? CrossProtect.Unprotect(File.ReadAllBytes(CacheFilePath),
                    new byte[0],
                    DataProtectionScope.CurrentUser)
                : null);
            }
        }

        private static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                lock (FileLock)
                {
                    // reflect changes in the persistent store
                    File.WriteAllBytes(CacheFilePath,
                        CrossProtect.Protect(args.TokenCache.SerializeMsalV3(),
                        new byte[0],
                        DataProtectionScope.CurrentUser)
                    );
                }
            }
        }
    }
}
