using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using TeamsCLI.Authentication;

namespace TeamsCLI
{
    class DeviceCodeAuthProvider : IAuthenticationProvider
    {
        private IPublicClientApplication _msalClient;
        private readonly string[] _scopes;
        private IAccount _userAccount;

        public DeviceCodeAuthProvider(string appId, string[] scopes, string tenantId)
        {
            _scopes = scopes;

            _msalClient = PublicClientApplicationBuilder
                .Create(appId)
                .WithAuthority(AadAuthorityAudience.AzureAdMyOrg, true)
                .WithTenantId(tenantId)
                .Build();
            TokenCacheHelper.EnableSerialization(_msalClient.UserTokenCache);

            //AccountSelector();
        }

        public void AccountSelector()
        {
            Console.WriteLine("Attempting to get user account from cache...");
            var _accounts = _msalClient.GetAccountsAsync().Result.ToArray();

            var cantAccounts = _accounts.Count();
            if (cantAccounts > 1)
            {
                int choice = -1;

                while (choice < 0 || choice > cantAccounts)
                {
                    Console.WriteLine("Choose the account you want to use:");
                    Console.WriteLine("0. Sign in with another account");
                    for(int i = 0; i < cantAccounts; i++)
                    {
                        Console.WriteLine($"{i + 1}. {_accounts[i].Username}");
                    }

                    try
                    {
                        choice = int.Parse(Console.ReadLine());
                    }
                    catch (System.FormatException)
                    {
                        choice = -1;
                    }
                }

                if (choice == 0)
                {
                    _userAccount = null;
                    return;
                }

                _userAccount = _accounts[choice - 1];

            }
            else
            {
                _userAccount = _accounts.FirstOrDefault();
            }
        }

        public async Task<IEnumerable<IAccount>> GetAccounts()
        {
            var _accounts = await _msalClient.GetAccountsAsync();
            return _accounts;
        }

        public void SetAccount(IAccount account)
        {
            _userAccount = account;
        }

        public async Task<string> GetAccessToken()
        {
            // If there is no saved user account, the user must sign-in
            if (_userAccount == null)
            {
                try
                {
                    // Invoke device code flow so user can sign-in with a browser
                    var result = await _msalClient.AcquireTokenWithDeviceCode(_scopes, callback =>
                    {
                        Console.WriteLine(callback.Message);
                        return Task.FromResult(0);
                    }).ExecuteAsync();

                    _userAccount = result.Account;
                    return result.AccessToken;
                }
                catch (Exception exception)
                {
                    Console.WriteLine($"Error getting access token: {exception.Message}");
                    return null;
                }
            }
            else
            {
                // If there is an account, call AcquireTokenSilent
                // By doing this, MSAL will refresh the token automatically if
                // it is expired. Otherwise it returs the cached token.

                var result = await _msalClient
                    .AcquireTokenSilent(_scopes, _userAccount)
                    .ExecuteAsync();

                return result.AccessToken;
            }
        }

        // This is the required function to implement IAuthenticationProvider
        // The Graph SDK will call this function each time it makes a Graph call.
        public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage)
        {
            requestMessage.Headers.Authorization =
                new AuthenticationHeaderValue("bearer", await GetAccessToken());
        }
    }
}
