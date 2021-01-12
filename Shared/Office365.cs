namespace Zebble
{
    using Microsoft.Identity.Client;
    using System;
    using System.Linq;
    using System.Net.Http;
    using System.Threading.Tasks;
    using System.IO;
    using Olive;

    public class Office365
    {
        public static string ClientId { get; private set; }
        public static string[] Scopes { get; private set; }
        public static string UserName { get; private set; }

        static IPublicClientApplication _publicClientApplication;
        static string _tokenForUser;
        static DateTimeOffset? _expiration;

        public static void Initialize(string clientId, string[] scopes, string redirectUri)
        {
            if (clientId.IsEmpty())
            {
                Log.For<Office365>().Error(null, "Please provide the ClientId!");
                return;
            }

            ClientId = clientId;
            Scopes = scopes;

            _publicClientApplication = PublicClientApplicationBuilder.Create(clientId).WithRedirectUri(redirectUri).Build();
        }

        public static async Task SignIn()
        {
            AuthenticationResult result = null;

            try
            {
                var accounts = await _publicClientApplication.GetAccountsAsync();
                result = await _publicClientApplication.AcquireTokenSilent(Scopes, accounts.First()).ExecuteAsync();
            }
            catch
            {
                if (_tokenForUser.IsEmpty() || _expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
#if ANDROID
                    result = await _publicClientApplication.AcquireTokenInteractive(Scopes)
                        .WithParentActivityOrWindow(UIRuntime.CurrentActivity)
                        .ExecuteAsync();
#elif IOS
                    result = await _publicClientApplication.AcquireTokenInteractive(Scopes)
                        .WithParentActivityOrWindow(UIRuntime.Window.RootViewController)
                        .ExecuteAsync();
#else
                    result = await _publicClientApplication.AcquireTokenInteractive(Scopes)
                        .ExecuteAsync();
#endif
                }
            }

            _tokenForUser = result?.AccessToken;
            _expiration = result?.ExpiresOn;
            UserName = result?.Account.Username;

            await Task.CompletedTask;
        }

        public static async Task<T> GetRequest<T>(string url)
        {
            try
            {
                using var client = CreateHttpClient();

                var response = await client.GetAsync(url.AsUri());

                response.EnsureSuccessStatusCode();

                return await ReadAs<T>(response);
            }
            catch (Exception e)
            {
                Log.For<Office365>().Error(e);
            }

            return default;
        }

        public static async Task<T> PostRequest<T>(string url, string body)
        {
            try
            {
                using var client = CreateHttpClient();

                var response = await client.PostAsync(url.AsUri(), new StringContent(body, System.Text.Encoding.UTF8, "application/json"));

                response.EnsureSuccessStatusCode();

                return await ReadAs<T>(response);
            }
            catch (Exception e)
            {
                Log.For<Office365>().Error(e);
            }

            return default;
        }

        public static async Task SignOut()
        {
            foreach (var account in await _publicClientApplication.GetAccountsAsync())
                await _publicClientApplication.RemoveAsync(account);

            _tokenForUser = null;
        }

        static HttpClient CreateHttpClient()
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + _tokenForUser);
            return client;
        }

        static async Task<T> ReadAs<T>(HttpResponseMessage response)
        {
            if (typeof(T) == typeof(byte[]))
                return (T)(object)await response.Content.ReadAsByteArrayAsync();

            if (typeof(T) == typeof(Stream))
                return (T)(object)await response.Content.ReadAsStreamAsync();

            if (typeof(T) == typeof(string))
                return (T)(object)await response.Content.ReadAsStringAsync();

            return default;
        }
    }
}
