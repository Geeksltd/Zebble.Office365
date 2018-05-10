namespace Zebble
{
    using Microsoft.Identity.Client;
    using System;
    using System.Linq;
    using System.Net.Http;
    using System.Threading.Tasks;

    public class Office365
    {
#if ANDROID
        public static UIParent UIParent { get; set; }

        static Office365()
        {
            UIParent = new UIParent(UIRuntime.CurrentActivity);
        }
#endif
        public static string UserEmail { get; private set; }
        public static string UserName { get; private set; }
        public static string ClientId { get; set; }
        public static string[] Scopes { get; set; }

        static PublicClientApplication IdentityClientApp;
        static string TokenForUser = null;
        static DateTimeOffset Expiration;

        public static Task Initialize(string clientId, string[] scopes, string redirectURI)
        {
            IdentityClientApp = new PublicClientApplication(clientId) { RedirectUri = redirectURI };

            if (string.IsNullOrEmpty(clientId))
            {
                Device.Log.Error("Please set the clientId first!");
                return Task.CompletedTask;
            }

            ClientId = clientId;
            Scopes = scopes;

            return Task.CompletedTask;
        }

        public static async Task SignIn()
        {
            AuthenticationResult authResult;
            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes, IdentityClientApp.Users.First());
                TokenForUser = authResult.AccessToken;

                UserEmail = authResult.User.DisplayableId;
                UserName = authResult.User.Name;
            }
            catch
            {
                if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
#if ANDROID
                    authResult = await IdentityClientApp.AcquireTokenAsync(Scopes, UIParent);
#else
                    authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);
#endif
                    TokenForUser = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;

                    UserEmail = authResult.User.DisplayableId;
                    UserName = authResult.User.Name;
                }
            }

            await Task.CompletedTask;
        }

        public static async Task<object> GetRequest(string url, RequestResponseTypes requestResponse = RequestResponseTypes.String)
        {
            object result = null;
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + TokenForUser);
                    var endpoint = new Uri(url);
                    var response = await client.GetAsync(endpoint);

                    if (response.IsSuccessStatusCode)
                    {
                        switch (requestResponse)
                        {
                            default:
                                result = response.Content.ReadAsStringAsync();
                                break;
                            case RequestResponseTypes.Stream:
                                result = response.Content.ReadAsStreamAsync();
                                break;
                            case RequestResponseTypes.ByteArray:
                                result = response.Content.ReadAsByteArrayAsync();
                                break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Device.Log.Error(e.Message);
            }

            return result;
        }

        public static async Task<object> PostRequest(string url, string body, RequestResponseTypes requestResponse = RequestResponseTypes.String)
        {
            object result = null;
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + TokenForUser);
                    var endpoint = new Uri(url);
                    var response = await client.PostAsync(endpoint, new StringContent(body, System.Text.Encoding.UTF8, "application/json"));

                    if (response.IsSuccessStatusCode)
                    {
                        switch (requestResponse)
                        {
                            default:
                                result = response.Content.ReadAsStringAsync();
                                break;
                            case RequestResponseTypes.Stream:
                                result = response.Content.ReadAsStreamAsync();
                                break;
                            case RequestResponseTypes.ByteArray:
                                result = response.Content.ReadAsByteArrayAsync();
                                break;
                        }
                    }
                    else
                    {
                        Device.Log.Error("We could not send the message: " + response.StatusCode);
                    }
                }
            }
            catch (Exception e)
            {
                Device.Log.Error(e.Message);
            }

            return result;
        }

        public static void SignOut()
        {
            foreach (var user in IdentityClientApp.Users)
            {
                IdentityClientApp.Remove(user);
            }

            TokenForUser = null;
        }
    }
}
