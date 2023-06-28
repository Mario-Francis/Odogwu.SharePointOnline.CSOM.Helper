using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace Odogwu.SharePointOnline.CSOM.Helper
{
    public class AuthenticationManager : IDisposable
    {
        private static readonly HttpClient httpClient = new HttpClient();

        // Token cache handling
        private static readonly SemaphoreSlim semaphoreSlimTokens = new SemaphoreSlim(1);
        private AutoResetEvent tokenResetEvent = null;
        private readonly ConcurrentDictionary<string, string> tokenCache = new ConcurrentDictionary<string, string>();
        private bool disposedValue;

        public Uri SiteUrl { get; set; }
        public string GrantType { get; set; }
        public string Resource { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string TenantName { get; set; }
        public int Timeout { get; set; } = 30000;

        public AuthenticationManager() { }

        public AuthenticationManager(string siteUrl, string grantType, string resource, string clientId, string clientSecret, string tenantName = null)
        {
            this.SiteUrl = new Uri(siteUrl);
            this.GrantType = grantType;
            this.Resource = resource;
            this.ClientId = clientId;
            this.ClientSecret = clientSecret;
            this.TenantName = tenantName;
        }

        internal class TokenWaitInfo
        {
            public RegisteredWaitHandle Handle = null;
        }

        public SP.ClientContext GetContext()
        {
            SP.ClientContext context = new SP.ClientContext(SiteUrl);
            context.ExecutingWebRequest += (sender, e) =>
            {
                string accessToken = EnsureAccessTokenAsync().GetAwaiter().GetResult();
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };
            context.RequestTimeout = Timeout;
            return context;
        }

        private async Task<string> EnsureAccessTokenAsync()
        {
            string accessTokenFromCache = TokenFromCache(SiteUrl, tokenCache);
            if (accessTokenFromCache == null)
            {
                await semaphoreSlimTokens.WaitAsync().ConfigureAwait(false);
                try
                {
                    // No async methods are allowed in a lock section
                    string accessToken = await AcquireTokenAsync().ConfigureAwait(false);
                    AddTokenToCache(SiteUrl, tokenCache, accessToken);

                    // Register a thread to invalidate the access token once's it's expired
                    tokenResetEvent = new AutoResetEvent(false);
                    TokenWaitInfo wi = new TokenWaitInfo();
                    wi.Handle = ThreadPool.RegisterWaitForSingleObject(
                        tokenResetEvent,
                        async (state, timedOut) =>
                        {
                            if (!timedOut)
                            {
                                TokenWaitInfo wi = (TokenWaitInfo)state;
                                if (wi.Handle != null)
                                {
                                    wi.Handle.Unregister(null);
                                }
                            }
                            else
                            {
                                try
                                {
                                    // Take a lock to ensure no other threads are updating the SharePoint Access token at this time
                                    await semaphoreSlimTokens.WaitAsync().ConfigureAwait(false);
                                    RemoveTokenFromCache(SiteUrl, tokenCache);
                                    Console.WriteLine($"Cached token for resource {SiteUrl.DnsSafeHost} and clientId {ClientId} expired");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Something went wrong during cache token invalidation: {ex.Message}");
                                    RemoveTokenFromCache(SiteUrl, tokenCache);
                                }
                                finally
                                {
                                    semaphoreSlimTokens.Release();
                                }
                            }
                        },
                        wi,
                        (uint)CalculateThreadSleep(accessToken).TotalMilliseconds,
                        true
                    );

                    return accessToken;
                }
                finally
                {
                    semaphoreSlimTokens.Release();
                }
            }
            else
                return accessTokenFromCache;
        }

        private async Task<string> AcquireTokenAsync()
        {
            var postData = new List<KeyValuePair<string, string>>();
            postData.Add(new KeyValuePair<string, string>("grant_type", GrantType));
            postData.Add(new KeyValuePair<string, string>("resource", $"{Resource}/{SiteUrl.DnsSafeHost}@{TenantName}"));
            postData.Add(new KeyValuePair<string, string>("client_id", $"{ClientId}@{TenantName}"));
            postData.Add(new KeyValuePair<string, string>("client_secret", $"{ClientSecret}"));

            using (var content = new FormUrlEncodedContent(postData))
            {
                content.Headers.Clear();
                content.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                var result = await httpClient.PostAsync($"https://accounts.accesscontrol.windows.net/{TenantName}/tokens/OAuth/2", content)
                    .ContinueWith((response) =>
                    {
                        return response.Result.Content.ReadAsStringAsync().Result;
                    })
                    .ConfigureAwait(false);

                var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
                var token = tokenResult.GetProperty("access_token").GetString();
                return token;
            }
        }

        private static string TokenFromCache(Uri web, ConcurrentDictionary<string, string> tokenCache)
        {
            if (tokenCache.TryGetValue(web.DnsSafeHost, out string accessToken))
                return accessToken;

            return null;
        }

        private static void AddTokenToCache(Uri web, ConcurrentDictionary<string, string> tokenCache, string newAccessToken)
        {
            if (tokenCache.TryGetValue(web.DnsSafeHost, out string currentAccessToken))
                tokenCache.TryUpdate(web.DnsSafeHost, newAccessToken, currentAccessToken);
            else
                tokenCache.TryAdd(web.DnsSafeHost, newAccessToken);
        }

        private static void RemoveTokenFromCache(Uri web, ConcurrentDictionary<string, string> tokenCache)
        {
            tokenCache.TryRemove(web.DnsSafeHost, out string currentAccessToken);
        }

        private static TimeSpan CalculateThreadSleep(string accessToken)
        {
            var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(accessToken);
            var lease = GetAccessTokenLease(token.ValidTo);
            lease = TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
            return lease;
        }

        private static TimeSpan GetAccessTokenLease(DateTime expiresOn)
        {
            DateTime now = DateTime.UtcNow;
            DateTime expires = expiresOn.Kind == DateTimeKind.Utc ? expiresOn : TimeZoneInfo.ConvertTimeToUtc(expiresOn);
            TimeSpan lease = expires - now;
            return lease;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (tokenResetEvent != null)
                    {
                        tokenResetEvent.Set();
                        tokenResetEvent.Dispose();
                    }
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
