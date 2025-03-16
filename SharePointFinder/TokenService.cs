using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace SharePointFinder
{
    public class TokenService
    {
        private static readonly HttpClient client = new HttpClient();

        public static async Task<string> GetAccessTokenAsync(string refreshToken, string domainName)
        {
            // קבלת ה-Tenant ID מהדומיין
            string tenantId = await CheckTenantID.GetTenantID(domainName);

            if (string.IsNullOrEmpty(tenantId))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[-] Failed to retrieve Tenant ID for {domainName}");
                Console.ResetColor();
                return null;
            }

            // URL עם ה-Tenant ID שהתקבל
            string url = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            var postData = new Dictionary<string, string>
            {
                { "client_id", "d3590ed6-52b3-4102-aeff-aad2292ab01c" },
                { "scope", "https://webshell.suite.office.com/.default" },
                { "grant_type", "refresh_token" },
                { "refresh_token", refreshToken }
            };

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new FormUrlEncodedContent(postData)
            };

            try
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"[+] Requesting access token for webshell.suite.office.com from Tenant ID: {tenantId}...");
                Console.ResetColor();

                HttpResponseMessage response = await client.SendAsync(request);
                string jsonResponse = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[-] Failed to get access token. HTTP {response.StatusCode}");
                    Console.ResetColor();
                    return null;
                }

                if (string.IsNullOrEmpty(jsonResponse))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[-] Received an empty response from the server.");
                    Console.ResetColor();
                    return null;
                }

                string accessToken = null;

                using (JsonDocument doc = JsonDocument.Parse(jsonResponse)) 
                {
                    JsonElement root = doc.RootElement;

                    if (root.TryGetProperty("access_token", out JsonElement tokenElement))
                    {
                        accessToken = tokenElement.GetString();
                    }
                }

                if (string.IsNullOrEmpty(accessToken))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[-] Access token not found in the response.");
                    Console.ResetColor();
                    return null;
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("[+] Access token retrieved successfully!");
                Console.ResetColor();

                return accessToken;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[-] Error: {ex.Message}");
                Console.ResetColor();
                return null;
            }
        }
    }
}
