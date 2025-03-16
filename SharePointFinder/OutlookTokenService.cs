using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace SharePointFinder
{
    public class OutlookTokenService
    {
        private static readonly HttpClient client = new HttpClient();

        public static async Task<string> GetOutlookAccessTokenAsync(string refreshToken, string domainName)
        {
            
            string tenantId = await CheckTenantID.GetTenantID(domainName);

            if (string.IsNullOrEmpty(tenantId))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[-] Failed to retrieve Tenant ID for {domainName}");
                Console.ResetColor();
                return null;
            }

            
            string url = $"https://login.microsoftonline.com/{tenantId}/oauth2/token?api-version=1.0";

            var postData = new Dictionary<string, string>
            {
                { "client_id", "d3590ed6-52b3-4102-aeff-aad2292ab01c" },
                { "resource", "https://outlook.office365.com/" },
                { "grant_type", "refresh_token" },
                { "refresh_token", refreshToken },
                { "scope", "openid" }
            };

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new FormUrlEncodedContent(postData)
            };

            try
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"[+] Requesting Outlook Access Token for Tenant ID: {tenantId}");
                Console.ResetColor();

                HttpResponseMessage response = await client.SendAsync(request);
                string jsonResponse = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[-] Failed to get Outlook access token. HTTP {response.StatusCode}");
                    Console.ResetColor();
                    return null;
                }

                using (JsonDocument doc = JsonDocument.Parse(jsonResponse)) 
                {
                    JsonElement root = doc.RootElement;
                    if (root.TryGetProperty("access_token", out JsonElement tokenElement))
                    {
                        string accessToken = tokenElement.GetString();

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("[+] Successfully retrieved Outlook Access Token!");
                        Console.ResetColor();

                        return accessToken;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("[-] No access token found in response.");
                        Console.ResetColor();
                        return null;
                    }
                }
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
