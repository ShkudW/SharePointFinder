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
        public static string FormattedUpn { get; private set; } 

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
                Console.ForegroundColor = ConsoleColor.DarkBlue;
                Console.WriteLine($"[*] Using Refresh Token for requesting Access_Token for outlook.office365  api.....");
                Console.ResetColor();

                HttpResponseMessage response = await client.SendAsync(request);
                string jsonResponse = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[-] Failed to get Outlook access token. HTTP {response.StatusCode}");
                    Console.WriteLine(" ");
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
                        Console.WriteLine("[+] Successfully retrieved Outlook.office365.com Access Token!");
                        Console.WriteLine(" ");
                        Console.ResetColor();


                        return accessToken;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("[-] No access token Outlook.office365.com found.");
                        Console.WriteLine(" ");
                        Console.ResetColor();
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[-] Error: {ex.Message}");
                Console.WriteLine(" ");
                Console.ResetColor();
                return null;
            }
        }

        private static string DecodeAccessToken(string accessToken)
        {
            try
            {
                
                string payloadPart = accessToken.Split('.')[1];

                
                int mod4 = payloadPart.Length % 4;
                if (mod4 > 0)
                {
                    payloadPart += new string('=', 4 - mod4);
                }

                
                byte[] decodedBytes = Convert.FromBase64String(payloadPart);
                string jsonPayload = Encoding.UTF8.GetString(decodedBytes);

                
                using (JsonDocument doc = JsonDocument.Parse(jsonPayload))
                {
                    JsonElement root = doc.RootElement;

                    
                    if (root.TryGetProperty("upn", out JsonElement upnElement) ||
                        root.TryGetProperty("preferred_username", out upnElement))
                    {
                        string upn = upnElement.GetString();

                        if (!string.IsNullOrEmpty(upn))
                        {
                            
                            return upn.Replace("@", "_").Replace(".", "_").ToLower();
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[-] Error decoding Access Token: {ex.Message}");
                Console.ResetColor();
                return null;
            }
        }
    }
}
