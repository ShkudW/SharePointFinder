using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace SharePointFinder
{
    public class SharePointService
    {
        private static readonly HttpClient client = new HttpClient();

        public static async Task<List<string>> GetSharePointDomainsAsync(string accessToken)
        {
            string url = "https://webshell.suite.office.com/api/myapps/GetAppDataCache?hasMailboxInCloud=true&culture=en-US";

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Add("Authorization", $"Bearer {accessToken}");

            try
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("[+] Fetching SharePoint domains...");
                Console.ResetColor();

                HttpResponseMessage response = await client.SendAsync(request);
                string jsonResponse = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[-] Failed to get SharePoint domains. HTTP {response.StatusCode}");
                    Console.ResetColor();
                    return new List<string>();
                }

                if (string.IsNullOrEmpty(jsonResponse))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[-] Received an empty response from the server.");
                    Console.ResetColor();
                    return new List<string>();
                }

                List<string> sharePointDomains = new List<string>();

                using (JsonDocument doc = JsonDocument.Parse(jsonResponse)) // שימוש בבלוק using עבור C# 7.3
                {
                    JsonElement root = doc.RootElement;

                    if (!root.TryGetProperty("FirstParty", out JsonElement firstParty) ||
                        !firstParty.TryGetProperty("Apps", out JsonElement appsArray))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("[-] Unexpected JSON format, missing 'FirstParty.Apps'.");
                        Console.ResetColor();
                        return new List<string>();
                    }

                    foreach (JsonElement app in appsArray.EnumerateArray())
                    {
                        if (app.TryGetProperty("LaunchFullUrl", out JsonElement urlElement))
                        {
                            string launchUrl = urlElement.GetString();
                            if (!string.IsNullOrEmpty(launchUrl) && launchUrl.Contains("sharepoint.com"))
                            {
                                try
                                {
                                    string domain = new Uri(launchUrl).Host;
                                    if (!sharePointDomains.Contains(domain))
                                    {
                                        sharePointDomains.Add(domain);
                                    }
                                }
                                catch (UriFormatException)
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine($"[-] Invalid URL format: {launchUrl}");
                                    Console.ResetColor();
                                }
                            }
                        }
                    }
                }

                if (sharePointDomains.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("[-] No SharePoint domains found.");
                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"[+] Found {sharePointDomains.Count} SharePoint domains:");
                    foreach (string domain in sharePointDomains)
                    {
                        Console.WriteLine($"    - {domain}");
                    }
                    Console.ResetColor();
                }

                return sharePointDomains;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[-] Error: {ex.Message}");
                Console.ResetColor();
                return new List<string>();
            }
        }
    }
}
