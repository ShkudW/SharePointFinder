using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace SharePointFinder
{
    public class DeviceCodeFlow
    {
        private static readonly HttpClient client = new HttpClient();
        private const string ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
        private const string DeviceCodeUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode";
        private const string TokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        private const string RefreshTokenFile = "refresh_token.txt";

        public static async Task<string> GetDeviceCodeAsync()
        {
            var postData = new Dictionary<string, string>
            {
                { "client_id", ClientId },
                { "scope", "https://graph.microsoft.com/.default offline_access" }
            };

            HttpResponseMessage response = await client.PostAsync(DeviceCodeUrl, new FormUrlEncodedContent(postData));
            response.EnsureSuccessStatusCode();

            string jsonResponse = await response.Content.ReadAsStringAsync();
            using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
            {
                JsonElement root = doc.RootElement;

                string deviceCode = root.GetProperty("device_code").GetString();
                string userCode = root.GetProperty("user_code").GetString();
                string verificationUri = root.GetProperty("verification_uri").GetString();

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\n[+] Device Code Flow Started!");
                Console.WriteLine($"[+] User Code: {userCode}");
                Console.WriteLine($"[+] Verification URL: {verificationUri}");
                Console.WriteLine("[+] Please enter the code at the given URL.");
                Console.ResetColor();

                return deviceCode;
            }
        }

        public static async Task GetTokensAsync(string deviceCode)
        {
            var postData = new Dictionary<string, string>
            {
                { "client_id", ClientId },
                { "grant_type", "urn:ietf:params:oauth:grant-type:device_code" },
                { "device_code", deviceCode }
            };

            while (true)
            {
                try
                {
                    HttpResponseMessage response = await client.PostAsync(TokenUrl, new FormUrlEncodedContent(postData));
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                    {
                        JsonElement root = doc.RootElement;

                        if (root.TryGetProperty("error", out JsonElement error) && error.GetString() == "authorization_pending")
                        {
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine("Waiting for authorization... polling for Access Token.");
                            Console.ResetColor();
                            await Task.Delay(3000); 
                            continue;
                        }

                        string accessToken = root.GetProperty("access_token").GetString();
                        string refreshToken = root.GetProperty("refresh_token").GetString();

                        
                        Console.WriteLine("\n[+] Authorization Successful!");
                        Console.WriteLine(" ");
                        Console.ForegroundColor = ConsoleColor.DarkGreen;
                        Console.WriteLine("[+] Access Token:");
                        Console.WriteLine("-----------------------");
                        Console.ResetColor();
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine(accessToken);
                        Console.WriteLine(" ");
                        Console.ResetColor();

                        Console.ForegroundColor = ConsoleColor.DarkGreen;
                        Console.WriteLine("\n[+] Refresh Token:");
                        Console.WriteLine("-----------------------");
                        Console.ResetColor();
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine(refreshToken);
                        Console.ResetColor();
                      

                        SaveRefreshToken(refreshToken);
                        break;
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[-] Error: {ex.Message}");
                    Console.ResetColor();
                    break;
                }
            }
        }

        
        private static void SaveRefreshToken(string refreshToken)
        {
            File.WriteAllText(RefreshTokenFile, refreshToken);
            Console.WriteLine($"\n[+] Refresh Token saved to {RefreshTokenFile}");
           
        }
    }
}
