using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;

namespace SharePointFinder
{
    public class CheckTenantID
    {
        private static readonly HttpClient client = new HttpClient();

        public static async Task<string> GetTenantID(string domain)
        {
            string url = $"https://login.microsoftonline.com/{domain}/v2.0/.well-known/openid-configuration";

            try
            {
                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string jsonResponse = await response.Content.ReadAsStringAsync();
                JObject json = JObject.Parse(jsonResponse);

                string issuer = json["issuer"]?.ToString();
                if (!string.IsNullOrEmpty(issuer))
                {
                    string tenantId = issuer.Replace("https://login.microsoftonline.com/", "").Replace("/v2.0", "");
                    return tenantId;
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[-] Error: {ex.Message}");
                Console.ResetColor();
            }

            return null;
        }
    }
}
