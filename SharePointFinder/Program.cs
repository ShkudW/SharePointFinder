using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SharePointFinder
{
    class Program
    {
        static async Task Main(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return;
            }

            string command = args[0].ToLower();

            switch (command)
            {
                case "devicecodeflow":
                    await RunDeviceCodeFlow();
                    break;

                case "find":
                    await RunFindCommand(args);
                    break;

                default:
                    if (command.StartsWith("/domainname:"))
                        await GetTenantId(command.Substring(12));
                    else
                        PrintUsage();
                    break;
            }
        }

        
        private static async Task RunDeviceCodeFlow()
        {
            Console.WriteLine("[*] Starting Device Code Flow...");
            string deviceCode = await DeviceCodeFlow.GetDeviceCodeAsync();
            if (!string.IsNullOrEmpty(deviceCode))
            {
                await DeviceCodeFlow.GetTokensAsync(deviceCode);
            }
        }

        
        private static async Task RunFindCommand(string[] args)
        {
            string word = null, refreshToken = null, domainName = null;

            foreach (var arg in args)
            {
                if (arg.StartsWith("/word:", StringComparison.OrdinalIgnoreCase))
                    word = arg.Substring(6);
                if (arg.StartsWith("/refreshtoken:", StringComparison.OrdinalIgnoreCase))
                    refreshToken = arg.Substring(14);
                if (arg.StartsWith("/domainname:", StringComparison.OrdinalIgnoreCase))
                    domainName = arg.Substring(12);
            }

            if (string.IsNullOrEmpty(word) || string.IsNullOrEmpty(refreshToken) || string.IsNullOrEmpty(domainName))
            {
                Console.WriteLine("Usage: program.exe find /word:\"search_term\" /refreshtoken:\"token\" /domainname:\"yourdomain.com\"");
                return;
            }

            Console.WriteLine($"[*] Searching for: {word}");
            Console.WriteLine($"[*] Using Refresh Token: {refreshToken}");
            Console.WriteLine($"[*] Domain: {domainName}");

           
            string tenantId = await CheckTenantID.GetTenantID(domainName);
            if (string.IsNullOrEmpty(tenantId))
            {
                Console.WriteLine("[-] Failed to retrieve Tenant ID.");
                return;
            }

            
            string webshellAccessToken = await TokenService.GetAccessTokenAsync(refreshToken, tenantId);
            if (string.IsNullOrEmpty(webshellAccessToken))
            {
                Console.WriteLine("[-] Failed to retrieve webshell access token.");
                return;
            }

            
            List<string> sharePointDomains = await SharePointService.GetSharePointDomainsAsync(webshellAccessToken);
            if (sharePointDomains.Count == 0)
            {
                Console.WriteLine("[-] No SharePoint domains found.");
                return;
            }

            Console.WriteLine("[+] Found SharePoint domains:");
            foreach (var spDomain in sharePointDomains)
                Console.WriteLine($"    - {spDomain}");

            
            string outlookAccessToken = await OutlookTokenService.GetOutlookAccessTokenAsync(refreshToken, tenantId);
            if (string.IsNullOrEmpty(outlookAccessToken))
            {
                Console.WriteLine("[-] Failed to retrieve Outlook access token.");
                return;
            }

            
            await SharePointSearchService.SearchFilesAsync(outlookAccessToken, sharePointDomains, word);
        }

        
        private static async Task GetTenantId(string domainName)
        {
            if (string.IsNullOrEmpty(domainName))
            {
                Console.WriteLine("Usage: program.exe /domainname:yourdomain.com");
                return;
            }

            Console.WriteLine($"[*] Retrieving Tenant ID for domain: {domainName}");
            string tenantId = await CheckTenantID.GetTenantID(domainName);

            if (!string.IsNullOrEmpty(tenantId))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"[+] Tenant ID for {domainName}: {tenantId}");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[-] Failed to retrieve Tenant ID.");
            }
            Console.ResetColor();
        }

        
        private static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  program.exe devicecodeflow");
            Console.WriteLine("  program.exe find /word:\"search_term\" /refreshtoken:\"token\" /domainname:\"yourdomain.com\"");
            Console.WriteLine("  program.exe /domainname:yourdomain.com");
        }
    }
}
