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

            PrintBanner();

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
                Console.WriteLine("Usage: SharePointFinder.exe find /word: /RefreshToken: /DomainNname:");
                return;
            }

            

            string tenantId = await CheckTenantID.GetTenantID(domainName);
            if (string.IsNullOrEmpty(tenantId))
            {
                Console.WriteLine("[-] Domain Name not found ");
                return;
            }

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(" ");
            Console.WriteLine($"[+] Searching for: {word}");
            Console.WriteLine($"[+] Using Refresh Token from user input");
            Console.WriteLine($"[+] Tenant target Name: {domainName}");
            Console.WriteLine($"[+] Tenant target ID: {tenantId}");
            Console.WriteLine(" ");
            Console.ResetColor();


            string webshellAccessToken = await TokenService.GetAccessTokenAsync(refreshToken, tenantId);
            if (string.IsNullOrEmpty(webshellAccessToken))
            {
                Console.WriteLine("[-] Failed to retrieve webshell.suite.com Access Token.");
                return;
            }


            List<string> sharePointDomains = await SharePointService.GetSharePointDomainsAsync(webshellAccessToken);
            if (sharePointDomains.Count == 0)
            {
                Console.WriteLine("[-] Share Point Domain Name not found");
                return;
            }

            
            string outlookAccessToken = await OutlookTokenService.GetOutlookAccessTokenAsync(refreshToken, tenantId);
            if (string.IsNullOrEmpty(outlookAccessToken))
            {
                Console.WriteLine("[-] Failed to retrieve Outlook Access Token.");
                return;
            }

            
            await SharePointSearchService.SearchFilesAsync(outlookAccessToken, sharePointDomains, word);
           
        }

        
        private static async Task GetTenantId(string domainName)
        {
            if (string.IsNullOrEmpty(domainName))
            {
                Console.WriteLine("Usage: SharePointFinder.exe /DomainName:");
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

        private static void PrintBanner()
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(@"
 ____  _                    ____       _       _    __ _           _           
/ ___|| |__   __ _ _ __ ___|  _ \ ___ (_)_ __ | |_ / _(_)_ __   __| | ___ _ __ 
\___ \| '_ \ / _` | '__/ _ \ |_) / _ \| | '_ \| __| |_| | '_ \ / _` |/ _ \ '__|
 ___) | | | | (_| | | |  __/  __/ (_) | | | | | |_|  _| | | | | (_| |  __/ |   
|____/|_| |_|\__,_|_|  \___|_|   \___/|_|_| |_|\__|_| |_|_| |_|\__,_|\___|_|   
                                                                              
       SharePoint Search Tool - By ShkudW
       https://github.com/ShkudW/SharePointFinder
");
            Console.ResetColor();
        }
        private static void PrintUsage()
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("Usage:");
            Console.WriteLine("------");
            Console.WriteLine(" ");
            Console.WriteLine("Getting Access_Token and Refresh_Token from Device Code Flow:");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("     [#] SharePointFinder.exe DeviceCodeFlow");
            Console.ResetColor();
            Console.WriteLine(" ");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("Getting Tenant ID:");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("     [#] SharePointFinder.exe /DomainName:");
            Console.ResetColor();
            Console.WriteLine(" ");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("Find interesting files in SharePoint:");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("     [#] SharePointFinder.exe find /word: /RefreshToken: /DomainName:");
            Console.WriteLine(" ");
            Console.ResetColor();

        }
    }
}
