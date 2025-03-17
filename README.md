# SharePointFinder

A simple tool for finding sensitive information in SharePoint.


Usage:

1. First use this function for getting refresh_token from device code flow api.
```powershell

SharePointFinder.exe DeviceCodeFlow
```

2. Now, use the refresh token for finding sensitive information:
```powershell

SharePointFinder.exe find /word:"Aa123456" /domain:domain.local /refreshtoken:1.AQQAGUvwznZ3lEq4....
```


Explain:
* The tool will use the refresh token to request an access token for the webshell.suite.com api.
* With the webshell.suite.com access token the tool will find the sharepoint domain names (it can be a different domain name than the tenant namd).
* The tool will use the refresh token to request an access token for the outlook.office365.com API
* With the outlook.office365.com access token the tool will send a web request (based on json format) that includes the sharepoint domain name and the word you want to search for.


![image](https://github.com/user-attachments/assets/1aadb457-dc61-4c0f-a4a8-5ec9e6e9f7f8)
