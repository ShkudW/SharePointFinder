using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace SharePointFinder
{
    public class SharePointSearchService
    {
        private static readonly HttpClient client = new HttpClient();

        public static async Task SearchFilesAsync(string accessToken, List<string> sharePointDomains, string searchWord)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[-] Error: No Outlook Access Token provided.");
                Console.ResetColor();
                return;
            }

            if (sharePointDomains == null || sharePointDomains.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("[-] Error: No SharePoint domains found.");
                Console.ResetColor();
                return;
            }

            string searchUrl = "https://outlook.office365.com/searchservice/api/v2/query";

            foreach (var domain in sharePointDomains)
            {
                string jsonBody = $@"{{
                  ""AnswerEntityRequest"": [
                        {{
                          ""Query"": {{
                             ""QueryString"": ""\""{searchWord}\"""",
                          }},
                          ""EntityTypes"": [
                            ""Building"",
                            ""EditorialQnA"",
                            ""Bookmark"",
                            ""People"",
                            ""Acronym"",
                            ""External"",
                            ""TuringQnA"",
                            ""Topic""
                          ],
                          ""From"": 0,
                          ""Size"": 10,
                          ""SupportedResultSourceFormats"": [
                            ""AdaptiveCard"",
                            ""EntityData"",
                            ""AdaptiveCardTemplateBinding""
                          ],
                          ""PreferredResultSourceFormat"": ""AdaptiveCard"",
                          ""EnableAsyncResolution"": true
                        }}
                      ],
                      ""EntityRequests"": [
                        {{
                          ""EntityType"": ""File"",
                          ""ContentSources"": [
                            ""SharePoint"",
                            ""OneDriveBusiness""
                          ],
                          ""Fields"": [
                            "".callerStack"",
                            "".correlationId"",
                            "".mediaBaseUrl"",
                            "".spResourceUrl"",
                            "".thumbnailUrl"",
                            ""AuthorOWSUSER"",
                            ""ContainerTypeId"",
                            ""ContentClass"",
                            ""ContentTypeId"",
                            ""Created"",
                            ""DefaultEncodingURL"",
                            ""DepartmentId"",
                            ""Description"",
                            ""DocId"",
                            ""EditorOWSUSER"",
                            ""FileExtension"",
                            ""Filename"",
                            ""GeoLocationSource"",
                            ""HitHighlightedSummary"",
                            ""IsContainer"",
                            ""IsHubSite"",
                            ""LastModifiedTime"",
                            ""LinkingUrl"",
                            ""ListID"",
                            ""ModifiedBy"",
                            ""MediaDuration"",
                            ""ParentLink"",
                            ""Path"",
                            ""PiSearchResultId"",
                            ""PictureThumbnailURL"",
                            ""ProgID"",
                            ""PromotedState"",
                            ""RelatedHubSites"",
                            ""SPWebUrl"",
                            ""SecondaryFileExtension"",
                            ""ServerRedirectedPreviewURL"",
                            ""ServerRedirectedUrl"",
                            ""ShortcutUrl"",
                            ""SiteId"",
                            ""SiteLogo"",
                            ""SiteTemplateId"",
                            ""SiteTitle"",
                            ""Title"",
                            ""UniqueID"",
                            ""UniqueId"",
                            ""ViewCount"",
                            ""ViewsLifeTimeUniqueUsers"",
                            ""WebId"",
                            ""isDocument"",
                            ""isexternalcontent"",
                            ""ListTemplateTypeId"",
                            ""PrivacyIndicator"",
                            ""ColorHex"",
                            ""ModifierUPNs"",
                            ""InformationProtectionLabelId"",
                            ""SiteSensitivityLabelID""
                          ],
                          ""Query"": {{
                            ""QueryString"": ""\""{searchWord}\"""",
                            ""DisplayQueryString"": ""\""{searchWord}\"""",
                            ""QueryTemplate"": """"
                          }},
                          ""Sort"": [
                            {{
                              ""Field"": ""PersonalScore"",
                              ""SortDirection"": ""Desc""
                            }}
                          ],
                          ""EnableQueryUnderstanding"": false,
                          ""EnableSpeller"": false,
                          ""IdFormat"": 0,
                          ""EnableResultAnnotations"": true,
                          ""ResultsMerge"": {{
                            ""Type"": ""Interleaved""
                          }},
                          ""FederationContext"": {{
                            ""SpoFederationContext"": {{
                              ""UserContextUrl"": ""https://{domain}/search/""
                            }}
                          }},
                          ""ExtendedQueries"": [
                            {{
                              ""SearchProvider"": ""SharePoint""
                            }}
                          ],
                          ""HitHighlight"": {{
                            ""HitHighlightedProperties"": [
                              """"
                            ],
                            ""SummaryLength"": 200
                          }}
                        }}
                      ],
                      ""Cvid"": ""11111111-1111-1111-1111-111111111111"",
                      ""LogicalId"": ""11111111-1111-1111-1111-111111111111"",
                      ""Culture"": ""en-us"",
                      ""UICulture"": ""en-us"",
                      ""TimeZone"": ""UTC"",
                      ""TextDecorations"": ""Off"",
                      ""Scenario"": {{
                        ""Name"": ""officehome"",
                        ""Dimensions"": [
                          {{
                            ""DimensionName"": ""QueryType"",
                            ""DimensionValue"": ""AllResults""
                          }},
                          {{
                            ""DimensionName"": ""FormFactor"",
                            ""DimensionValue"": ""Web""
                          }}
                        ]
                      }},
                      ""QueryAlterationOptions"": {{
                        ""EnableSuggestion"": true,
                        ""EnableAlteration"": true,
                        ""SupportedRecourseDisplayTypes"": [
                          ""ServiceSideRecourseLink""
                        ]
                      }},
                      ""WholePageRankingOptions"": {{
                        ""EnableEnrichedRanking"": true,
                        ""EnableLayoutHints"": true,
                        ""SupportedSerpRegions"": [
                          ""MainLine""
                        ],
                        ""EntityResultTypeRankingOptions"": [
                          {{
                            ""ResultType"": ""Answer"",
                            ""MaxEntitySetCount"": 6
                          }}
                        ],
                        ""MultiEntityMerge"": [
                          {{
                            ""EntityTypes"": [
                              ""File"",
                              ""External""
                            ],
                            ""Size"": 15,
                            ""From"": 0
                          }}
                        ],
                        ""SupportedRankingVersion"": 1
                      }}
                    }}";

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, searchUrl)
                {
                    Content = new StringContent(jsonBody, Encoding.UTF8, "application/json")
                };

                request.Headers.Add("Authorization", $"Bearer {accessToken}");
                request.Headers.Add("Accept", "application/json");

                try
                {
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine(" ");
                    Console.WriteLine("------------------------------------------------------------------------");
                    Console.WriteLine($"[+] Searching in: {domain}...");
                    Console.ResetColor();

                    HttpResponseMessage response = await client.SendAsync(request);
                    string jsonResponse = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"[-] Failed to search files in {domain}. HTTP {response.StatusCode}");
                        Console.ResetColor();
                        continue;
                    }

                    using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                    {
                        JsonElement root = doc.RootElement;
                        List<string> urls = new List<string>();

                        if (root.TryGetProperty("EntitySets", out JsonElement entitySets))
                        {
                            foreach (JsonElement entitySet in entitySets.EnumerateArray())
                            {
                                if (entitySet.TryGetProperty("ResultSets", out JsonElement resultSets))
                                {
                                    foreach (JsonElement resultSet in resultSets.EnumerateArray())
                                    {
                                        if (resultSet.TryGetProperty("Results", out JsonElement results))
                                        {
                                            foreach (JsonElement result in results.EnumerateArray())
                                            {
                                                if (result.TryGetProperty("Source", out JsonElement source))
                                                {
                                                    
                                                    string[] urlKeys = {
                                                        "LinkingUrl",
                                                        "Path",
                                                        "SPWebUrl",
                                                        "ServerRedirectedUrl",
                                                        "ServerRedirectedPreviewURL",
                                                        ".spResourceUrl"
                                                    };

                                                    foreach (var key in urlKeys)
                                                    {
                                                        if (source.TryGetProperty(key, out JsonElement urlElement))
                                                        {
                                                            string url = urlElement.GetString();
                                                            if (!string.IsNullOrEmpty(url))
                                                            {
                                                                urls.Add(url);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.WriteLine(" ");
                        Console.WriteLine($"[+] Total results found in {domain}: {urls.Count}");
                        Console.ResetColor();
                        Console.WriteLine(" ");


                        foreach (var url in urls)
                        {
                            Console.ForegroundColor = ConsoleColor.DarkCyan;
                            Console.WriteLine($"[URL]: {url}");
                            Console.ResetColor();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[-] Error: {ex.Message}");
                    Console.ResetColor();
                }
            }
        }
    }
}
