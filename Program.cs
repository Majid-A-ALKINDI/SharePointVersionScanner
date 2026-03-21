using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;

class Program
{
    static async Task Main(string[] args)
    {
        PrintBanner();
        Console.WriteLine("Choose an option:");
        Console.WriteLine("1: Type the target URLs");
        Console.WriteLine("2: Select target URL file");
        Console.Write("Enter your choice (1 or 2): ");
        string choice = Console.ReadLine()?.Trim();

        List<string> hosts = new List<string>();

        if (choice == "1")
        {
            Console.WriteLine("Enter SharePoint URLs :");
            string url;
            while (!string.IsNullOrWhiteSpace(url = Console.ReadLine()))
            {
                hosts.Add(url.Trim());
            }
        }
        else if (choice == "2")
        {
            Console.Write("Enter the path to the file containing URLs: ");
            string filePath = Console.ReadLine()?.Trim();
            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
            {
                string content = File.ReadAllText(filePath);
                var urlRegex = new Regex(@"https?://[^\s]+", RegexOptions.IgnoreCase);
                var matches = urlRegex.Matches(content);
                hosts = matches.Select(m => m.Value.TrimEnd(',', '.', ';')).ToList();
            }
            else
            {
                Console.WriteLine("Invalid file path.");
                return;
            }
        }
        else
        {
            Console.WriteLine("Invalid choice.");
            return;
        }

        if (hosts.Count == 0)
        {
            Console.WriteLine("No hosts provided.");
            return;
        }

        Console.Write("Enable verbose output? (y/n): ");
        bool verbose = string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase);

        Console.WriteLine($"Found {hosts.Count} hosts to scan.");

        using var client = new HttpClient();

        for (int i = 0; i < hosts.Count; i++)
        {
            var host = hosts[i];
            Console.WriteLine($"\n[{i + 1}/{hosts.Count}] Scanning: {host}");

            var paths = new[] { "/admin/_vti_bin/client.svc/ProcessQuery", "/en/_vti_bin/client.svc/ProcessQuery", "/_vti_bin/client.svc/ProcessQuery" };
            bool versionFound = false;
            string version = null;
            string versionName = null;
            string errorMessage = null;
            string headerDetectedVersionName = null;

            foreach (var pathSuffix in paths)
            {
                var url = $"{host}{pathSuffix}";
                // CSOM query to get CompatibilityLevel of the current web
                var body = @"<Request xmlns=""http://schemas.microsoft.com/sharepoint/clientquery/2009"" SchemaVersion=""15.0.0.0"" LibraryVersion=""16.0.0.0"" ApplicationName=""SharePoint Version Scanner"">
<Actions>
<Query Id=""1"" ObjectPathId=""2"">
<Query SelectAllProperties=""false"">
<Properties>
<Property Name=""CompatibilityLevel"" ScalarProperty=""true"" />
</Properties>
</Query>
</Query>
</Actions>
<ObjectPaths>
<StaticProperty Id=""2"" TypeId=""{3747adcd-a3c3-41b9-bfab-4a64dd100d0a}"" Name=""Current"" />
</ObjectPaths>
</Request>";


                var content = new StringContent(body, Encoding.UTF8, "text/xml");

                var request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Content = content;
                request.Headers.Add("Accept", "*/*");
                request.Headers.Add("Accept-Language", "en-US,en;q=0.9");
                request.Headers.Add("Accept-Encoding", "gzip, deflate, br");
                request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36");
                request.Headers.Add("X-Requested-With", "XMLHttpRequest");
                request.Headers.Add("X-RequestDigest", "0x8786E8C5E485BE472EB68BF2AFACA9F12732793250D757A734030E67B9DAC83CEC3C91B0690F4ACD08F8DF3785A50541965AB012E4D41A0F9BDB293C82E83F77,21 Mar 2026 17:03:06 -0000");
                request.Headers.Add("Sec-Ch-Ua", "\"Chromium\";v=\"145\", \"Not:A-Brand\";v=\"99\"");
                request.Headers.Add("Sec-Ch-Ua-Platform", "\"Windows\"");
                request.Headers.Add("Sec-Ch-Ua-Mobile", "?0");
                request.Headers.Add("Sec-Fetch-Site", "same-origin");
                request.Headers.Add("Sec-Fetch-Mode", "cors");
                request.Headers.Add("Sec-Fetch-Dest", "empty");
                request.Headers.Add("Priority", "u=1, i");
                request.Headers.Add("Cookie", "ARRAffinity=310b2fcc84cfe1c4978678496139361c99638187673a89655f07acaba6eb74b0; ARRAffinitySameSite=310b2fcc84cfe1c4978678496139361c99638187673a89655f07acaba6eb74b0; x-bni-fpc=67ee3b0d1215d4ddcca54b27f034a687; x-bni-rncf=n0jRaMZzSY3m66c7QxMD9rWPIXaOnhUq5YTSxAX8G8M=; acceptCookies=true; CookieCheck1=true; CookieCheck2=false; ga-disable-UA-54601551-1=false; CookieCheck3=true; CookieCheck4=false; BNES_ga-disable-UA-54601551-1=j1hAUPjgxdMW0fHT45tJzHDQBw8SU/GBz6+XgzPPnAB97V/O304Z53W3RvF0Jmq5dUa61sJB/ZtQF9OrNtR3SA==; ARRAffinity=310b2fcc84cfe1c4978678496139361c99638187673a89655f07acaba6eb74b0; ARRAffinitySameSite=310b2fcc84cfe1c4978678496139361c99638187673a89655f07acaba6eb74b0; WSS_FullScreenMode=false; BNES_ARRAffinity=Lojy16Sj0wT8bcvcmlbnMxhJqFEIrN9fdu5GJu6VWgEemxbvL8Hiy56NsKCXQ34BD/oY32Fq1+IVi9wBVMm+Mw4GLeeYA0uuppRKW0XTrpnHhCnv4ToB/Hrl0XRgBbWy5RMh0Kuxr2U=; BNES_ARRAffinitySameSite=kviATctt7QYXlChgMXvjqP8xv6crXngHxM9rsagLl/iWE1ygK4pit7xf9YvR/RWSUtzD9PAJM4Y3URBiFMJai4YK0CZzC02lpXAxV83zMd7r5cQRHttwApsECG7dlgANb1BhA1JWhbIEfwVNnudJgw==");
                // Dynamic headers based on host
                request.Headers.Add("Host", new Uri(host).Host);
                request.Headers.Add("Origin", host);
                request.Headers.Add("Referer", $"{host}/Pages/default.aspx");

                if (verbose)
                {
                    Console.WriteLine($"Request URL: {url}");
                    Console.WriteLine("Request Method: POST");
                    Console.WriteLine($"Request Headers: {string.Join(", ", request.Headers.Select(h => $"{h.Key}: {string.Join(", ", h.Value)}"))}");
                }

                try
                {
                    Console.Write($"  Querying {pathSuffix}... ");
                    var startTime = DateTime.Now;
                    var responseTask = client.SendAsync(request);

                    // ASCII-safe ascending animation while waiting
                    int step = 0;
                    while (!responseTask.IsCompleted)
                    {
                        step++;
                        if (step > 5) step = 1;

                        var meter = new string('=', step) + new string('.', 5 - step);
                        var statusLine = $"  Querying {pathSuffix}... [{meter}]";
                        Console.Write("\r" + statusLine.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : statusLine.Length));
                        await Task.Delay(200);
                    }

                    var response = await responseTask;
                    var elapsed = (DateTime.Now - startTime).TotalMilliseconds;
                    var doneLine = $"  Response received ({elapsed:F0}ms)";
                    Console.Write("\r" + doneLine.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : doneLine.Length));
                    Console.WriteLine();

                    var responseBody = await response.Content.ReadAsStringAsync();

                    if (verbose)
                    {
                        Console.WriteLine($"Response Status: {response.StatusCode}");
                        Console.WriteLine($"Response Headers: {string.Join(", ", response.Headers.Select(h => $"{h.Key}: {string.Join(", ", h.Value)}"))}");
                        Console.WriteLine($"Response Body: {responseBody}");
                    }

                    // Try to parse LibraryVersion from response, even if not success
                    try
                    {
                        var jsonArray = JsonDocument.Parse(responseBody).RootElement;
                        if (jsonArray.GetArrayLength() > 0)
                        {
                            var metadata = jsonArray[0];
                            if (metadata.TryGetProperty("LibraryVersion", out var libVersion))
                            {
                                version = libVersion.GetString();
                                versionName = DetermineSharePointVersion(version);
                                versionFound = true;
                                break;
                            }
                        }
                    }
                    catch (JsonException)
                    {
                        // Not JSON, try to extract from headers
                    }

                    // If no version found in body, try to detect from response headers
                    if (version == null && response.Headers.Contains("SPRequestGuid"))
                    {
                        var detectedFromHeaders = ExtractVersionFromHeaders(response.Headers);
                        if (detectedFromHeaders != null)
                        {
                            // Keep header-based detection as fallback, but still try all paths.
                            headerDetectedVersionName ??= detectedFromHeaders;
                        }
                    }

                    // If not success and no version found, continue to next path
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                    continue;
                }
            }

            // Fallback: try SharePoint context info endpoint and parse d:LibraryVersion from XML.
            if (!versionFound)
            {
                var contextInfoUrl = $"{host}/_api/contextinfo";
                using var contextInfoRequest = new HttpRequestMessage(HttpMethod.Post, contextInfoUrl);
                contextInfoRequest.Content = new StringContent(string.Empty, Encoding.UTF8, "application/json");
                contextInfoRequest.Headers.Add("Accept", "application/xml");
                contextInfoRequest.Headers.Add("Host", new Uri(host).Host);
                contextInfoRequest.Headers.Add("Origin", host);
                contextInfoRequest.Headers.Add("Referer", $"{host}/Pages/default.aspx");

                if (verbose)
                {
                    Console.WriteLine($"Request URL: {contextInfoUrl}");
                    Console.WriteLine("Request Method: POST");
                    Console.WriteLine($"Request Headers: {string.Join(", ", contextInfoRequest.Headers.Select(h => $"{h.Key}: {string.Join(", ", h.Value)}"))}");
                }

                try
                {
                    Console.Write("  Querying /_api/contextinfo... ");
                    var startTime = DateTime.Now;
                    var responseTask = client.SendAsync(contextInfoRequest);

                    int step = 0;
                    while (!responseTask.IsCompleted)
                    {
                        step++;
                        if (step > 5) step = 1;

                        var meter = new string('=', step) + new string('.', 5 - step);
                        var statusLine = $"  Querying /_api/contextinfo... [{meter}]";
                        Console.Write("\r" + statusLine.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : statusLine.Length));
                        await Task.Delay(200);
                    }

                    var response = await responseTask;
                    var elapsed = (DateTime.Now - startTime).TotalMilliseconds;
                    var doneLine = $"  Response received ({elapsed:F0}ms)";
                    Console.Write("\r" + doneLine.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : doneLine.Length));
                    Console.WriteLine();

                    var responseBody = await response.Content.ReadAsStringAsync();

                    if (verbose)
                    {
                        Console.WriteLine($"Response Status: {response.StatusCode}");
                        Console.WriteLine($"Response Headers: {string.Join(", ", response.Headers.Select(h => $"{h.Key}: {string.Join(", ", h.Value)}"))}");
                        Console.WriteLine($"Response Body: {responseBody}");
                    }

                    var contextLibraryVersion = ExtractLibraryVersionFromContextInfo(responseBody);
                    if (!string.IsNullOrEmpty(contextLibraryVersion))
                    {
                        version = contextLibraryVersion;
                        versionName = DetermineSharePointVersion(version);
                        versionFound = true;
                    }
                    else if (response.Headers.Contains("SPRequestGuid"))
                    {
                        var detectedFromHeaders = ExtractVersionFromHeaders(response.Headers);
                        if (detectedFromHeaders != null)
                        {
                            headerDetectedVersionName ??= detectedFromHeaders;
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                }
            }

            // If no LibraryVersion was found in ProcessQuery or context info, use header-based detection.
            if (!versionFound && headerDetectedVersionName != null)
            {
                version = "(detected from headers)";
                versionName = headerDetectedVersionName;
                versionFound = true;
            }

            // After trying all paths, print the final result
            if (versionFound)
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("================================================================================================================================================");
                Console.Write($"Host: {host}, LibraryVersion: ");
                var originalColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write(version);
                Console.ForegroundColor = originalColor;
                Console.Write(", Version: ");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(versionName);
                Console.ForegroundColor = originalColor;
                Console.WriteLine();
                Console.WriteLine("================================================================================================================================================");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("================================================================================================================================================");
                Console.Write($"Host: {host}, Error: ");
                var originalColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write(errorMessage ?? "No SharePoint version detected");
                Console.ForegroundColor = originalColor;
                Console.WriteLine();
                Console.WriteLine("================================================================================================================================================");
            }
        }
    }

    static string DetermineSharePointVersion(string libraryVersion)
    {
        if (string.IsNullOrEmpty(libraryVersion)) return "Unknown";

        var parts = libraryVersion.Split('.');
        if (parts.Length < 2) return "Unknown";

        int major = int.Parse(parts[0]);
        int minor = int.Parse(parts[1]);

        return (major, minor) switch
        {
            (14, _) => "SharePoint 2010",
            (15, _) => "SharePoint 2013",
            (16, 0) => "SharePoint 2016/2019/Online/SE",
            (16, _) => "SharePoint Online or 2019",
            (17, _) => "SharePoint Server Subscription Edition",
            _ => $"Unknown (LibraryVersion: {libraryVersion})"
        };
    }

    static void PrintBanner()
    {
        Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
        Console.WriteLine("|");
        Console.WriteLine("|                                           SharePoint Version Scanner v1.2                                            |");
        Console.WriteLine("|                                                Built by Majed alkindi                                                |");
        Console.WriteLine("|");
        Console.WriteLine("|                              This script checks SharePoint endpoints and version hints                               |");
        Console.WriteLine("|                                       using ProcessQuery and response headers                                        |");
        Console.WriteLine("|");
        Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
    }

    static string ExtractVersionFromHeaders(HttpResponseHeaders headers)
    {
        // Check for X-SharePointHealthScore header (indicates SharePoint)
        if (headers.TryGetValues("X-SharePointHealthScore", out var values))
        {
            // If this header exists, it's SharePoint
            // Try to determine version from other hints
            if (headers.TryGetValues("X-MS-InvokeApp", out var invokeApp))
            {
                // Modern SharePoint Online typically has this header
                return "SharePoint Online (detected from X-MS-InvokeApp header)";
            }
        }

        // Check for SPRequestGuid (SharePoint GUID format)
        if (headers.TryGetValues("SPRequestGuid", out var guidValues))
        {
            var guidStr = guidValues.FirstOrDefault();
            if (guidStr != null && guidStr.Contains("-"))
            {
                // SPRequestGuid present indicates modern SharePoint
                return "SharePoint Online or Modern (detected from SPRequestGuid header)";
            }
        }

        return null;
    }

    static string ExtractLibraryVersionFromContextInfo(string xml)
    {
        if (string.IsNullOrWhiteSpace(xml))
        {
            return null;
        }

        var match = Regex.Match(xml, @"<d:LibraryVersion>([^<]+)</d:LibraryVersion>", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            return match.Groups[1].Value.Trim();
        }

        return null;
    }
}
