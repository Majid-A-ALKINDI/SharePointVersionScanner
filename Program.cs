using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;
using System.Net;

class Program
{
    static async Task Main(string[] args)
    {
        PrintBanner();
        await FetchAndPrintLatestUpdates();
        while (true)
        {
            Console.WriteLine("Choose an option:");
            WriteMenuOption("1", "Type the target URLs", ConsoleColor.Cyan);
            WriteMenuOption("2", "Select target URL file", ConsoleColor.Cyan);
            WriteMenuOption("3", "Filter output file by SharePoint version", ConsoleColor.Cyan);
            WriteMenuOption("0", "Exit", ConsoleColor.DarkYellow);
            Console.Write("Enter your choice (0, 1, 2, or 3): ");
            string choice = Console.ReadLine()?.Trim();

            if (choice == "0")
            {
                Console.WriteLine("Exiting...");
                break;
            }

            if (choice == "3")
            {
                Console.Write("Enter the path to the output file (default: output.txt): ");
                string outputFilePath = Console.ReadLine()?.Trim();
                if (string.IsNullOrWhiteSpace(outputFilePath))
                    outputFilePath = "output.txt";

                outputFilePath = Path.GetFullPath(outputFilePath);

                if (!File.Exists(outputFilePath))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"File not found: {outputFilePath}");
                    Console.ResetColor();
                    continue;
                }

            Console.WriteLine("Select filter type:");
            WriteMenuOption("1", "SharePoint Server Subscription Edition", ConsoleColor.Green);
            WriteMenuOption("2", "SharePoint Server 2019", ConsoleColor.Green);
            WriteMenuOption("3", "SharePoint Server 2016", ConsoleColor.Green);
            WriteMenuOption("4", "Error or Unknown host", ConsoleColor.Red);
            Console.Write("Enter your choice (1, 2, 3, or 4): ");
            string versionChoice = Console.ReadLine()?.Trim();

            string filterKeyword = versionChoice switch
            {
                "1" => "Subscription",
                "2" => "2019",
                "3" => "2016",
                "4" => "ERROR_OR_UNKNOWN",
                _ => null
            };

                if (filterKeyword == null)
                {
                    Console.WriteLine("Invalid choice.");
                    continue;
                }

            string versionLabel = versionChoice switch
            {
                "1" => "SharePoint Server Subscription Edition",
                "2" => "SharePoint Server 2019",
                "3" => "SharePoint Server 2016",
                "4" => "Error or Unknown host",
                _ => ""
            };

                bool useLessThanFilter = false;
                Version thresholdVersion = null;
                if (versionChoice is "1" or "2" or "3")
                {
                    Console.WriteLine("Choose action:");
                    WriteMenuOption("1", "Print all matching versions", ConsoleColor.Cyan);
                    WriteMenuOption("2", "Look for versions less than a value", ConsoleColor.Yellow);
                    Console.Write("Enter your choice (1 or 2): ");
                    string actionChoice = Console.ReadLine()?.Trim();

                    if (actionChoice == "2")
                    {
                        Console.Write("Enter version threshold (example: 16.0.10417.20027): ");
                        string thresholdText = Console.ReadLine()?.Trim();
                        if (!TryParseVersion(thresholdText, out thresholdVersion))
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Invalid version format. Use numbers like 16.0.10417.20027");
                            Console.ResetColor();
                            continue;
                        }
                        useLessThanFilter = true;
                    }
                    else if (actionChoice != "1")
                    {
                        Console.WriteLine("Invalid choice.");
                        continue;
                    }
                }

                var lines = File.ReadAllLines(outputFilePath);
                var matches = lines
                    .Where(l => l.StartsWith("Host:", StringComparison.OrdinalIgnoreCase)
                             && (filterKeyword == "ERROR_OR_UNKNOWN"
                                 ? l.Contains("Error:", StringComparison.OrdinalIgnoreCase)
                                   || l.Contains("Unknown", StringComparison.OrdinalIgnoreCase)
                                 : l.Contains(filterKeyword, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                if (useLessThanFilter)
                {
                    matches = matches
                        .Where(l => IsVersionLessThan(l, thresholdVersion))
                        .ToList();
                }

                Console.WriteLine();
                Console.WriteLine($"Filtering: {outputFilePath}");
                Console.WriteLine($"Results for: {versionLabel}");
                if (useLessThanFilter)
                {
                    Console.WriteLine($"Condition: LibraryVersion < {thresholdVersion}");
                }
                Console.WriteLine(new string('=', 120));

                if (matches.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("No matching entries found.");
                    Console.ResetColor();
                }
                else
                {
                    foreach (var line in matches)
                    {
                        PrintColoredFileResultLine(line);
                    }
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"Total matches: {matches.Count}");
                    Console.ResetColor();
                }

                Console.WriteLine(new string('=', 120));
                Console.WriteLine();
                continue;
            }

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
                    hosts = ExtractHttpsHostsFromText(content);
                }
                else
                {
                    Console.WriteLine("Invalid file path.");
                    continue;
                }
            }
            else
            {
                Console.WriteLine("Invalid choice.");
                continue;
            }

            if (hosts.Count == 0)
            {
                Console.WriteLine("No hosts provided.");
                continue;
            }

        Console.Write("Enable verbose output? (y/n): ");
        bool verbose = string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase);

        Console.Write("Ignore SSL certificate errors? (y/n): ");
        bool ignoreSslErrors = string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase);

        Console.Write("Save output to file? (y/n): ");
        bool saveOutput = string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase);
        string outputPath = string.Empty;
        var outputLines = new List<string>();

        if (saveOutput)
        {
            Console.Write("Enter output file path (default: output.txt): ");
            outputPath = Console.ReadLine()?.Trim();
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                outputPath = "output.txt";
            }

            outputPath = Path.GetFullPath(outputPath);
            Console.WriteLine($"Output will be saved to: {outputPath}");

            outputLines.Add("SharePoint Version Scanner Results");
            outputLines.Add($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            outputLines.Add(new string('=', 120));
        }

        Console.WriteLine($"Found {hosts.Count} hosts to scan.");

        using var clientHandler = new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli
        };
        if (ignoreSslErrors)
        {
            clientHandler.ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;
        }
        using var client = new HttpClient(clientHandler);

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
                    errorMessage = GetDetailedErrorMessage(ex);
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
                    errorMessage = GetDetailedErrorMessage(ex);
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

                if (saveOutput)
                {
                    outputLines.Add($"Host: {host}, LibraryVersion: {version}, Version: {versionName}");
                }
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

                if (saveOutput)
                {
                    outputLines.Add($"Host: {host}, Error: {errorMessage ?? "No SharePoint version detected"}");
                }
            }
        }

            if (saveOutput)
            {
                try
                {
                    File.WriteAllLines(outputPath, outputLines);
                    Console.WriteLine($"Output saved to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to save output file: {ex.Message}");
                    Console.ResetColor();
                }
            }

            Console.WriteLine();
            Console.WriteLine("Scan completed. Returning to main menu...");
            Console.WriteLine();
        }
    }

    static async Task FetchAndPrintLatestUpdates()
    {
        try
        {
            using var http = new HttpClient();
            http.Timeout = TimeSpan.FromSeconds(15);

            using var req = new HttpRequestMessage(HttpMethod.Get,
                "https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates");
            req.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
            req.Headers.Add("Accept", "text/html,application/xhtml+xml");
            req.Headers.Add("Accept-Language", "en-US,en;q=0.9");

            var fetchTask = http.SendAsync(req);

            int step = 0;
            while (!fetchTask.IsCompleted)
            {
                step++;
                if (step > 5) step = 1;
                var meter = new string('=', step) + new string('.', 5 - step);
                var statusLine = $"  Fetching latest SharePoint updates... [{meter}]";
                Console.Write("\r" + statusLine.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : statusLine.Length));
                await Task.Delay(200);
            }

            var res = await fetchTask;
            var html = await res.Content.ReadAsStringAsync();

            var doneLine = "  Fetching latest SharePoint updates... done.";
            Console.Write("\r" + doneLine.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : doneLine.Length));
            Console.WriteLine();
            Console.WriteLine();

            var sections = new[]
            {
                ("SharePoint Server Subscription Edition update history", "SharePoint Server Subscription Edition"),
                ("SharePoint 2019 update history", "SharePoint Server 2019"),
                ("SharePoint 2016 update history", "SharePoint Server 2016")
            };

            Console.WriteLine("  Latest SharePoint Updates  (source: learn.microsoft.com/en-us/officeupdates/sharepoint-updates)");
            Console.WriteLine("  " + new string('-', 97));
            Console.WriteLine("  " + "Product".PadRight(42) + "KB".PadRight(14) + "Version".PadRight(22) + "Date");
            Console.WriteLine("  " + new string('-', 97));

            bool anyFound = false;
            foreach (var (sectionTitle, label) in sections)
            {
                var sectionIdx = html.IndexOf(sectionTitle, StringComparison.OrdinalIgnoreCase);
                if (sectionIdx < 0) continue;

                var slice = html.Substring(sectionIdx);

                var kbMatch = Regex.Match(slice, @"KB\s+(\d+)", RegexOptions.IgnoreCase);
                if (!kbMatch.Success) continue;

                var afterKb = slice.Substring(kbMatch.Index);
                var versionMatch = Regex.Match(afterKb, @"(\d+\.\d+\.\d+\.\d+)");
                if (!versionMatch.Success) continue;

                var afterVersion = afterKb.Substring(versionMatch.Index);
                var dateMatch = Regex.Match(afterVersion,
                    @"((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+(?:\d+,\s+\d{4}|\d{4}))");
                if (!dateMatch.Success) continue;

                var kb = "KB " + kbMatch.Groups[1].Value;
                var version = versionMatch.Groups[1].Value;
                var date = dateMatch.Groups[1].Value.Trim();

                var orig = Console.ForegroundColor;
                Console.Write("  " + label.PadRight(42));
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write(kb.PadRight(14));
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write(version.PadRight(22));
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(date);
                Console.ForegroundColor = orig;
                Console.WriteLine();
                anyFound = true;
            }

            if (!anyFound)
                Console.WriteLine("  (No update data found — page structure may have changed)");

            Console.WriteLine("  " + new string('-', 97));
            Console.WriteLine();
        }
        catch (TaskCanceledException)
        {
            var msg = "  Fetching latest SharePoint updates... (timed out, skipping)";
            Console.Write("\r" + msg.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : msg.Length));
            Console.WriteLine();
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            var msg = $"  Fetching latest SharePoint updates... (skipped: {ex.Message})";
            Console.Write("\r" + msg.PadRight(Console.WindowWidth > 0 ? Console.WindowWidth - 1 : msg.Length));
            Console.WriteLine();
            Console.WriteLine();
        }
    }

    static List<string> ExtractHttpsHostsFromText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return new List<string>();
        }

        // Extract domains or IPv4 values and normalize them to https://host
        var hostRegex = new Regex(@"((?:\d{1,3}\.){3}\d{1,3}|(?:[a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})",
            RegexOptions.IgnoreCase);

        var hosts = hostRegex.Matches(text)
            .Select(m => m.Groups[1].Value.Trim().TrimEnd(',', '.', ';'))
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Select(h => $"https://{h}")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return hosts;
    }

    static string GetDetailedErrorMessage(Exception ex)
    {
        if (ex == null)
        {
            return "Unknown error";
        }

        if (ex.InnerException != null && !string.IsNullOrWhiteSpace(ex.InnerException.Message))
        {
            return $"{ex.Message} (Inner: {ex.InnerException.Message})";
        }

        return ex.Message;
    }

    static string DetermineSharePointVersion(string libraryVersion)
    {
        if (string.IsNullOrEmpty(libraryVersion)) return "Unknown";

        var parts = libraryVersion.Split('.');
        if (parts.Length < 2) return "Unknown";

        if (!int.TryParse(parts[0], out int major) || !int.TryParse(parts[1], out int minor))
            return $"Unknown (LibraryVersion: {libraryVersion})";

        if (major == 14) return "SharePoint Server 2010";
        if (major == 15) return "SharePoint Server 2013";

        if (major == 16 && minor == 0 && parts.Length >= 3)
        {
            var build = parts[2];
            // Identify SharePoint 16.0 by the build number's prefix
            // 16.0.19xxx => SharePoint Server Subscription Edition
            // 16.0.10xxx => SharePoint Server 2019
            // 16.0.5xxx  => SharePoint Server 2016
            if (build.StartsWith("19")) return "SharePoint Server Subscription Edition";
            if (build.StartsWith("10")) return "SharePoint Server 2019";
            if (build.StartsWith("5"))  return "SharePoint Server 2016";
            return $"SharePoint 16.0 (build: {libraryVersion})";
        }

        if (major == 17) return "SharePoint Server Subscription Edition";

        return $"Unknown (LibraryVersion: {libraryVersion})";
    }

    static void PrintBanner()
    {
        var originalColor = Console.ForegroundColor;

        Console.ForegroundColor = ConsoleColor.DarkCyan;
        Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
        Console.WriteLine("|");

        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("|                                           SharePoint Version Scanner v1.2                                            |");

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("|                                                 Built by Majid alkindi                                               |");

        Console.ForegroundColor = ConsoleColor.DarkCyan;
        Console.WriteLine("|");

        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine("|                              This script checks SharePoint endpoints and version hints                               |");
        Console.WriteLine("|                                       using ProcessQuery and response headers                                        |");

        Console.ForegroundColor = ConsoleColor.DarkCyan;
        Console.WriteLine("|");
        Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");

        Console.ForegroundColor = originalColor;
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

    static bool IsVersionLessThan(string outputLine, Version threshold)
    {
        if (threshold == null || string.IsNullOrWhiteSpace(outputLine))
        {
            return false;
        }

        var match = Regex.Match(outputLine, @"LibraryVersion:\s*([^,]+)", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return false;
        }

        var versionText = match.Groups[1].Value.Trim();
        if (!TryParseVersion(versionText, out var lineVersion))
        {
            return false;
        }

        return lineVersion < threshold;
    }

    static bool TryParseVersion(string value, out Version version)
    {
        version = null;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var match = Regex.Match(value.Trim(), @"^\d+\.\d+\.\d+\.\d+$");
        if (!match.Success)
        {
            return false;
        }

        return Version.TryParse(match.Value, out version);
    }

    static void WriteMenuOption(string key, string label, ConsoleColor textColor)
    {
        var originalColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.Write($"{key}: ");
        Console.ForegroundColor = textColor;
        Console.WriteLine(label);
        Console.ForegroundColor = originalColor;
    }

    static void PrintColoredFileResultLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return;
        }

        var originalColor = Console.ForegroundColor;

        if (line.Contains("Error:", StringComparison.OrdinalIgnoreCase)
            || line.Contains("Unknown", StringComparison.OrdinalIgnoreCase))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(line);
            Console.ForegroundColor = originalColor;
            return;
        }

        var hostMatch = Regex.Match(line, @"Host:\s*([^,]+)", RegexOptions.IgnoreCase);
        var libMatch = Regex.Match(line, @"LibraryVersion:\s*([^,]+)", RegexOptions.IgnoreCase);
        var versionMatch = Regex.Match(line, @"Version:\s*(.+)$", RegexOptions.IgnoreCase);

        if (hostMatch.Success || libMatch.Success || versionMatch.Success)
        {
            Console.Write("Host: ");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write(hostMatch.Success ? hostMatch.Groups[1].Value.Trim() : "N/A");
            Console.ForegroundColor = originalColor;

            Console.Write(", LibraryVersion: ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write(libMatch.Success ? libMatch.Groups[1].Value.Trim() : "N/A");
            Console.ForegroundColor = originalColor;

            Console.Write(", Version: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(versionMatch.Success ? versionMatch.Groups[1].Value.Trim() : "N/A");
            Console.ForegroundColor = originalColor;
            return;
        }

        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine(line);
        Console.ForegroundColor = originalColor;
    }
}
