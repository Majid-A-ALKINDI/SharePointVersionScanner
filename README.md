# SharePoint Version Scanner

A C# console application that scans a list of SharePoint hosts to determine their version by making POST requests to `/en/_vti_bin/client.svc/ProcessQuery`.

## Usage

Run the application: `dotnet run`

The application will prompt you to choose how to provide the SharePoint site URLs:
1. **Type the target URLs**: Enter URLs manually, one per line, and press Enter on an empty line to finish.
2. **Select target file**: Provide the path to a text file. The app will extract all URLs (starting with http:// or https://) from the file content and scan them.

The application will then scan each URL and output the SharePoint version based on the LibraryVersion. It provides verbose output showing request details, response status, headers, and body for each scan. The POST request includes necessary headers like X-RequestDigest, User-Agent, and cookies for SharePoint authentication and compatibility.

## Requirements

- .NET 8.0 or later
- Internet access to the SharePoint sites

## Notes

- This tool assumes the sites are accessible without authentication. For authenticated sites, additional code for authentication (e.g., using SharePointOnlineCredentials) may be needed.
- The version detection is based on the LibraryVersion from the ProcessQuery response:
  - 14.x: SharePoint 2010
  - 15.x: SharePoint 2013
  - 16.0: SharePoint 2016/2019/Online/SE
  - 16.x: SharePoint Online or 2019
  - 17.x: SharePoint Server Subscription Edition