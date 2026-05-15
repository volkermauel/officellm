using System.Text;

namespace OfficeMcpServer.Models;

/// <summary>
/// Serves a templated manifest.xml with dynamic URLs inferred from the request Host header.
/// Placeholders: {{SCHEME}}, {{HOST}}, {{BASE_URL}}
/// </summary>
public static class ManifestTemplate
{
    private const string Template = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<OfficeApp
  xmlns=""http://schemas.microsoft.com/office/appforoffice/1.1""
  xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
  xmlns:bt=""http://schemas.microsoft.com/office/officeappbasictypes/1.0""
  xmlns:ov=""http://schemas.microsoft.com/office/taskpaneappversionoverrides/ov""
  xsi:type=""TaskPaneApp"">

  <!-- Basic Info -->
  <Id>a1b2c3d4-e5f6-7890-abcd-ef1234567890</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Office LLM Harness</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue=""Office LLM Harness""/>
  <Description DefaultValue=""Connects PowerPoint to Open WebUI via a local MCP server for LLM-assisted workflows.""/>
  <IconUrl DefaultValue=""{{BASE_URL}}/assets/icon-32.png""/>
  <HighResolutionIconUrl DefaultValue=""{{BASE_URL}}/assets/icon-64.png""/>

  <!-- Support -->
  <SupportUrl DefaultValue=""{{BASE_URL}}/""/>

  <!-- Domains (needed for localhost dev) -->
  <AppDomains>
    <Domain>localhost</Domain>
    <Domain>127.0.0.1</Domain>
  </AppDomains>

  <!-- Default Settings (overrides in VersionOverrides for production) -->
  <Hosts>
    <Host name=""Presentation""/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue=""{{BASE_URL}}/index.html""/>
    <RequestedWidth>400</RequestedWidth>
  </DefaultSettings>

  <!-- Permissions -->
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>";

    /// <summary>
    /// Renders the manifest template with URLs from the request.
    /// </summary>
    public static string Render(HttpRequest request)
    {
        string scheme = request.Scheme; // "http" or "https"
        string host = request.Host.ToString(); // "localhost:3000" or "officellm.apps.rp.alliance.co.uk"
        string baseUrl = $"{scheme}://{host}";

        return Template
            .Replace("{{SCHEME}}", scheme)
            .Replace("{{HOST}}", host)
            .Replace("{{BASE_URL}}", baseUrl);
    }

    /// <summary>
    /// Returns the content type for the manifest.
    /// </summary>
    public static string ContentType => "application/xml";

    /// <summary>
    /// Returns the encoding for the manifest.
    /// </summary>
    public static Encoding Encoding => Encoding.UTF8;
}
