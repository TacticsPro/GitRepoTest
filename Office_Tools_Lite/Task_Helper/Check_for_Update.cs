using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.UI.Xaml;

namespace Office_Tools_Lite.Task_Helper;

public class Check_for_Update
{
    private static readonly HttpClient client = new HttpClient();

    public Version CurrentVersion { get; }

    public Check_for_Update()
    {
        var assembly = Assembly.GetExecutingAssembly();
        CurrentVersion = assembly.GetName().Version;
    }

    public async Task<(bool isUpdateAvailable, string latestVersion, string downloadUrl, string releaseNotes)> IsUpdateAvailableAsync(Window mainWindow)
    {
        try
        {
            //var uri = new Uri("https://raw.githubusercontent.com/TacticsPro/Office_Tools_Latest_Version/refs/heads/main/Release_info.json");
            var infoPath = FinderService.GetCachedInfoPath();
            var uri = new Uri(infoPath);

            var response = await client.GetStringAsync(uri);

            //Use source-generated serializer
           var versionInfo = JsonSerializer.Deserialize(response, versionContext.Default.VersionInfos);


            if (versionInfo == null)
                return (false, null, null, null);

            string latestVersion_lite = versionInfo.lite_latest_version;
            string downloadUrl_lite = versionInfo.lite_msix_download_url;
            string releaseNotes_lite = versionInfo.lite_release_notes;

            if (Version.TryParse(latestVersion_lite, out var latest) && CurrentVersion.CompareTo(latest) < 0)
            {
                return (true, latestVersion_lite, downloadUrl_lite, releaseNotes_lite);
            }

            return (false, latestVersion_lite, downloadUrl_lite, releaseNotes_lite);
        }
        catch
        {
            return (false, null, null, null);
        }
    }

    public async Task<string> getfullversion()
    {
        try
        {
            //var uri = new Uri("https://raw.githubusercontent.com/TacticsPro/Office_Tools_Latest_Version/refs/heads/main/Release_info.json");
            var infoPath = FinderService.GetCachedInfoPath();
            var uri = new Uri(infoPath);

            var response = await client.GetStringAsync(uri);

            //Use source-generated serializer
            var versionInfo = JsonSerializer.Deserialize(response, versionContext.Default.VersionInfos);


            if (versionInfo == null)
                return null;

   
            string downloadUrl_msix = versionInfo.msix_download_url;

            if (Version.TryParse(downloadUrl_msix, out var latest) && CurrentVersion.CompareTo(latest) < 0)
            {
                return downloadUrl_msix;
            }

            return downloadUrl_msix;
        }
        catch
        {
            return null;
        }
    }
    public class VersionInfos
    {
        public string lite_latest_version { get; set; }
        public string lite_msix_download_url { get; set;}
        public string lite_release_notes { get; set; }
        public string msix_download_url { get; set;}
    }
}
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(Check_for_Update.VersionInfos))]
internal partial class versionContext : JsonSerializerContext
{
}