using System.Globalization;
using System.Text.Json;

namespace Office_Tools_Lite.Task_Helper;
public static class TimeService
{
    public static async Task<DateTime> GetInternetTimeAsync()
    {
        var istZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");

        using (var client = new HttpClient())
        {
            client.Timeout = TimeSpan.FromSeconds(3);

            // First attempt: GitHub HEAD
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Head, "https://github.com");
                var response = await client.SendAsync(request);

                if (response.Headers.Date.HasValue)
                {
                    var utcTime = response.Headers.Date.Value.UtcDateTime;
                    return TimeZoneInfo.ConvertTimeFromUtc(utcTime, istZone);
                }
            }
            catch { }

            // Second attempt: timeapi.io
            try
            {
                var response = await client.GetStringAsync("https://timeapi.io/api/Time/current/zone?timeZone=Asia/Kolkata");
                var jsonDoc = JsonDocument.Parse(response);
                var istDateTimeString = jsonDoc.RootElement.GetProperty("dateTime").GetString();
                return DateTime.Parse(istDateTimeString, CultureInfo.InvariantCulture);
            }
            catch { }
        }

        // Final fallback
        try
        {
            return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, istZone);
        }
        catch
        {
            return DateTime.Now;
        }
    }
}
