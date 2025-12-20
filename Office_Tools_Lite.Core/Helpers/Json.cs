using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace Office_Tools_Lite.Core.Helpers;

public static class Json
{
    public static async Task<T?> ToObjectAsync<T>(string value)
    {
        return await Task.Run(() =>
        {
            var typeInfo = (JsonTypeInfo<T>)AppJsonContext.Default.GetTypeInfo(typeof(T))!;
            return JsonSerializer.Deserialize(value, typeInfo);
        });
    }

    public static async Task<string> StringifyAsync(object value)
    {
        return await Task.Run(() =>
        {
            var type = value.GetType();
            var typeInfo = AppJsonContext.Default.GetTypeInfo(type)
                           ?? throw new NotSupportedException($"Type not registered in AppJsonContext: {type}");

            return JsonSerializer.Serialize(value, typeInfo);
        });
    }
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(Dictionary<string, object>))]
public partial class AppJsonContext : JsonSerializerContext
{
}
