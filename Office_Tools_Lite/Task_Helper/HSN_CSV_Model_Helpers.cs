using System.Text;
using System.Text.Json.Serialization;
using static Office_Tools_Lite.Task_Helper.HSN_CSV_Model_Helpers;

namespace Office_Tools_Lite.Task_Helper;

public static class HSN_CSV_Model_Helpers
{
    #region Load HSN.json file
    public static async Task<(HashSet<string> HSNCodes,Dictionary<string, string> HSNDescriptions,HashSet<string> UQCCodes)> LoadValidCodesAsync()
    {
        try
        {
            var hsnTempPath = Path.Combine(Path.GetTempPath(), "HSN.json");
            var hsnJsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Task_Helper", "HSN.json");

            string hsnPath = File.Exists(hsnTempPath) ? hsnTempPath : hsnJsonPath;

            if (!File.Exists(hsnPath))
                return (null, null, null);

            var jsonContent = await File.ReadAllTextAsync(hsnPath);

            // Deserialize into the *new* shape
            var data = System.Text.Json.JsonSerializer.Deserialize(jsonContent, HSN_CSV_JsonContext.Default.HSN_CSV_Data);

            if (data == null)
                return (null, null, null);

            // ---- HSN codes -------------------------------------------------
            var hsnSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var descMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in data.HsnCodes)
            {
                if (!string.IsNullOrWhiteSpace(item.Code))
                {
                    hsnSet.Add(item.Code.Trim());
                    descMap[item.Code.Trim()] = item.Description?.Trim() ?? string.Empty;
                }
            }

            // ---- UQC -------------------------------------------------------
            var uqcSet = new HashSet<string>(data.UQCCodes, StringComparer.OrdinalIgnoreCase);

            return (hsnSet, descMap, uqcSet);
        }
        catch
        {
            return (null, null, null);
        }
    }

    public class HSNCodeItem
    {
        [JsonPropertyName("code")]
        public string Code { get; set; } = string.Empty;

        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;
    }

    public class HSN_CSV_Data
    {
        [JsonPropertyName("hsn_codes")]
        public List<HSNCodeItem> HsnCodes { get; set; } = new();

        [JsonPropertyName("UQC")]
        public List<string> UQCCodes { get; set; } = new();
    }
    #endregion

    #region HSN CSV Parser
    public class HSN_CSV_Records
    {
        public string HSN { get; set; }
        public string Description { get; set; }
        public string UQC { get; set; }
        public decimal Total_Quantity { get; set; }
        public decimal Total_Value { get; set; }
        public decimal Taxable_Value { get; set; }
        public decimal Integrated_Tax_Amount { get; set; }
        public decimal Central_Tax_Amount { get; set; }
        public decimal State_UT_Tax_Amount { get; set; }
        public decimal Cess_Amount { get; set; }
        public int Rate { get; set; }
    }

    public static class CsvManualParser
    {
        public static List<HSN_CSV_Records> ParseHSNEntries(string csvFilePath)
        {
            var entries = new List<HSN_CSV_Records>();
            var lines = File.ReadAllLines(csvFilePath);

            if (lines.Length < 2)
                return entries;

            var headers = ParseCsvLine(lines[0]);

            for (var i = 1; i < lines.Length; i++)
            {
                var fields = ParseCsvLine(lines[i]);

                if (i > 10)
                {
                    break; // Limit to first 10 data rows for performance
                }
                var entry = new HSN_CSV_Records
                {
                    HSN = GetField(fields, headers, "HSN"),
                    Description = GetField(fields, headers, "Description"),
                    UQC = GetField(fields, headers, "UQC"),
                    Total_Quantity = ParseDecimal(GetField(fields, headers, "Total Quantity")),
                    Total_Value = ParseDecimal(GetField(fields, headers, "Total Value")),
                    Taxable_Value = ParseDecimal(GetField(fields, headers, "Taxable Value")),
                    Integrated_Tax_Amount = ParseDecimal(GetField(fields, headers, "Integrated Tax Amount")),
                    Central_Tax_Amount = ParseDecimal(GetField(fields, headers, "Central Tax Amount")),
                    State_UT_Tax_Amount = ParseDecimal(GetField(fields, headers, "State/UT Tax Amount")),
                    Cess_Amount = ParseDecimal(GetField(fields, headers, "Cess Amount")),
                    Rate = ParseInt(GetField(fields, headers, "Rate"))
                };

                entries.Add(entry);
            }

            return entries;
        }

        public static string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            var currentField = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField.ToString().Trim());
                    currentField.Clear();
                    continue;
                }

                currentField.Append(c);
            }

            // Add the last field
            fields.Add(currentField.ToString().Trim());

            return fields.ToArray();
        }

        private static string GetField(string[] fields, string[] headers, string columnName)
        {
            for (var i = 0; i < headers.Length; i++)
            {
                if (string.Equals(headers[i], columnName, StringComparison.OrdinalIgnoreCase))
                    return i < fields.Length ? fields[i].Trim() : "";
            }
            return "";
        }

        private static decimal ParseDecimal(string value)
        {
            return decimal.TryParse(value, out var result) ? result : 0m;
        }

        private static int ParseInt(string value)
        {
            return int.TryParse(value, out var result) ? result : 0;
        }
    }
    #endregion

}

[JsonSerializable(typeof(HSN_CSV_Model_Helpers.HSN_CSV_Data))] // enable if error check done by HSN.json
[JsonSerializable(typeof(HSNCodeItem))]
public partial class HSN_CSV_JsonContext : JsonSerializerContext
{
}