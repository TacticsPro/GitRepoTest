using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.UI.Xaml.Controls;

namespace Office_Tools_Lite.Task_Helper;

public static class GSTR1_json_generator
{
    // ---------- Decimal converter that writes numbers with two decimal places ----------
    public class DecimalTwoPlacesConverter : JsonConverter<decimal>
    {
        public override decimal Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) =>
            reader.GetDecimal();

        public override void Write(Utf8JsonWriter writer, decimal value, JsonSerializerOptions options)
        {
            // Write numeric value with two decimals (e.g., 118.00)
            var str = value.ToString("F2", CultureInfo.InvariantCulture);
            writer.WriteRawValue(str);
        }
    }

    // ---------- GSTIN validation ----------
    static bool IsValidGSTIN(string gst) =>
    !string.IsNullOrWhiteSpace(gst) && gst.Length == 15;  // No format check

    // ---------- Main conversion function ----------
    public static async Task ConvertExcelToGstr1Json(string inputXlsxPath, string outputJsonPath, string myGstin, string fp)
    {
        if (!File.Exists(inputXlsxPath)) throw new FileNotFoundException("Input not found", inputXlsxPath);
        if (string.IsNullOrWhiteSpace(myGstin)) throw new ArgumentException("Provide your GSTIN (myGstin).");

        var myStateCode = myGstin.Length >= 2 ? myGstin.Substring(0, 2) : "00";

        using var wb = new XLWorkbook(inputXlsxPath);
        var sheetName = "Ready Data";
        if (!wb.Worksheets.Contains(sheetName))
        {
            await ShowDialog.ShowMsgBox("Error", "Worksheet 'Ready Data' not found in the Excel file.", "OK", null, 1, App.MainWindow);
            return;
        }

        var ws = wb.Worksheet(sheetName);
        var headerRow = ws.Row(1);
        int lastCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;

        var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= lastCol; c++)
        {
            var val = headerRow.Cell(c).GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(val) && !headerMap.ContainsKey(val))
                headerMap[val] = c;
            
        }

        static string GetCellString(IXLRow row, Dictionary<string, int> map, string header)
        {
            if (!map.TryGetValue(header, out var col)) return "";
            return row.Cell(col).GetValue<string>()?.Trim() ?? "";
        }

        static decimal GetCellDecimal(IXLRow row, Dictionary<string, int> map, string header)
        {
            if (!map.TryGetValue(header, out var col)) return 0m;
            var cell = row.Cell(col);
            // If numeric cell type, read directly
            if (cell.DataType == XLDataType.Number)
            {
                try { return Convert.ToDecimal(cell.GetDouble()); } catch { }
            }
            var s = cell.GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(s)) return 0m;
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return d;
            return 0m;
        }

        var gstr1 = new Gstr1Root { gstin = myGstin, fp = fp, gt = 0m, cur_gt = 0m };
        // NIL / Non-GST / Exempt accumulators
        decimal totalNonGst = 0m;
        decimal totalNilRated = 0m;
        decimal totalExempt = 0m;

        var b2bMap = new Dictionary<string, Dictionary<string, B2BInvoice>>(StringComparer.OrdinalIgnoreCase);
        var b2csMap = new Dictionary<string, B2CS>(StringComparer.OrdinalIgnoreCase);

        // Track invalid GSTINs for warnings (non-empty attempts only)
        var invalidGstins = new List<(int rowNum, string gstValue, int length)>();

        var invoiceNumbersSeen = new List<string>();
        var invoiceNumbersNumeric = new List<int>();

        var used = ws.RangeUsed();
        int lastRow = used?.LastRowUsed()?.RowNumber() ?? ws.LastRowUsed().RowNumber();

        for (int r = 2; r <= 15; r++) // Upto 15 rows
        {
            var row = ws.Row(r);

            var billNo = GetCellString(row, headerMap, "Bill No.");
            var dateStr = GetCellString(row, headerMap, "Date");
            var custGstin = GetCellString(row, headerMap, "GSTIN");
            var invoiceValue = GetCellDecimal(row, headerMap, "Total Invoice Value");
            var totalTaxable = GetCellDecimal(row, headerMap, "Total Taxable");

            if (string.IsNullOrWhiteSpace(billNo) && string.IsNullOrWhiteSpace(custGstin) && totalTaxable == 0m && invoiceValue == 0m)
                continue;

            if (!string.IsNullOrWhiteSpace(billNo))
            {
                invoiceNumbersSeen.Add(billNo);
                if (int.TryParse(billNo, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
                    invoiceNumbersNumeric.Add(n);
            }

            string idt = "";
            if (DateTime.TryParse(dateStr, out var dt))
                idt = dt.ToString("dd-MM-yyyy");
            else
            {
                try { idt = row.Cell(headerMap.ContainsKey("Date") ? headerMap["Date"] : 0).GetDateTime().ToString("dd-MM-yyyy"); }
                catch { idt = dateStr; }
            }

            // Determine B2B / B2CS using rule: valid GSTIN only => B2B
            bool isB2B = IsValidGSTIN(custGstin);
            bool isB2CS = !isB2B;

            // Track invalid non-empty GSTINs (skip noise for very short/empty)
            if (!isB2B && !string.IsNullOrWhiteSpace(custGstin) && custGstin.Length >= 3)
            {
                invalidGstins.Add((r, custGstin, custGstin.Length));
            }

            string custState = myStateCode;
            bool isInterState = false;
            if (isB2B)
            {
                custState = custGstin.Substring(0, 2);
                isInterState = !string.Equals(custState, myStateCode, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                // B2CS: POS = my state by default
                custState = myStateCode;
                isInterState = false;
            }

            // Read slab columns
            var tx0 = GetCellDecimal(row, headerMap, "Taxable 0%");
            var tx5 = GetCellDecimal(row, headerMap, "Taxable 5%");
            var cgst5 = GetCellDecimal(row, headerMap, "CGST 2.5%");
            var sgst5 = GetCellDecimal(row, headerMap, "SGST 2.5%");

            var tx12 = GetCellDecimal(row, headerMap, "Taxable 12%");
            var cgst6 = GetCellDecimal(row, headerMap, "CGST 6%");
            var sgst6 = GetCellDecimal(row, headerMap, "SGST 6%");

            var tx18 = GetCellDecimal(row, headerMap, "Taxable 18%");
            var cgst9 = GetCellDecimal(row, headerMap, "CGST 9%");
            var sgst9 = GetCellDecimal(row, headerMap, "SGST 9%");

            var tx28 = GetCellDecimal(row, headerMap, "Taxable 28%");
            var cgst14 = GetCellDecimal(row, headerMap, "CGST 14%");
            var sgst14 = GetCellDecimal(row, headerMap, "SGST 14%");

            var itx0 = GetCellDecimal(row, headerMap, "Inter State 0%");
            var itx5 = GetCellDecimal(row, headerMap, "Inter State 5%");
            var i5 = GetCellDecimal(row, headerMap, "IGST 5%");

            var itx12 = GetCellDecimal(row, headerMap, "Inter State 12%");
            var i12 = GetCellDecimal(row, headerMap, "IGST 12%");

            var itx18 = GetCellDecimal(row, headerMap, "Inter State 18%");
            var i18 = GetCellDecimal(row, headerMap, "IGST 18%");

            var itx28 = GetCellDecimal(row, headerMap, "Inter State 28%");
            var i28 = GetCellDecimal(row, headerMap, "IGST 28%");

            var cess = GetCellDecimal(row, headerMap, "CESS");

            // Read NON-GST column from Excel
            var nonGst = GetCellDecimal(row, headerMap, "Non GST");

            if (nonGst > 0m)
            {
                totalNonGst += nonGst;
                continue;       // <-- SKIPS B2B/B2CS processing
            }


            // Helpers
            void AddB2BItem(string ctinKey, string invoiceKey, B2BItem item)
            {
                if (!b2bMap.TryGetValue(ctinKey, out var invMap))
                {
                    invMap = new Dictionary<string, B2BInvoice>(StringComparer.OrdinalIgnoreCase);
                    b2bMap[ctinKey] = invMap;
                }
                if (!invMap.TryGetValue(invoiceKey, out var inv))
                {
                    inv = new B2BInvoice { inum = invoiceKey, idt = idt, val = invoiceValue, pos = custState, itms = new List<B2BItemContainer>() };
                    invMap[invoiceKey] = inv;
                }
                inv.itms.Add(new B2BItemContainer { num = inv.itms.Count + 1, itm_det = item });
            }

            void AddB2CsAggregate(decimal rate, string sply_ty, string pos, decimal txval, decimal camtVal, decimal samtVal, decimal iamtVal, decimal csamtVal)
            {
                var key = $"{rate}|{sply_ty}|{pos}";
                if (!b2csMap.TryGetValue(key, out var ag))
                {
                    ag = new B2CS { rt = rate, sply_ty = sply_ty, pos = pos, typ = "OE", txval = 0m, camt = 0m, samt = 0m, iamt = 0m, csamt = 0m };
                    b2csMap[key] = ag;
                }
                ag.txval += txval;
                ag.camt += camtVal;
                ag.samt += samtVal;
                ag.iamt += iamtVal;
                ag.csamt += csamtVal;
            }

            // Process slabs - when isB2B => create B2B invoice items; else aggregate to b2cs
            // RATE 0
            if ((tx0 > 0m) || (itx0 > 0m))
            {
                if (isB2B)
                {
                    if (isInterState)
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = itx0, rt = 0m, camt = 0m, samt = 0m, iamt = 0m, csamt = 0m });
                    else
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = tx0, rt = 0m, camt = 0m, samt = 0m, iamt = 0m, csamt = 0m });
                }
                else
                {
                    if (itx0 > 0m) AddB2CsAggregate(0m, "INTER", myStateCode, itx0, 0m, 0m, 0m, cess);
                    else AddB2CsAggregate(0m, "INTRA", myStateCode, tx0, 0m, 0m, 0m, cess);
                }
            }

            // RATE 5
            if ((tx5 > 0m) || (itx5 > 0m))
            {
                if (isB2B)
                {
                    if (isInterState)
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = itx5, rt = 5m, camt = 0m, samt = 0m, iamt = i5, csamt = 0m });
                    else
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = tx5, rt = 5m, camt = cgst5, samt = sgst5, iamt = 0m, csamt = 0m });
                }
                else
                {
                    if (itx5 > 0m) AddB2CsAggregate(5m, "INTER", myStateCode, itx5, 0m, 0m, i5, 0m);
                    else AddB2CsAggregate(5m, "INTRA", myStateCode, tx5, cgst5, sgst5, 0m, 0m);
                }
            }

            // RATE 12
            if ((tx12 > 0m) || (itx12 > 0m))
            {
                if (isB2B)
                {
                    if (isInterState)
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = itx12, rt = 12m, camt = 0m, samt = 0m, iamt = i12, csamt = 0m });
                    else
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = tx12, rt = 12m, camt = cgst6, samt = sgst6, iamt = 0m, csamt = 0m });
                }
                else
                {
                    if (itx12 > 0m) AddB2CsAggregate(12m, "INTER", myStateCode, itx12, 0m, 0m, i12, 0m);
                    else AddB2CsAggregate(12m, "INTRA", myStateCode, tx12, cgst6, sgst6, 0m, 0m);
                }
            }

            // RATE 18
            if ((tx18 > 0m) || (itx18 > 0m))
            {
                if (isB2B)
                {
                    if (isInterState)
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = itx18, rt = 18m, camt = 0m, samt = 0m, iamt = i18, csamt = 0m });
                    else
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = tx18, rt = 18m, camt = cgst9, samt = sgst9, iamt = 0m, csamt = 0m });
                }
                else
                {
                    if (itx18 > 0m) AddB2CsAggregate(18m, "INTER", myStateCode, itx18, 0m, 0m, i18, 0m);
                    else AddB2CsAggregate(18m, "INTRA", myStateCode, tx18, cgst9, sgst9, 0m, 0m);
                }
            }

            // RATE 28
            if ((tx28 > 0m) || (itx28 > 0m))
            {
                if (isB2B)
                {
                    if (isInterState)
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = itx28, rt = 28m, camt = 0m, samt = 0m, iamt = i28, csamt = 0m });
                    else
                        AddB2BItem(custGstin, billNo, new B2BItem { txval = tx28, rt = 28m, camt = cgst14, samt = sgst14, iamt = 0m, csamt = 0m });
                }
                else
                {
                    if (itx28 > 0m) AddB2CsAggregate(28m, "INTER", myStateCode, itx28, 0m, 0m, i28, 0m);
                    else AddB2CsAggregate(28m, "INTRA", myStateCode, tx28, cgst14, sgst14, 0m, 0m);
                }
            }

            // Add cess for B2CS if present (naive: aggregated under 0% group if not distributed)
            if (cess > 0m && isB2CS)
            {
                AddB2CsAggregate(0m, "INTRA", myStateCode, 0m, 0m, 0m, 0m, cess);
            }

            // accumulate grand totals
            gstr1.gt += totalTaxable;
            gstr1.cur_gt += invoiceValue;
        }

        // Build final lists (B2B & B2CS)
        foreach (var kv in b2bMap)
        {
            var b2bEntry = new B2B { ctin = kv.Key, inv = kv.Value.Values.ToList() };
            // Remove any entries that accidentally have ctin empty/null - defensive
            if (!string.IsNullOrWhiteSpace(b2bEntry.ctin)) gstr1.b2b.Add(b2bEntry);
        }

        gstr1.b2cs = b2csMap.Values.ToList();

        // Build doc_issue (invoice range)
        var docIssue = new DocIssue();
        var invDoc = new DocDet { doc_num = 1, doc_typ = "Invoices for outward supply", docs = new List<Doc>() };
        if (invoiceNumbersNumeric.Count > 0)
        {
            var min = invoiceNumbersNumeric.Min();
            var max = invoiceNumbersNumeric.Max();
            var tot = invoiceNumbersNumeric.Distinct().Count();
            invDoc.docs.Add(new Doc { num = 1, from = min.ToString(CultureInfo.InvariantCulture), to = max.ToString(CultureInfo.InvariantCulture), totnum = tot, cancel = 0, net_issue = tot });
        }
        else if (invoiceNumbersSeen.Count > 0)
        {
            var from = invoiceNumbersSeen.First();
            var to = invoiceNumbersSeen.Last();
            var tot = invoiceNumbersSeen.Distinct().Count();
            invDoc.docs.Add(new Doc { num = 1, from = from, to = to, totnum = tot, cancel = 0, net_issue = tot });
        }
        else
        {
            invDoc.docs.Add(new Doc { num = 1, from = "0", to = "0", totnum = 0, cancel = 0, net_issue = 0 });
        }
        docIssue.doc_det.Add(invDoc);
        gstr1.doc_issue = docIssue;

        var outputDir = Path.GetDirectoryName(outputJsonPath);

        // Emit warnings for invalid GSTINs (hard stop: no JSON if issues)
        if (invalidGstins.Any())
        {
            var invalidCount = invalidGstins.Count;
            var details = string.Join("\n", invalidGstins.Select(iv => $"Row {iv.rowNum}: '{iv.gstValue}' (length: {iv.length})"));

            var warningMsg = $"Warning: invalid GSTIN(s) detected.\n\n" +
                                "GSTIN must be exactly 15 characters in format (e.g., '29ABCDE1234F1Z5').\n\n" +
                                $"Details:\n{details}\n\n" +
                                "Review and correct your Excel before uploading to GST portal.";


            try
            {
                var logPath = Path.Combine(outputDir, "gstr1_warnings.log");
                File.WriteAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}: {warningMsg}\n");
    
                await ShowDialog.ShowMsgBox("GSTIN Validation Warning", warningMsg, "OK", null, 1, App.MainWindow);
                return; // Do not proceed to write JSON
            }
            catch (Exception Ex)
            {
                Console.WriteLine($"Log File Write Failed: {Ex.Message}");
            }
            
        }

        // After the loop, replace the nil assembly with this:
        NilSupply? nil = null;  // Nullable declaration; starts as null (omitted in JSON via DefaultIgnoreCondition)

        decimal totalNilExemptNonGst = totalNilRated + totalExempt + totalNonGst;
        if (totalNilExemptNonGst > 0m)
        {
            nil = new NilSupply();  // Create only if needed

            // Add INTRAB2C entry (default; adjust if you track INTER separately)
            nil.inv.Add(new NilSupplyInvoice
            {
                sply_ty = "INTRAB2C",  // Or dynamically: "INTERB2C" if interstate non-GST >0
                expt_amt = totalExempt,
                nil_amt = totalNilRated,
                ngsup_amt = totalNonGst
            });

        }

        gstr1.nil = nil;  // null if no data → omitted in JSON thanks to JsonIgnoreCondition

        // Serialize with decimal converter (two decimals)
        var options = new JsonSerializerOptions { WriteIndented = true };
        options.Converters.Add(new DecimalTwoPlacesConverter());

        var json = JsonSerializer.Serialize( gstr1,Gstr1JsonContext.Default.Gstr1Root);
        File.WriteAllText(outputJsonPath, json);
        await ShowDialog.ShowMsgBox("Success", $"Successfully saved GSTR-1 JSON file:\n{outputJsonPath}", "OK", null, 1, App.MainWindow);

        Process.Start("explorer.exe", outputDir);

    }

    // ---------- Models ----------
    public class Gstr1Root
    {
        public string gstin { get; set; }
        public string fp { get; set; }
        public decimal gt { get; set; }
        public decimal cur_gt { get; set; }
        public List<B2B> b2b { get; set; } = new();
        public List<B2CS> b2cs { get; set; } = new();
        public NilSupply nil { get; set; }
        public DocIssue doc_issue { get; set; }
    }

    public class B2B { public string ctin { get; set; }
    public List<B2BInvoice> inv { get; set; } = new(); }
    public class B2BInvoice
    {
        public string inum { get; set; }
        public string idt { get; set; } // dd-MM-yyyy
        public decimal val { get; set; }
        public string pos { get; set; }
        public string rchrg { get; set; } = "N";
        public List<B2BItemContainer> itms { get; set; } = new();
        public string inv_typ { get; set; } = "R";
    }
    public class B2BItemContainer
    { 
        public int num { get; set; } = 1;
        public B2BItem itm_det { get; set; }
    }
    public class B2BItem
    {
        public decimal txval { get; set; }
        public decimal rt { get; set; }
        public decimal camt { get; set; }    // CGST
        public decimal samt { get; set; }    // SGST
        public decimal iamt { get; set; }    // IGST
        public decimal csamt { get; set; }
    }

    public class B2CS
    {
        public decimal rt { get; set; }
        public string sply_ty { get; set; } // INTRA / INTER
        public string pos { get; set; }
        public string typ { get; set; } = "OE";
        public decimal txval { get; set; }
        public decimal camt { get; set; }
        public decimal samt { get; set; }
        public decimal iamt { get; set; }
        public decimal csamt { get; set; }
    }

    public class NilSupply
    {
        public List<NilSupplyInvoice> inv { get; set; } = new();
    }

    public class NilSupplyInvoice
    {
        public string sply_ty { get; set; }
        public decimal expt_amt { get; set; }
        public decimal nil_amt { get; set; }
        public decimal ngsup_amt { get; set; }
    }

    public class DocIssue { public List<DocDet> doc_det { get; set; } = new(); }
    public class DocDet { public int doc_num { get; set; }
    public string doc_typ { get; set; }
    public List<Doc> docs { get; set; } = new(); }
    public class Doc
    { 
        public int num { get; set; }
        public string from { get; set; }
        public string to { get; set; }
        public int totnum { get; set; }
        public int cancel { get; set; }
        public int net_issue { get; set; } 
    }
}

[JsonSourceGenerationOptions(WriteIndented = true,DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    Converters = new[] { typeof(GSTR1_json_generator.DecimalTwoPlacesConverter) }
)]

[JsonSerializable(typeof(GSTR1_json_generator.Gstr1Root))]
[JsonSerializable(typeof(GSTR1_json_generator.B2B))]
[JsonSerializable(typeof(GSTR1_json_generator.B2BInvoice))]
[JsonSerializable(typeof(GSTR1_json_generator.B2BItemContainer))]
[JsonSerializable(typeof(GSTR1_json_generator.B2BItem))]
[JsonSerializable(typeof(GSTR1_json_generator.B2CS))]
[JsonSerializable(typeof(GSTR1_json_generator.NilSupply))]
[JsonSerializable(typeof(GSTR1_json_generator.NilSupplyInvoice))]
[JsonSerializable(typeof(GSTR1_json_generator.DocIssue))]
[JsonSerializable(typeof(GSTR1_json_generator.DocDet))]
[JsonSerializable(typeof(GSTR1_json_generator.Doc))]
public partial class Gstr1JsonContext : JsonSerializerContext
{
}