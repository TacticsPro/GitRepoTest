using System.Data;
using System.Xml.Linq;
using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Office_Tools_Lite.Task_Helper;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.XML_Generator;

public class Master_Excel_To_Xml_Converter_6
{
    private DataTable _table;

    #region Execute Main Task
    public async Task Execute(Window mainWindow)
    {
        // File picker dialog
        var filePicker = new FileOpenPicker();
        var hwnd = WindowNative.GetWindowHandle(mainWindow);
        InitializeWithWindow.Initialize(filePicker, hwnd);

        filePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        filePicker.FileTypeFilter.Add(".xlsx");

        var file = await filePicker.PickSingleFileAsync();
        if (file != null)
        {
            try
            {
                string excelPath = file.Path; // Use file.Path instead of dialog.FileName
                _table = await ReadExcel(excelPath);
                if (_table == null) return;

                XDocument xml = BuildLedgerXml(_table);
                string outputPath = Path.Combine(Path.GetDirectoryName(excelPath), "Ledger_Masters_6.xml");
                xml.Save(outputPath);
                await ShowDialog.ShowMsgBox("Success", "XML file generated:\n" + outputPath, "Ok", null, 1, mainWindow);

                // Open the output folder in Explorer
                var outputFolderPath = Path.GetDirectoryName(outputPath); // Get the directory of the output file
                System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
                var dataPackage = new DataPackage();
                dataPackage.SetText(outputPath);
                Clipboard.SetContent(dataPackage);
            }
            catch (Exception ex)
            {
                await ShowDialog.ShowMsgBox("Exception", $"Error: {ex.Message}", "OK", null, 1, mainWindow);
            }
        }
    }
    #endregion

    #region Read Excel Data
    private async Task<DataTable> ReadExcel(string path)
    {
        try
        {
            using (var workbook = new XLWorkbook(path))
            {
                var sheetName = "Ready Data";
                if (!workbook.Worksheets.Contains(sheetName))
                {
                    await ShowDialog.ShowMsgBox("Error", "Worksheet 'Ready Data' not found in the Excel file.", "OK", null, 1, App.MainWindow);
                    return null;
                }

                var worksheet = workbook.Worksheet(sheetName);
                // Disable AutoFilter if it exists
                if (worksheet.AutoFilter.IsEnabled)
                {
                    worksheet.AutoFilter.Clear();
                }

                var range = worksheet.RangeUsed();
                if (range == null || range.RowCount() == 0)
                {
                    await ShowDialog.ShowMsgBox("Error", "No data found in the 'Ready Data' worksheet.", "OK", null, 1, App.MainWindow);
                    return null;
                }

                return range.AsTable().AsNativeDataTable();
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Error reading Excel file: {ex.Message}", "OK", null, 1, App.MainWindow);
            return null;
        }
    }
    #endregion

    #region Build XML
    private XDocument BuildLedgerXml(DataTable table)
    {
        //var ledgers = table.AsEnumerable().Select(row => CreateLedgerElement(row, table));
        int maxRows = 15; // LIMIT to 15 rows
        var ledgers = table.AsEnumerable().Take(maxRows).Select(row => CreateLedgerElement(row, table));

        var xml = new XDocument(
            new XElement("ENVELOPE",
                new XElement("HEADER",
                    new XElement("TALLYREQUEST", "Import Data")
                ),
                new XElement("BODY",
                    new XElement("IMPORTDATA",
                        new XElement("REQUESTDESC",
                            new XElement("REPORTNAME", "All Masters"),
                            new XElement("STATICVARIABLES",
                                new XElement("SVCURRENTCOMPANY", "")
                            )
                        ),
                        new XElement("REQUESTDATA", ledgers)
                    )
                )
            )
        );

        return xml;
    }
    #endregion

    #region Create Voucher Element
    private XElement CreateLedgerElement(DataRow row, DataTable table)
    {
        string name = row["Name"].ToString();
        string parent = row["Parent"].ToString();
        //string ledgerType = table.Columns.Contains("LEDGERTYPE") ? row["LEDGERTYPE"].ToString() : "Regular";
        string gstin = table.Columns.Contains("GSTIN") ? row["GSTIN"].ToString() : "";
        string state = table.Columns.Contains("State") ? row["State"].ToString() : "";
        string country = "India";
        string applicablefrom = "20170401";

        XElement ledger = new XElement("TALLYMESSAGE",
            new XElement("LEDGER",
                new XAttribute("NAME", name),
                new XAttribute("RESERVEDNAME", ""),
                new XElement("NAME", name),
                new XElement("PARENT", parent),
                new XElement("GSTREGISTRATIONTYPE", !string.IsNullOrEmpty(gstin) ? "Regular" : "Unregistered"),
                new XElement("PARTYGSTIN", gstin),
                new XElement("ISGSTAPPLICABLE", !string.IsNullOrEmpty(gstin) ? "Yes" : "No"),
                new XElement("FBTCATEGORY", "Not Applicable"),
                new XElement("ISBILLWISEON", "No"),
                new XElement("ISCOSTCENTRESON", "No"),
                new XElement("ISINTERESTON", "No"),
                new XElement("USEFORGAINLOSS", "No"),
                new XElement("USEFORCOMPUTATION", "No"),
                new XElement("USEFORVAT", "No"),
                new XElement("IGNORETDSEXEMPT", "No"),
                new XElement("ISDEEMEDPOSITIVE", "No"),
                new XElement("AFFECTSSTOCK", "No"),
                new XElement("FORPAYROLL", "No"),
                new XElement("ISABCENABLED", "No"),
                new XElement("ISCOSTTRACKINGON", "No"),
                new XElement("STATENAME", state),

                //Applicable for Tally 6
                new XElement("LEDMAILINGDETAILS.LIST",
                new XElement("APPLICABLEFROM", applicablefrom),
                new XElement("MAILINGNAME", name),
                new XElement("STATE", state),
                new XElement("COUNTRY", country)),

                new XElement("LEDGSTREGDETAILS.LIST",
                new XElement("APPLICABLEFROM", applicablefrom),
                new XElement("GSTREGISTRATIONTYPE", !string.IsNullOrEmpty(gstin) ? "Regular" : "Unregistered"),
                new XElement("GSTIN", gstin))
        )
        );

        return ledger;
    }
    #endregion

}