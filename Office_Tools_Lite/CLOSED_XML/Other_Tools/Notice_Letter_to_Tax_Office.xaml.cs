using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Office_Tools_Lite.CLOSED_XML.Other_Tools;

public sealed partial class Notice_Letter_to_Tax_Office : Page
{
    public Notice_Letter_to_Tax_Office()
    {
        this.InitializeComponent();
    }

    private async void GenerateLetterButton_Click(object sender, Microsoft.UI.Xaml.RoutedEventArgs e)
    {

        // Collect inputs
        string firmName = FirmNameTextBox.Text;
        string propName = PropNameTextBox.Text;
        string address = AddressTextBox.Text;
        string gstin = GSTINTextBox.Text;
        string year = YearTextBox.Text;
        string place = PlaceTextBox.Text;
        string gstr3b = GSTR3BTextBox.Text;
        string gstr2a2b = GSTR2A2BTextBox.Text;
        string noticeITC = NoticeITCTextBox.Text;

        var requiredFields = new (TextBox, string)[]
        {
            (FirmNameTextBox, "Firm Name"),
            (PropNameTextBox, "Prop Name"),
            (AddressTextBox, "Address Field"),
            (GSTINTextBox, "GSTIN"),
            (YearTextBox, "Financial Year"),
            (PlaceTextBox, "Place"),
            (GSTR3BTextBox, "GSTR-3B ITC"),
            (GSTR2A2BTextBox, "GSTR-2A/2B ITC"),
            (NoticeITCTextBox, "Notice ITC")
        };

        foreach (var (field, name) in requiredFields)
        {
            if (string.IsNullOrWhiteSpace(field.Text))
            {
                await ShowDialog.ShowMsgBox("Warning", $" '{name}' cannot be empty.", "OK", null, 1, App.MainWindow);
                ProcessingText.Visibility = Visibility.Collapsed;
                return;
            }
        }
        ProcessingText.Visibility = Visibility.Visible;

        // Show folder picker
        var folderPicker = new FolderPicker();
        folderPicker.SuggestedStartLocation = PickerLocationId.Desktop;
        folderPicker.FileTypeFilter.Add("*");

        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(folderPicker, hwnd);

        var folder = await folderPicker.PickSingleFolderAsync();
        if (folder == null)
        {
            ProcessingText.Visibility = Visibility.Collapsed;
            return;
        }


        // Generate the Word document
        string fileName = $"{firmName}_Letter.docx";
        string filePath = Path.Combine(folder.Path, fileName);

        try
        {
            GenerateWordDocument(filePath, firmName, propName, address, gstin, year, place, gstr3b, gstr2a2b, noticeITC);

            await ShowDialog.ShowMsgBox("Success", $"Letter generated successfully! Saved to {filePath}", "OK", null, 1, App.MainWindow);
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Failed to generate letter: {ex.Message}", "OK", null, 1, App.MainWindow);
            ProcessingText.Visibility = Visibility.Collapsed;
        }
        ProcessingText.Visibility = Visibility.Collapsed;
        var outputFolderPath = Path.GetDirectoryName(filePath); // Get the directory of the output file
        System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
    }

    private void GenerateWordDocument(string filePath, string firmName, string propName, string address, string gstin, string year, string place, string gstr3b, string gstr2a2b, string noticeITC)
    {
        // Create a new document
        using (var doc = DocX.Create(filePath))
        {
            // From Section
            var fromParagraph = doc.InsertParagraph("From,")
                .AppendLine($"\t{firmName}");

            if (!string.IsNullOrEmpty(propName))
            {
                fromParagraph.AppendLine($"\tProp: {propName}");
            }

            // Ensure the address has a leading tab on all lines  
            var addressLines = address.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in addressLines)
            {
                fromParagraph.AppendLine($"\t{line}");
            }

            fromParagraph.AppendLine($"\tGSTIN: {gstin}");

            // To Section
            var toParagraph = doc.InsertParagraph("\nTo,");
            var toLines = new string[] { "Commercial Tax Officer", "<SGSTTO>", "<Place>" };
            foreach (var line in toLines)
            {
                toParagraph.AppendLine($"\t{line}");
            }

            // Subject
            var subjectParagraph = doc.InsertParagraph("\nSubject: Response to Notice for Excess Input Tax Credit Claimed")
                .Bold()
                .Alignment = Alignment.left;

            // Letter Body
            var bodyText = $@"
    Dear Sir/Madam,

            I am writing in response to the notice regarding the excess Input Tax Credit (ITC) claimed due to a mismatch in GSTR-2B/2A. For the year {year}, I claimed an ITC amount of ₹{gstr3b} in GSTR-3B as per our tax invoice. However, as per the notice, the available ITC declared amount is ₹{noticeITC}. Upon reviewing my GSTR-2B/2A, I observed that the actual ITC available is ₹{gstr2a2b}. Hence, we have not claimed excess ITC. This discrepancy appears to be due to an error in the initial claim.
    As per your reference, I have attached the downloaded GSTR-2B/2A statement and forwarded it to your Email-ID.

            I kindly request you to consider the actual ITC available as per GSTR-2B/2A and make the necessary corrections. I assure you that there was no intention to claim excess ITC, and the discrepancy was purely unintentional.

    Thank you for your understanding and cooperation.

    Date: {DateTime.Now:dd-MM-yyyy}								Yours Faithfully
    Place: {place}
    ";
            var bodyParagraph = doc.InsertParagraph(bodyText)
            .Alignment = Alignment.left;


            // Save the document
            doc.Save();
        }
    }

}
