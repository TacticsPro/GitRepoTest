using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Windows.System;
using static Office_Tools_Lite.Task_Helper.MachineLicenseInfo;

namespace Office_Tools_Lite.Task_Helper;

public class Finder
{
    private static readonly HttpClient client = new HttpClient(); // Reuse HttpClient
    private static byte[] key;
    private static byte[] iv;
    private static readonly DateTime OfflineExpiryDate = new DateTime(0);
    private static readonly string configFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Microsoft", "Runtime_req", "OTEW.enc");
    private static readonly string firstTimeOpenFilePath = Path.Combine(Path.GetTempPath(), "firstTimeOpen.txt");
    private static readonly string hsnTempPath = Path.Combine(Path.GetTempPath(), "HSN.json");

    private DateTime currentDate;
    private readonly JsonDocument secureConfig;
    private int initialofflinecount = 3;
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]

    private static extern bool GetVolumeInformation(
            string lpRootPathName,
            StringBuilder lpVolumeNameBuffer,
            int nVolumeNameSize,
            out uint lpVolumeSerialNumber,
            out uint lpMaximumComponentLength,
            out uint lpFileSystemFlags,
            StringBuilder lpFileSystemNameBuffer,
            int nFileSystemNameSize);


    public Finder()
    {
        // Set GitHub timeout (so app won't freeze)
        client.Timeout = TimeSpan.FromSeconds(10);

        // Load and parse Required_WU.dll (decrypted JSON with secrets)
        byte[] encBytes = Transformation.GetTransformedFileContent("Required_WU.dll");
        if (encBytes == null)
            throw new InvalidOperationException("Required_WU.dll not found or not decrypted.");

        string json = Encoding.UTF8.GetString(encBytes);
        secureConfig = JsonDocument.Parse(json);

        // Create local folder if needed
        Directory.CreateDirectory(Path.GetDirectoryName(configFilePath));

        LoadonLaunch();
        _ = InitializeDate();
        _ = RetryPendingEmails();
        _ = InitializeKeysAsync(); // Initialize keys asynchronously
        //_ = DownloadHSNFile();
    }

    #region OnLaunch
    public string GetReleasePage()
    {
        return secureConfig.RootElement
            .GetProperty("ExternalLinks")
            .GetProperty("ReleasePage")
            .GetString()
            ?? throw new InvalidOperationException("ReleasePage is missing");
    }

    private void LoadonLaunch()
    {
        // Use secret GitHub PAT from encrypted config
        string gstring = secureConfig.RootElement.GetProperty("String").GetString()
                ?? throw new InvalidOperationException("String is missing");

        client.DefaultRequestHeaders.Add("Authorization", "token " + gstring);
        client.DefaultRequestHeaders.Add("User-Agent", "Office_Tools_Lite");
        client.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");
    }

    public void Dispose()
    {
        secureConfig?.Dispose();
    }
    #endregion

    #region Initialize Key
    private static async Task InitializeKeysAsync()
    {
        // Skip if already initialized
        if (key != null && iv != null)
        {
            return;
        }

        try
        {
            var prompt = new Prompt();
            string firstpart = await prompt.PrompterK();
            string secondpart = await prompt.PrompterI();
            key = Convert.FromBase64String(firstpart);
            iv = Convert.FromBase64String(secondpart);
            //await ShowDialog.ShowMsgBox("Success", "InitializeKeysAsync", "OK", null, 1, App.MainWindow);
        }
        catch (FormatException ex)
        {
            //await ShowDialog.ShowMsgBox1("Failed", $"Invalid Base64 string for key or IV: {ex.Message}", App.MainWindow);
            Console.WriteLine($"Invalid Base64 string for key or IV: {ex.Message}");
            throw new InvalidOperationException("Failed to initialize encryption keys.");
        }
    }
    #endregion


    #region Initialize internet Time
    public async Task InitializeDate()
    {
        try
        {
            DateTime istTime = await TimeService.GetInternetTimeAsync();
            currentDate = istTime;
        }
        catch
        {
            currentDate = DateTime.Now;
        }
    }

    public void ForceSetCurrentDate(DateTime dt)
    {
        currentDate = dt;
    }
    #endregion

    #region Get Machine Details
    public static string GetMachineIdentifier()
    {
        string currentMachineId = Get_UUID.GetMachineIdentifier().Replace("-", "");
        return currentMachineId;
    }

    #endregion

    #region Encryption
    private static string EncryptData(string data)
    {
        // Ensure keys are initialized
        if (key == null || iv == null)
        {
            Task.Run(async () => await InitializeKeysAsync()).GetAwaiter().GetResult();
        }

        try
        {
            //File.AppendAllText("debug.log", $"Encrypting data: {data}\n");
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;

                using (MemoryStream ms = new MemoryStream())
                using (CryptoStream cs = new CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write))
                using (StreamWriter sw = new StreamWriter(cs))
                {
                    sw.Write(data);
                    sw.Close();
                    string encrypted = Convert.ToBase64String(ms.ToArray());
                    //File.AppendAllText("debug.log", $"Encryption successful, length: {encrypted.Length}\n");
                    return encrypted;
                }
            }
        }
        catch (CryptographicException ex)
        {
            //File.AppendAllText("error.log", $"Cryptographic Error in EncryptData: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            throw;
        }
        catch (Exception ex)
        {
            //File.AppendAllText("error.log", $"Unexpected Error in EncryptData: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            throw;
        }
    }
    #endregion

    #region Decryption
    private static string DecryptData(string data)
    {
        // Ensure keys are initialized
        if (key == null || iv == null)
        {
            Task.Run(async () => await InitializeKeysAsync()).GetAwaiter().GetResult();
        }

        using (Aes aes = Aes.Create())
        {
            aes.Key = key;
            aes.IV = iv;

            using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(data)))
            using (CryptoStream cs = new CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Read))
            using (StreamReader sr = new StreamReader(cs))
            {
                return sr.ReadToEnd();
            }
        }
    }
    #endregion

    #region Save Details
    private static void SaveDetailsToConfig(string productId, string emailId, string licenceStatus = "Expired", int offlineruncount = 0, DateTime expiryDate = default, string ActivateMode = "Offline", DateTime Activationtime = default)
    {
        try
        {
            var info = new MachineLicenseInfo
            {
                machineId = GetMachineIdentifier(),
                productId = productId,
                emailId = emailId,
                licenceStatus = licenceStatus,
                offlineruncount = offlineruncount,
                expiryDate = expiryDate,
                ActivateMode = ActivateMode,
                Activationtime = Activationtime,
            };

            // Use the source-generated context for serialization
            string json = JsonSerializer.Serialize(info, MachineLicenseInfoJsonContext.Default.MachineLicenseInfo);
            //File.AppendAllText("debug.log", $"Serialized JSON: {json}\n");

            //var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true, WriteIndented = true }; // hidden because Reflection-based serialization has been disabled for this application
            //string json = JsonSerializer.Serialize(info, options);

            if (string.IsNullOrEmpty(json) || json == "{}")
            {
                //File.AppendAllText("error.log", $"Serialization failed: Empty JSON for MachineLicenseInfo: productId={productId}, emailId={emailId}\n");
                throw new InvalidOperationException("Serialization resulted in empty JSON");
            }

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(configFilePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
                //File.AppendAllText("debug.log", $"Created directory: {directory}\n");
            }

            //File.AppendAllText("debug.log", $"Encrypting JSON for configFilePath: {configFilePath}\n");
            string encrypted = EncryptData(json);
            //File.AppendAllText("debug.log", $"Encrypted data length: {encrypted.Length}\n");

            //File.AppendAllText("debug.log", $"Writing to configFilePath: {configFilePath}\n");
            File.WriteAllText(configFilePath, encrypted);
            //File.AppendAllText("debug.log", $"Successfully wrote to configFilePath: {configFilePath}\n");
        }
        catch (JsonException ex)
        {
            //File.AppendAllText("error.log", $"JSON Serialization Error in SaveDetailsToConfig: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            throw;
        }
        catch (CryptographicException ex)
        {
            //File.AppendAllText("error.log", $"Encryption Error in SaveDetailsToConfig: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            throw;
        }
        catch (IOException ex)
        {
            //File.AppendAllText("error.log", $"File I/O Error in SaveDetailsToConfig: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            throw;
        }
        catch (UnauthorizedAccessException ex)
        {
            //File.AppendAllText("error.log", $"Access Denied Error in SaveDetailsToConfig: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            throw;
        }
        catch (Exception ex)
        {
            //File.AppendAllText("error.log", $"Unexpected Error in SaveDetailsToConfig: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            throw;
        }
    }
    #endregion

    #region Load Licence
    private async Task<(DateTime expiryDate, string productId, string licenceStatus)> LoadLicence()
    {
        if (!File.Exists(configFilePath))
        {
            //File.AppendAllText("error.log", "Licence File Does not exist \" +\r\n");
            SaveDetailsToConfig(string.Empty, string.Empty);
            return (OfflineExpiryDate, string.Empty, string.Empty);
        }
        else
        {
            //File.AppendAllText("error.log", "Licence File exist \" +\r\n");

            try
            {
                //File.AppendAllText("error.log", "Entering Try case \" +\r\n");
                string encryptedData = File.ReadAllText(configFilePath);
                string decryptedJson = DecryptData(encryptedData);

                // First, parse into JsonDocument to validate JSON
                using var jsonDocument = JsonDocument.Parse(decryptedJson);
                var root = jsonDocument.RootElement;
                //File.AppendAllText("error.log", $"{root} \" +\r\n");

                try
                {
                    // Manually extract properties
                    var data = new MachineLicenseInfo
                    {
                        machineId = root.GetProperty("machineId").GetString() ?? string.Empty,
                        expiryDate = root.GetProperty("expiryDate").GetDateTime(),
                        productId = root.GetProperty("productId").GetString() ?? string.Empty,
                        emailId = root.GetProperty("emailId").GetString() ?? string.Empty,
                        licenceStatus = root.GetProperty("licenceStatus").GetString() ?? string.Empty,
                        offlineruncount = root.GetProperty("offlineruncount").GetInt32(),
                        ActivateMode = root.GetProperty("ActivateMode").GetString() ?? string.Empty,
                        Activationtime = root.GetProperty("Activationtime").GetDateTime()
                    };
                    return (data.expiryDate, data.productId, data.licenceStatus);

                }
                catch (JsonException ex)
                {
                    //File.AppendAllText("error.log", $"JSON Deserialization Error: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                    throw;
                }
                catch (ArgumentException ex)
                {
                    //File.AppendAllText("error.log", $"Argument Deserialization Error: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                    throw;
                }
                catch (InvalidOperationException ex)
                {
                    //File.AppendAllText("error.log", $"Operation Deserialization Error: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                    throw;
                }
                catch (Exception ex)
                {
                    //File.AppendAllText("error.log", $"Unexpected Deserialization Error: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                    throw;
                }
            }
            catch (Exception ex)
            {
                //File.AppendAllText("error.log", $"Outer exception in LoadLicence: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                SaveDetailsToConfig(string.Empty, string.Empty);
                await ShowDialog.ShowMsgBox("Exceptions", "error occurred while licence check.\nContinue with Search Licence. ", "OK", null, 1, App.MainWindow);
                return (OfflineExpiryDate, string.Empty, string.Empty);
            }
        }
    }
    #endregion

    #region Get Licence Deatils
    public async Task<(string EmailId, string productId, string machineId, DateTime Activationtime)> GetLicenseDetails()
    {
        try
        {
            if (!File.Exists(configFilePath))
            {
                return (string.Empty, string.Empty, string.Empty, default);
            }

            string encryptedData = File.ReadAllText(configFilePath);
            string decryptedJson = DecryptData(encryptedData);

            // Manually parse JSON using JsonDocument
            using var jsonDocument = JsonDocument.Parse(decryptedJson);
            var root = jsonDocument.RootElement;

            // Extract properties with default values if missing
            var data = new MachineLicenseInfo
            {
                machineId = root.TryGetProperty("machineId", out var machineIdProp) ? machineIdProp.GetString() ?? string.Empty : string.Empty,
                productId = root.TryGetProperty("productId", out var productIdProp) ? productIdProp.GetString() ?? string.Empty : string.Empty,
                emailId = root.TryGetProperty("emailId", out var emailIdProp) ? emailIdProp.GetString() ?? string.Empty : string.Empty,
                Activationtime = root.TryGetProperty("Activationtime", out var activationTimeProp) && DateTime.TryParse(activationTimeProp.GetString(), out var activationTime) ? activationTime : default
            };

            return (data.emailId, data.productId, data.machineId, data.Activationtime);
        }
        catch (JsonException ex)
        {
            //File.AppendAllText("error.log", $"JSON Parsing Error in GetLicenseDetails: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            return (string.Empty, string.Empty, string.Empty, default);
        }
        catch (Exception ex)
        {
            //File.AppendAllText("error.log", $"Unexpected Error in GetLicenseDetails: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
            return (string.Empty, string.Empty, string.Empty, default);
        }
    }
    #endregion

    #region Send Product Key Date
    private async Task SendProductKEY_Date_Time(string product, string userEmailId, string expiryDate, string machineID = "", string oldProductId = null)
    {
        const int maxRetries = 3;
        const int retryDelayMs = 2000;
        int attempt = 0;
        bool success = false;
        Exception lastException = null;

        while (attempt < maxRetries && !success)
        {
            attempt++;
            try
            {
                DateTime dateTime = DateTime.Now;
                var fromAddress = new MailAddress("gpenmail@gmail.com", "Office Tools Lite");
                var toAddress = new MailAddress("c.rakshith@ymail.com");
                var toUserAddress = new MailAddress(userEmailId);
                string fstring = secureConfig.RootElement.GetProperty("Strings").GetString()
                         ?? throw new InvalidOperationException("Strings is missing in Required.dll");


                using (var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(fromAddress.Address, fstring),
                    Timeout = 5000
                })
                {
                    var messageToSupport = new MailMessage(fromAddress, toAddress)
                    {
                        Subject = "Office Tools Lite Activation",
                        Body = $"Email ID: {userEmailId}\n" +
                               $"Product ID: {product}\n" +
                               $"Activated Time: {dateTime}\n" +
                               (oldProductId != null ? $"\nOld Product ID: {oldProductId}" : "\n") +
                               $"Machine ID: {machineID}\n"
                    };

                    var messageToUser = new MailMessage(fromAddress, toUserAddress)
                    {
                        Subject = "Office Tools Lite Activation",
                        Body = $"🎉 Congratulations!\n" +
                               "You have successfully activated Office_Tools_Lite." +
                               "Please save the following details for future reference:\n" +
                               $"🆔 Product ID: {product}\n" +
                               $"⏰ Activated On: {dateTime}\n" +
                               (oldProductId != null ? $"Old Product ID: {oldProductId}\n" : "") +
                               $"Machine ID: {machineID}\n" +
                               "If you face any issues while using the Office_Tools_Lite application, feel free to reach out to us via email. We're here to help!\n\n" +
                               "Thank you for choosing Office_Tools_Lite.\n" +
                               "Wishing you a smooth and productive experience ahead!\n\n" +
                               "Warm regards,\n" +
                               "Office_Tools_Lite Support Team"
                    };

                    await smtp.SendMailAsync(messageToSupport);
                    await smtp.SendMailAsync(messageToUser);

                    messageToSupport.Dispose();
                    messageToUser.Dispose();
                }

                await UpdateGitHubFiles(product, userEmailId, dateTime.ToString(), expiryDate, oldProductId);
                success = true;
            }
            catch (SmtpException ex)
            {
                lastException = ex;
                if (attempt < maxRetries)
                    await Task.Delay(retryDelayMs);
            }
            catch (Exception ex)
            {
                lastException = ex;
                break;
            }
        }

        if (!success)
        {
            string pendingDir = Path.GetDirectoryName(configFilePath);

            // 🔁 Delete any previous pending email JSON files
            foreach (var oldFile in Directory.GetFiles(pendingDir, "pending_email_*.json"))
            {
                try
                {
                    File.Delete(oldFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete old pending file: {ex.Message}");
                }
            }

            // Save new pending email JSON file
            File.WriteAllText(
                Path.Combine(pendingDir, $"pending_email_{product}_{DateTime.Now.Ticks}.json"),
                JsonSerializer.Serialize(new
                {
                    product,
                    userEmailId,
                    expiryDate,
                    machineID,
                    oldProductId
                }));

            throw new Exception($"Failed to send confirmation emails: {lastException?.Message}", lastException);
        }
    }

    #endregion

    #region Retry Email Send
    public async Task RetryPendingEmails()
    {
        string pendingEmailDir = Path.GetDirectoryName(configFilePath);
        var pendingFiles = Directory.GetFiles(pendingEmailDir, "pending_email_*.json");

        foreach (var file in pendingFiles)
        {
            try
            {
                string jsonContent = File.ReadAllText(file);
                var emailData = JsonSerializer.Deserialize(jsonContent, PendingEmailDataJsonContext.Default.PendingEmailData)
                    ?? throw new JsonException($"Failed to deserialize pending email: {file}");

                Console.WriteLine($"Retrying pending email: {file} -> {emailData.userEmailId}");

                await SendProductKEY_Date_Time(
                    emailData.product,
                    emailData.userEmailId,
                    emailData.expiryDate,
                    emailData.machineID,
                    emailData.oldProductId);

                // Delete file only after successful resend
                File.Delete(file);
                Console.WriteLine($"Successfully sent and deleted: {file}");
            }
            catch (JsonException ex)
            {
                Console.WriteLine($"Corrupted JSON in pending file: {file} => {ex.Message}");
                try
                {
                    File.Delete(file);
                    Console.WriteLine($"Deleted corrupted file: {file}");
                }
                catch (Exception deleteEx)
                {
                    Console.WriteLine($"Error deleting corrupted file: {deleteEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Retry failed for {file}: {ex.Message}");
            }
        }
    }
    #endregion

    //#region Search for Existing Licence
    //private async Task<MachineLicenseInfo.Product> SearchforExistingLicence(string machineId)
    //{
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "Product_IDs/sold_product_ids.json";
    //    string apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{filePath}";

    //    try
    //    {
    //        var response = await client.GetAsync(apiUrl);
    //        if (response.StatusCode == HttpStatusCode.NotFound)
    //        {
    //            return null; // No sold products file, so no existing licence
    //        }
    //        response.EnsureSuccessStatusCode();
    //        string responseContent = await response.Content.ReadAsStringAsync();
    //        var jsonDocument = JsonDocument.Parse(responseContent);
    //        var root = jsonDocument.RootElement;

    //        string contentBase64 = root.GetProperty("content").GetString();
    //        string jsonContent = Encoding.UTF8.GetString(Convert.FromBase64String(contentBase64));
    //        var productsDocument = JsonDocument.Parse(jsonContent);
    //        var productsRoot = productsDocument.RootElement;

    //        var soldProductIds = productsRoot.GetProperty("sold_product_ids").EnumerateArray();
    //        foreach (var product in soldProductIds)
    //        {
    //            string activatedMachineId = product.TryGetProperty("ActivatedMachineId", out var machineProp) ? machineProp.GetString() : null;
    //            if (activatedMachineId == machineId)
    //            {
    //                // Found matching machine ID, construct and return the Product object
    //                return new MachineLicenseInfo.Product
    //                {
    //                    ProductId = product.GetProperty("ProductId").GetString(),
    //                    EmailId = product.GetProperty("EmailId").GetString(),
    //                    ActivatedMachineId = activatedMachineId,
    //                    ActivationTime = product.GetProperty("ActivationTime").GetString(),
    //                    ExpiryDate = product.GetProperty("ExpiryDate").GetString(),
    //                    LicenceStatus = product.GetProperty("LicenceStatus").GetString()
    //                };
    //            }
    //        }
    //        return null; // No matching machine ID found
    //    }
    //    catch (HttpRequestException ex)
    //    {
    //        Console.WriteLine($"HTTP Error in SearchforExistingLicence: {ex.Message}");
    //        return null;
    //    }
    //    catch (JsonException ex)
    //    {
    //        Console.WriteLine($"JSON Error in SearchforExistingLicence: {ex.Message}");
    //        return null;
    //    }
    //}
    //#endregion

    //#region Search for Licence
    //private async Task SearchforLicence()
    //{
    //    string currentMachineId = GetMachineIdentifier();
    //    var existingProduct = await SearchforExistingLicence(currentMachineId);

    //    if (existingProduct != null)
    //    {
    //        try
    //        {
    //            DateTime expiryDate = DateTime.ParseExact(existingProduct.ExpiryDate, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
    //            DateTime activationTime = DateTime.Parse(existingProduct.ActivationTime); // Assuming ActivationTime is in a parsable format

    //            // Save details to config with "Active" status if the licence exists
    //            SaveDetailsToConfig(existingProduct.ProductId,existingProduct.EmailId, "Active",0,expiryDate,"Online",activationTime);
    //            File.WriteAllText(firstTimeOpenFilePath, "True");

    //            if (currentDate.Date > expiryDate)
    //            {
    //                await ShowDialog.ShowMsgBox("Warining", "Licence has been expired. Please activate manually", "OK", null, 1, App.MainWindow);
    //            }
    //            else
    //            {
    //                await ShowDialog.ShowMsgBox("Acivated", "Licence Exists and Activated successfully. Please Restart the App", "OK", null, 1, App.MainWindow);

    //            }

    //        }
    //        catch (FormatException ex)
    //        {
    //            await ShowDialog.ShowMsgBox("Error", "Error activating licence. Contact support.", "OK", null, 1, App.MainWindow);
    //        }
    //    }
    //    else
    //    {
    //        await ShowDialog.ShowMsgBox("Warning", "No existing licence found for this machine. Please activate manually.", "OK", null, 1, App.MainWindow);
    //    }
    //}
    //#endregion

    #region Get New Expiry Date
    public async Task<DateTime> GetNewExpiryDate(Window mainWindow)
    {
        var (expiryDate, productId, licenceStatus) = await LoadLicence();

        if (expiryDate.Date >= currentDate.Date & !string.IsNullOrEmpty(productId))
        {
            return expiryDate;
        }

        var panel = new StackPanel
        {
            Margin = new Thickness(10),
            VerticalAlignment = VerticalAlignment.Center
        };

        panel.Children.Add(new TextBlock
        {
            Text = "Activation required. Do you want to proceed?",
            Margin = new Thickness(0, 0, 0, 20),
            TextAlignment = TextAlignment.Center
        });

        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center
        };

        var activateBtn = new Button
        {
            Content = "Activate",
            Width = 100,
            Height = 35,
            Margin = new Thickness(5),
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0x83, 0x66, 0xEC)), // #8366EC
            Foreground = new SolidColorBrush(Colors.White),
            CornerRadius = new CornerRadius(8)
        };

        var cancelBtn = new Button
        {
            Content = "Cancel",
            Width = 100,
            Height = 35,
            Margin = new Thickness(5),
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xFF, 0x99, 0x99)), // #FF9999 (light red for cancel)
            CornerRadius = new CornerRadius(8)
        };

        string releasePage = GetReleasePage();
        var hyperlinkButton = new HyperlinkButton
        {
            Content = "Click here to visit web",
            NavigateUri = new Uri($"{releasePage}"),
            Margin = new Thickness(0, 0, 0, 20),
            HorizontalAlignment = HorizontalAlignment.Center
        };

        panel.Children.Add(hyperlinkButton);

        buttonPanel.Children.Add(activateBtn);
        buttonPanel.Children.Add(cancelBtn);
        panel.Children.Add(buttonPanel);

        var helpBtn = new Button
        {
            Content = "Get Help",
            Width = 80,
            Margin = new Thickness(0, 10, 0, 10),
            HorizontalAlignment = HorizontalAlignment.Center,
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xB1, 0xFD, 0xBA)), // #B1FDBA
            Foreground = new SolidColorBrush(Colors.Black),
            CornerRadius = new CornerRadius(8)
        };
        panel.Children.Add(helpBtn);

        var dialog = new ContentDialog
        {
            Title = "Activation Check",
            Content = panel,
            CloseButtonText = "",
            XamlRoot = mainWindow.Content.XamlRoot
        };

        // Flags to decide what to do AFTER the dialog closes
        bool shouldActivate = false;
        bool shouldReactivate = false;
        bool shouldsearchlicence = false;

        activateBtn.Click += (s, e) =>
        {
            shouldActivate = true;
            dialog.Hide();
        };

        cancelBtn.Click += (s, e) =>
        {
            dialog.Hide();
        };

        helpBtn.Click += async (s, e) =>
        {
            var htmlFileUri = App.GetCachedHtmlPath();
            await Launcher.LaunchUriAsync(new Uri(htmlFileUri));
        };

        await dialog.ShowAsync(); // Wait for user choice

        if (shouldActivate)
        {
            return await ActivateLicence(mainWindow);
        }
        //else if (shouldReactivate)
        //{
        //    return await ReActivateLicence(mainWindow);
        //}


        // Cancel or failure
        Environment.Exit(0);
        return DateTime.MinValue;
    }

    #endregion

    #region Activate Licence
    private async Task<DateTime> ActivateLicence(Window mainWindow)
    {
        bool isActivated = false;
        DateTime newExpiryDate = DateTime.Now.Date.AddDays(365);

        var panel = new StackPanel { Margin = new Thickness(10) };

        var userEmailTxt = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
        panel.Children.Add(new TextBlock { Text = "Email ID:", Margin = new Thickness(0, 0, 0, 5) });
        panel.Children.Add(userEmailTxt);

        var productTxt = new TextBox { Margin = new Thickness(0, 0, 0, 10), Text = "12345678-Lite", IsEnabled = false };
        panel.Children.Add(new TextBlock { Text = "Enter Product ID:", Margin = new Thickness(0, 0, 0, 5) });
        panel.Children.Add(productTxt);

        var activationMsg = new TextBlock { Text = "Enter your details to activate.", Margin = new Thickness(0, 10, 0, 0) };

        var activateBtn = new Button
        {
            Content = "Activate",
            Width = 120,
            Height = 30,
            Margin = new Thickness(5),
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0x83, 0x66, 0xEC)), // #8366EC
            Foreground = new SolidColorBrush(Colors.White),
        };

        var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
        buttonPanel.Children.Add(activateBtn);

        panel.Children.Add(buttonPanel);
        panel.Children.Add(activationMsg);

        var dialog = new ContentDialog
        {
            Title = "Activate License",
            Content = panel,
            CloseButtonText = "Cancel",
            XamlRoot = mainWindow.Content.XamlRoot
        };

        activateBtn.Click += async (s, e) =>
        {
            activateBtn.IsEnabled = false;
            activationMsg.Text = "Please wait...";

            string product = productTxt.Text.Trim();
            string email = userEmailTxt.Text.Trim();

            if (string.IsNullOrWhiteSpace(product) || string.IsNullOrWhiteSpace(email))
            {
                activationMsg.Text = "Product ID and Email are required.";
                activateBtn.IsEnabled = true;
                return;
            }

            if (!IsValidEmail(email))
            {
                activationMsg.Text = "Invalid Email format.";
                activateBtn.IsEnabled = true;
                return;
            }

            try
            {

                await SendProductKEY_Date_Time(product, email, newExpiryDate.ToString("dd-MM-yyyy"), GetMachineIdentifier());
                SaveDetailsToConfig(product, email, "Active", 0, newExpiryDate, "Online", DateTime.Now);
                File.WriteAllText(firstTimeOpenFilePath, "True");

                isActivated = true;
                dialog.Hide();
                await Task.Delay(500);
                await ShowDialog.ShowMsgBox("Success", "Product Activated Successfully!", "OK", null, 1, mainWindow);
            }
            catch (Exception ex)
            {
                dialog.Hide();
                await Task.Delay(200);
                await ShowDialog.ShowMsgBox("Error", $"Activation failed: {ex.Message}", "Ok", null, 1, mainWindow);
            }
        };

        await dialog.ShowAsync();

        if (!isActivated)
        {
            Environment.Exit(0);
        }

        return newExpiryDate;
    }

    #endregion

    //#region ReActivate Licence
    //private async Task<DateTime> ReActivateLicence(Window mainWindow)
    //{
    //    bool isActivated = false;
    //    DateTime newExpiryDate = DateTime.MinValue;

    //    var panel = new StackPanel { Margin = new Thickness(10) };

    //    var userEmailTxt = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
    //    var oldProductTxt = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
    //    var newProductTxt = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
    //    var activateBtn = new Button { Content = "Activate", Width = 120, Margin = new Thickness(0, 10, 0, 10),
    //        Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0x83, 0x66, 0xEC)), // #8366EC
    //        Foreground = new SolidColorBrush(Colors.White),
    //    };
    //    var searchforlicenceBtn = new Button { Content = "Search for Licence", Margin = new Thickness(150, -51, 0, 0),
    //        Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xC7, 0x63, 0xEF)), // #C763EF
    //        Foreground = new SolidColorBrush(Colors.White),
    //    };
    //    var activationMsg = new TextBlock { Text = "Enter your details to re-activate.", Margin = new Thickness(0, 10, 0, 0) };

    //    panel.Children.Add(new TextBlock { Text = "Email ID:" });
    //    panel.Children.Add(userEmailTxt);
    //    panel.Children.Add(new TextBlock { Text = "Old Product ID:" });
    //    panel.Children.Add(oldProductTxt);
    //    panel.Children.Add(new TextBlock { Text = "New Product ID:" });
    //    panel.Children.Add(newProductTxt);
    //    panel.Children.Add(activateBtn);
    //    panel.Children.Add(searchforlicenceBtn);
    //    panel.Children.Add(activationMsg);

    //    var dialog = new ContentDialog
    //    {
    //        Title = "Re-Activate License",
    //        Content = panel,
    //        CloseButtonText = "Cancel",
    //        XamlRoot = mainWindow.Content.XamlRoot
    //    };

    //    activateBtn.Click += async (s, e) =>
    //    {
    //        activateBtn.IsEnabled = false;
    //        activationMsg.Text = "Please wait...";

    //        string email = userEmailTxt.Text.Trim();
    //        string oldId = oldProductTxt.Text.Trim();
    //        string newId = newProductTxt.Text.Trim();

    //        if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(oldId) || string.IsNullOrWhiteSpace(newId))
    //        {
    //            activationMsg.Text = "All fields are required.";
    //            activateBtn.IsEnabled = true;
    //            return;
    //        }

    //        if (!IsValidEmail(email))
    //        {
    //            activationMsg.Text = "Invalid Email format.";
    //            activateBtn.IsEnabled = true;
    //            return;
    //        }

    //        try
    //        {

    //            await SendProductKEY_Date_Time(newId, email, newExpiryDate.ToString("dd-MM-yyyy"), GetMachineIdentifier(), oldId);
    //            SaveDetailsToConfig(newId, email, "Active", 0, newExpiryDate, "Online", DateTime.Now);
    //            File.WriteAllText(firstTimeOpenFilePath, "True");

    //            isActivated = true;
    //            dialog.Hide();
    //            await Task.Delay(500);
    //            await ShowDialog.ShowMsgBox("Success", "Product Re-Activated Successfully!", "OK", null, 1, mainWindow);
    //        }
    //        catch (Exception ex)
    //        {
    //            dialog.Hide();
    //            await Task.Delay(500);
    //            await ShowDialog.ShowMsgBox("Error", $"Re-Activation failed: {ex.Message}", "Ok", null, 1, mainWindow);
    //        }
    //    };

    //    //bool shouldsearchlicence = false;
    //    searchforlicenceBtn.Click += async (s, e) =>
    //    {

    //        dialog.Hide();
    //        isActivated = true; // just for showing messagebox from SearchforLicence
    //        await Task.Delay(500);
    //        Environment.Exit(0);

    //    };

    //    await dialog.ShowAsync();

    //    if (!isActivated)
    //    {
    //        Environment.Exit(0);
    //    }

    //    return newExpiryDate;
    //}

    //#endregion

    //#region Get github username
    //private async Task<string> GetAuthenticatedUserName()
    //{
    //    var response = await client.GetAsync("https://api.github.com/user");
    //    response.EnsureSuccessStatusCode();
    //    var content = await response.Content.ReadAsStringAsync();

    //    using var jsonDoc = JsonDocument.Parse(content);
    //    var root = jsonDoc.RootElement;
    //    return root.GetProperty("login").GetString()
    //        ?? throw new InvalidOperationException("Login property not found in JSON response");
    //}
    //#endregion

    //#region Validate Product ID Available
    //private async Task<AvailableProductInfo> ValidateProductIdAvailable(string productId)
    //{
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "Product_IDs/available_product_ids.json";
    //    string apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{filePath}";

    //    try
    //    {
    //        var response = await client.GetAsync(apiUrl);
    //        if (response.StatusCode == HttpStatusCode.NotFound)
    //        {
    //            return null;
    //        }

    //        response.EnsureSuccessStatusCode();
    //        string responseContent = await response.Content.ReadAsStringAsync();
    //        using var jsonDocument = JsonDocument.Parse(responseContent);
    //        var root = jsonDocument.RootElement;
    //        string contentBase64 = root.GetProperty("content").GetString();
    //        string jsonContent = Encoding.UTF8.GetString(Convert.FromBase64String(contentBase64));
    //        using var productsDocument = JsonDocument.Parse(jsonContent);
    //        var productsRoot = productsDocument.RootElement;

    //        var availableProductIds = productsRoot.GetProperty("available_product_ids").EnumerateArray();
    //        foreach (var product in availableProductIds)
    //        {
    //            if (product.GetProperty("ProductId").GetString() == productId)
    //            {
    //                return new AvailableProductInfo
    //                {
    //                    ProductId = product.GetProperty("ProductId").GetString(),
    //                    Days = product.GetProperty("Days").GetString()
    //                };
    //            }
    //        }

    //        return null; // Product not found
    //    }
    //    catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException || ex.CancellationToken.IsCancellationRequested)
    //    {
    //        // This is the REAL timeout case
    //        Console.WriteLine("Error in Validate ProductId Available IDs with timed out (slow internet)");
    //        return null;
    //    }
    //    catch (HttpRequestException ex)
    //    {
    //        Console.WriteLine("Error in Validate ProductId Available IDs with Network error: " + ex.Message);
    //        return null;
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine("Error in Validate ProductId Available IDs with Unexpected error in validation: " + ex.GetType().Name);
    //        return null;
    //    }
    //}
    //#endregion

    //#region Validate Product ID Sold
    //private async Task<bool> ValidateProductIdSold(string productId)
    //{
    //    //string repoOwner = "TacticsPro";
    //    //string repoOwner = await GetAuthenticatedUserName();
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "Product_IDs/sold_product_ids.json";
    //    string apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{filePath}";

    //    try
    //    {
    //        var response = await client.GetAsync(apiUrl);
    //        if (response.StatusCode == HttpStatusCode.NotFound)
    //        {
    //            return false; // No sold products file, so product ID is not sold
    //        }
    //        response.EnsureSuccessStatusCode();
    //        string responseContent = await response.Content.ReadAsStringAsync();
    //        using var jsonDocument = JsonDocument.Parse(responseContent);
    //        var root = jsonDocument.RootElement;

    //        string contentBase64 = root.GetProperty("content").GetString();
    //        string jsonContent = Encoding.UTF8.GetString(Convert.FromBase64String(contentBase64));
    //        using var productsDocument = JsonDocument.Parse(jsonContent);
    //        var productsRoot = productsDocument.RootElement;

    //        var soldProductIds = productsRoot.GetProperty("sold_product_ids").EnumerateArray();
    //        foreach (var product in soldProductIds)
    //        {
    //            if (product.GetProperty("ProductId").GetString() == productId)
    //            {

    //                return true;
    //            }
    //        }
    //        return false;
    //    }
    //    catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException || ex.CancellationToken.IsCancellationRequested)
    //    {
    //        // This is the REAL timeout case
    //        Console.WriteLine("License Validation failed at Sold Product IDs with timed out (slow internet)");
    //        return false;
    //    }
    //    catch (HttpRequestException ex)
    //    {
    //        Console.WriteLine("License Validation failed at Sold Product IDs with Network error: " + ex.Message);
    //        return false;
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine("License Validation failed at Sold Product IDs with Unexpected error in validation: " + ex.GetType().Name);
    //        return false;
    //    }
    //}
    //#endregion

    //#region Get Old Expiry Date
    //private async Task<DateTime> GetOldExpiry_from_sold_product(string productId)
    //{
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "Product_IDs/sold_product_ids.json";
    //    string apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{filePath}";

    //    try
    //    {
    //        var response = await client.GetAsync(apiUrl);
    //        if (response.StatusCode == HttpStatusCode.NotFound)
    //        {
    //            return DateTime.MinValue; // No sold products file, so return default
    //        }
    //        response.EnsureSuccessStatusCode();
    //        string responseContent = await response.Content.ReadAsStringAsync();
    //        using var jsonDocument = JsonDocument.Parse(responseContent);
    //        var root = jsonDocument.RootElement;

    //        string contentBase64 = root.GetProperty("content").GetString();
    //        string jsonContent = Encoding.UTF8.GetString(Convert.FromBase64String(contentBase64));
    //        using var productsDocument = JsonDocument.Parse(jsonContent);
    //        var productsRoot = productsDocument.RootElement;

    //        var soldProductIds = productsRoot.GetProperty("sold_product_ids").EnumerateArray();
    //        foreach (var product in soldProductIds)
    //        {
    //            if (product.GetProperty("ProductId").GetString() == productId)
    //            {
    //                string expiryDateStr = product.GetProperty("ExpiryDate").GetString();
    //                if (DateTime.TryParseExact(expiryDateStr, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime expiryDate))
    //                {
    //                    return expiryDate;
    //                }
    //                return DateTime.MinValue; // Invalid date format
    //            }
    //        }
    //        return DateTime.MinValue; // Product not found
    //    }
    //    catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException || ex.CancellationToken.IsCancellationRequested)
    //    {
    //        // This is the REAL timeout case
    //        Console.WriteLine("Error in GetOldExpiry_from_sold_product IDs with timed out (slow internet)");
    //        return DateTime.MinValue;
    //    }
    //    catch (HttpRequestException ex)
    //    {
    //        Console.WriteLine("Error in GetOldExpiry_from_sold_product IDs with Network error: " + ex.Message);
    //        return DateTime.MinValue;
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine("Error in GetOldExpiry_from_sold_product IDs with Unexpected error in validation: " + ex.GetType().Name);
    //        return DateTime.MinValue;
    //    }
    //}
    //#endregion

    //#region Get Expiry and Status
    //private async Task<(DateTime expiryDate, string licenceStatus)> GetExpiryAndStatusFromSoldProduct(string productId)
    //{
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "Product_IDs/sold_product_ids.json";
    //    string apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{filePath}";

    //    try
    //    {
    //        var response = await client.GetAsync(apiUrl);
    //        if (response.StatusCode == HttpStatusCode.NotFound)
    //        {
    //            return (DateTime.MinValue, string.Empty); // No sold products file, so return default
    //        }
    //        response.EnsureSuccessStatusCode();
    //        string responseContent = await response.Content.ReadAsStringAsync();
    //        using var jsonDocument = JsonDocument.Parse(responseContent);
    //        var root = jsonDocument.RootElement;

    //        string contentBase64 = root.GetProperty("content").GetString();
    //        string jsonContent = Encoding.UTF8.GetString(Convert.FromBase64String(contentBase64));
    //        using var productsDocument = JsonDocument.Parse(jsonContent);
    //        var productsRoot = productsDocument.RootElement;

    //        var soldProductIds = productsRoot.GetProperty("sold_product_ids").EnumerateArray();
    //        foreach (var product in soldProductIds)
    //        {
    //            if (product.GetProperty("ProductId").GetString() == productId)
    //            {
    //                string expiryDateStr = product.GetProperty("ExpiryDate").GetString();
    //                string licenceStatus = product.GetProperty("LicenceStatus").GetString() ?? string.Empty;
    //                if (DateTime.TryParseExact(expiryDateStr, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime expiryDate))
    //                {
    //                    return (expiryDate, licenceStatus);
    //                }
    //                return (DateTime.MinValue, string.Empty); // Invalid date format
    //            }
    //        }
    //        return (DateTime.MinValue, string.Empty); // Product not found
    //    }
    //    catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException || ex.CancellationToken.IsCancellationRequested)
    //    {
    //        // This is the REAL timeout case
    //        Console.WriteLine("Error in GetExpiry And Status From Sold Product IDs with timed out (slow internet)");
    //        return (DateTime.MinValue, string.Empty);
    //    }
    //    catch (HttpRequestException ex)
    //    {
    //        Console.WriteLine("Error in GetExpiry AndStatus From Sold Product IDs with Network error: " + ex.Message);
    //        return (DateTime.MinValue, string.Empty);
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine("Error in GetExpiry AndStatus From Sold Product IDs with Unexpected error in validation: " + ex.GetType().Name);
    //        return (DateTime.MinValue, string.Empty);
    //    }
    //}
    //#endregion

    #region Update Github Files
    private async Task UpdateGitHubFiles(string productId, string emailId, string activationTime, string expiryDate, string oldProductId = null)
    {
        string repoOwner = "TacticsPro";
        string repoName = "Office_Tools_Private";
        string liteVersionPath = "Product_IDs/lite_version.json";
        string liteVersionApiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{liteVersionPath}";
        string currentMachineId = GetMachineIdentifier();

        try
        {
            // Update Product_IDs/lite_version.json
            string liteVersionJsonContent;
            string liteVersionFileSha;
            var liteVersionResponse = await client.GetAsync(liteVersionApiUrl);
            if (liteVersionResponse.StatusCode == HttpStatusCode.NotFound)
            {
                liteVersionJsonContent = JsonSerializer.Serialize(new liteVersionProducts { liteVersionProductIds = new List<Product>() }, liteVersionProductsJsonContext.Default.liteVersionProducts);
                liteVersionFileSha = null;
            }
            else
            {
                liteVersionResponse.EnsureSuccessStatusCode();
                string responseContent = await liteVersionResponse.Content.ReadAsStringAsync();
                using var jsonDocument = JsonDocument.Parse(responseContent);
                var root = jsonDocument.RootElement;
                liteVersionFileSha = root.GetProperty("sha").GetString();
                string contentBase64 = root.GetProperty("content").GetString();
                liteVersionJsonContent = Encoding.UTF8.GetString(Convert.FromBase64String(contentBase64));
            }

            // Deserialize using source-generated context
            var liteVersionProducts = JsonSerializer.Deserialize(liteVersionJsonContent, liteVersionProductsJsonContext.Default.liteVersionProducts)
                ?? throw new JsonException("Failed to deserialize liteVersion_products.json");

            var existingProduct = liteVersionProducts.liteVersionProductIds.FirstOrDefault(p => p.ProductId == productId);
            if (existingProduct != null)
            {
                existingProduct.ActivationTime = activationTime;
                existingProduct.ExpiryDate = expiryDate;
                existingProduct.LicenceStatus = "Active";
            }
            else
            {
                if (!string.IsNullOrEmpty(oldProductId))
                {
                    liteVersionProducts.liteVersionProductIds.RemoveAll(p => p.ProductId == oldProductId);
                }
                liteVersionProducts.liteVersionProductIds.Add(new Product
                {
                    ProductId = productId,
                    EmailId = emailId,
                    ActivatedMachineId = currentMachineId,
                    ActivationTime = activationTime,
                    ExpiryDate = expiryDate,
                    LicenceStatus = "Active"
                });
            }
            string updatedliteVersionJson = JsonSerializer.Serialize(liteVersionProducts, liteVersionProductsJsonContext.Default.liteVersionProducts);
            string updatedliteVersionContentBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(updatedliteVersionJson));

            var liteVersionUpdatePayload = new GitHubUpdatePayload
            {
                Message = $"Add product {productId} for {emailId}" + (oldProductId != null ? $" and remove old product {oldProductId}" : ""),
                Content = updatedliteVersionContentBase64,
                Sha = liteVersionFileSha
            };
            var liteVersionContent = new StringContent(
                JsonSerializer.Serialize(liteVersionUpdatePayload, GitHubUpdatePayloadJsonContext.Default.GitHubUpdatePayload),
                Encoding.UTF8,
                "application/json");
            var liteVersionUpdateResponse = await client.PutAsync(liteVersionApiUrl, liteVersionContent);

            if (liteVersionUpdateResponse.StatusCode == HttpStatusCode.Conflict)
            {
                var retryResponse = await client.GetAsync(liteVersionApiUrl);
                retryResponse.EnsureSuccessStatusCode();
                string retryContent = await retryResponse.Content.ReadAsStringAsync();
                using var retryDocument = JsonDocument.Parse(retryContent);
                liteVersionFileSha = retryDocument.RootElement.GetProperty("sha").GetString();
                liteVersionUpdatePayload = new GitHubUpdatePayload
                {
                    Message = $"Add product {productId} for {emailId}" + (oldProductId != null ? $" and remove old product {oldProductId}" : ""),
                    Content = updatedliteVersionContentBase64,
                    Sha = liteVersionFileSha
                };
                liteVersionContent = new StringContent(
                    JsonSerializer.Serialize(liteVersionUpdatePayload, GitHubUpdatePayloadJsonContext.Default.GitHubUpdatePayload),
                    Encoding.UTF8,
                    "application/json");
                liteVersionUpdateResponse = await client.PutAsync(liteVersionApiUrl, liteVersionContent);
            }
            liteVersionUpdateResponse.EnsureSuccessStatusCode();

        }
        catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException || ex.CancellationToken.IsCancellationRequested)
        {
            // This is the REAL timeout case
            Console.WriteLine("GitHub Update Failed with timed out (slow internet)");
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine("GitHub Update Failed with Network error: " + ex.Message);
            throw new Exception($"GitHub Update Failed: {ex.Message}\nResponse: {ex.InnerException?.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("GitHub Update Failed WITH Unexpected error in validation: " + ex.GetType().Name);
        }
    }
    #endregion

    #region Validate email
    private bool IsValidEmail(string email)
    {
        try
        {
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }
    #endregion

    //#region Getting Update latest version info
    //public static async Task<VersionInfo> GetLatestVersionInfoAsync()
    //{
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "Office_Tools_Latest_Version/Office_Tools_WinUi3/MSIX/latest_version_winui3.json";
    //    string apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{filePath}";

    //    try
    //    {
    //        // Set User-Agent header as required by GitHub API
    //        client.DefaultRequestHeaders.UserAgent.TryParseAdd("OfficeToolsFinder");

    //        // Make the HTTP GET request to the GitHub API
    //        HttpResponseMessage response = await client.GetAsync(apiUrl);
    //        response.EnsureSuccessStatusCode(); // Throws if the response is not successful

    //        // Read the response content as a string
    //        string jsonResponse = await response.Content.ReadAsStringAsync();

    //        // Parse the GitHub API response (which includes base64-encoded content)
    //        using JsonDocument doc = JsonDocument.Parse(jsonResponse);
    //        JsonElement root = doc.RootElement;

    //        // Extract the base64-encoded content
    //        string base64Content = root.GetProperty("content").GetString();
    //        // Decode the base64 content to get the JSON string
    //        byte[] decodedBytes = Convert.FromBase64String(base64Content);
    //        string jsonContent = Encoding.UTF8.GetString(decodedBytes);

    //        // Parse the decoded JSON to extract VersionInfo using the source-generated context
    //        VersionInfo versionInfo = JsonSerializer.Deserialize(jsonContent, SourceGenContext.Default.VersionInfo);

    //        return versionInfo;
    //    }
    //    catch (HttpRequestException ex)
    //    {
    //        throw new Exception("Failed to fetch version info from GitHub.", ex);
    //    }
    //    catch (JsonException ex)
    //    {
    //        throw new Exception("Failed to parse version info JSON.", ex);
    //    }
    //    catch (Exception ex)
    //    {
    //        throw new Exception("An error occurred while retrieving version info.", ex);
    //    }
    //}
    //#endregion

    #region Getting info path
    public static async Task<InfoPath> GetInfoPath()
    {
        string repoOwner = "TacticsPro";
        string repoName = "Office_Tools_Private";
        string filePath = "Version_Info/info_path.json";
        string apiUrl = $"https://api.github.com/repos/{repoOwner}/{repoName}/contents/{filePath}";

        try
        {
            // Set User-Agent header as required by GitHub API
            client.DefaultRequestHeaders.UserAgent.TryParseAdd("OfficeToolsFinder");

            // Make the HTTP GET request to the GitHub API
            HttpResponseMessage response = await client.GetAsync(apiUrl);
            response.EnsureSuccessStatusCode(); // Throws if the response is not successful

            // Read the response content as a string
            string jsonResponse = await response.Content.ReadAsStringAsync();

            // Parse the GitHub API response (which includes base64-encoded content)
            using JsonDocument doc = JsonDocument.Parse(jsonResponse);
            JsonElement root = doc.RootElement;

            // Extract the base64-encoded content
            string base64Content = root.GetProperty("content").GetString();
            // Decode the base64 content to get the JSON string
            byte[] decodedBytes = Convert.FromBase64String(base64Content);
            string jsonContent = Encoding.UTF8.GetString(decodedBytes);

            // Parse the decoded JSON to extract infopath using the source-generated context
            InfoPath infopath = JsonSerializer.Deserialize(jsonContent, SourceGenContext.Default.InfoPath);

            return infopath;
        }
        catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException || ex.CancellationToken.IsCancellationRequested)
        {
            throw new Exception("Timed out (slow internet).", ex);
        }
        catch (HttpRequestException ex)
        {
            throw new Exception("Failed to fetch version info from GitHub.", ex);
        }
        catch (JsonException ex)
        {
            throw new Exception("Failed to parse version info JSON.", ex);
        }
        catch (Exception ex)
        {
            throw new Exception("An error occurred while retrieving version info.", ex);
        }
    }
    #endregion

    //#region Write HSN.json
    //public async static Task DownloadHSNFile() // for bigger size any type of file
    //{
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "HSN/HSN.json";
    //    string rawUrl = $"https://raw.githubusercontent.com/{repoOwner}/{repoName}/main/{filePath}";
    //    try
    //    {
    //        client.DefaultRequestHeaders.UserAgent.TryParseAdd("OfficeToolsFinder");

    //        string jsonContent = await client.GetStringAsync(rawUrl);
    //        File.WriteAllText(hsnTempPath, jsonContent);
    //    }
    //    catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException || ex.CancellationToken.IsCancellationRequested)
    //    {
    //        throw new Exception("Timed out (slow internet).", ex);
    //    }
    //    catch (HttpRequestException ex)
    //    {
    //        throw new Exception("Failed to fetch from GitHub.", ex);
    //    }
    //    catch (JsonException ex)
    //    {
    //        throw new Exception("Failed to parse version info JSON.", ex);
    //    }
    //    catch (Exception ex)
    //    {
    //        throw new Exception("An error occurred.", ex);
    //    }
    //}
    //#endregion

    //#region Download File
    //public async static Task DownloadHSNFile() // for bigger size any type of file
    //{
    //    string repoOwner = "TacticsPro";
    //    string repoName = "Office_Tools_Private";
    //    string filePath = "HSN/HSN.json";
    //    string rawUrl = $"https://raw.githubusercontent.com/{repoOwner}/{repoName}/main/{filePath}";

    //    client.DefaultRequestHeaders.UserAgent.TryParseAdd("OfficeToolsFinder");

    //    using var response = await client.GetAsync(rawUrl, HttpCompletionOption.ResponseHeadersRead);
    //    response.EnsureSuccessStatusCode();

    //    await using var fs = new FileStream(hsnTempPath, FileMode.Create, FileAccess.Write, FileShare.None);
    //    await response.Content.CopyToAsync(fs);
    //}

    //#endregion

}

#region Other Classes
internal class MachineLicenseInfo
{
    [JsonPropertyName("machineId")]
    public string machineId { get; set; } = string.Empty;

    [JsonPropertyName("expiryDate")]
    public DateTime expiryDate
    {
        get; set;
    }

    [JsonPropertyName("productId")]
    public string productId { get; set; } = string.Empty;

    [JsonPropertyName("emailId")]
    public string emailId { get; set; } = string.Empty;

    [JsonPropertyName("licenceStatus")]
    public string licenceStatus { get; set; } = string.Empty;

    [JsonPropertyName("offlineruncount")]
    public int offlineruncount
    {
        get; set;
    }

    [JsonPropertyName("ActivateMode")]
    public string ActivateMode { get; set; } = string.Empty;

    [JsonPropertyName("Activationtime")]
    public DateTime Activationtime
    {
        get; set;
    }

    public class Product
    {
        [JsonPropertyName("ProductId")]
        public string ProductId
        {
            get; set;
        }

        [JsonPropertyName("EmailId")]
        public string EmailId
        {
            get; set;
        }

        [JsonPropertyName("ActivatedMachineId")]
        public string ActivatedMachineId
        {
            get; set;
        }

        [JsonPropertyName("ActivationTime")]
        public string ActivationTime
        {
            get; set;
        }

        [JsonPropertyName("ExpiryDate")]
        public string ExpiryDate
        {
            get; set;
        }

        [JsonPropertyName("LicenceStatus")]
        public string LicenceStatus
        {
            get; set;
        }
    }

    public class AvailableProductInfo
    {
        [JsonPropertyName("ProductId")]
        public string ProductId
        {
            get; set;
        }

        [JsonPropertyName("Days")]
        public string Days
        {
            get; set;
        }
    }

    public class SoldProducts
    {
        [JsonPropertyName("sold_product_ids")]
        public List<Product> SoldProductIds { get; set; } = new List<Product>();
    }
    public class liteVersionProducts
    {
        [JsonPropertyName("lite_version")]
        public List<Product> liteVersionProductIds { get; set; } = new List<Product>();
    }

    public class AvailableProducts
    {
        [JsonPropertyName("available_product_ids")]
        public List<AvailableProductInfo> AvailableProductIds { get; set; } = new List<AvailableProductInfo>();
    }

    public class DeletedProducts
    {
        [JsonPropertyName("deleted_product_ids")]
        public List<string> DeletedProductIds { get; set; } = new List<string>();
    }

    public class GitHubUpdatePayload
    {
        [JsonPropertyName("message")]
        public string Message
        {
            get; set;
        }

        [JsonPropertyName("content")]
        public string Content
        {
            get; set;
        }

        [JsonPropertyName("sha")]
        public string Sha
        {
            get; set;
        }
    }

    public class PendingEmailData
    {
        [JsonPropertyName("product")]
        public string product
        {
            get; set;
        }

        [JsonPropertyName("userEmailId")]
        public string userEmailId
        {
            get; set;
        }

        [JsonPropertyName("expiryDate")]
        public string expiryDate
        {
            get; set;
        }

        [JsonPropertyName("machineID")]
        public string machineID
        {
            get; set;
        }

        [JsonPropertyName("oldProductId")]
        public string oldProductId
        {
            get; set;
        }
    }

}
//public class VersionInfo
//{
//    public string latest_version { get; set; }
//    public string download_url { get; set; }
//    public string release_notes { get; set; }
//    public string info_path { get; set; }
//}
public class InfoPath
{
    public string info_path
    {
        get; set;
    }
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo))]
internal partial class MachineLicenseInfoJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.Product))]
internal partial class ProductJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.AvailableProductInfo))]
internal partial class AvailableProductInfoJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.SoldProducts))]
internal partial class SoldProductsJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.liteVersionProducts))]
internal partial class liteVersionProductsJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.AvailableProducts))]
internal partial class AvailableProductsJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.DeletedProducts))]
internal partial class DeletedProductsJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.GitHubUpdatePayload))]
internal partial class GitHubUpdatePayloadJsonContext : JsonSerializerContext
{
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(MachineLicenseInfo.PendingEmailData))]
internal partial class PendingEmailDataJsonContext : JsonSerializerContext
{
}

//[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
//[JsonSerializable(typeof(VersionInfo))]
//internal partial class SourceGenContext : JsonSerializerContext
//{
//}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(InfoPath))]
internal partial class SourceGenContext : JsonSerializerContext
{
}

#endregion