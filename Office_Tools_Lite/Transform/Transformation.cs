using System.Security.Cryptography;
using Office_Tools_Lite.Task_Helper;
namespace Office_Tools_Lite;

public class Transformation
{
    // Dictionary to store transformed file contents in memory
    private static Dictionary<string, byte[]> transformedFileContents = new Dictionary<string, byte[]>();

    public async Task TransformFiles()
    {
        var prompt = new Prompt();
        // Use await to get the string results from async methods
        string firstpart = await prompt.PrompterK();
        string secondpart = await prompt.PrompterI();

        byte[] key = Convert.FromBase64String(firstpart);
        byte[] iv = Convert.FromBase64String(secondpart);

        string DynamicLibrariesFolderPath = Path.Combine(AppContext.BaseDirectory, "DynamicLibraries");

        Console.WriteLine($"DynamicLibraries folder path: {DynamicLibrariesFolderPath}");

        if (!Directory.Exists(DynamicLibrariesFolderPath))
        {
            Console.WriteLine("DynamicLibraries folder not found.");
            return;
        }

        string[] filesToTransform = Directory.GetFiles(DynamicLibrariesFolderPath, "*.dll");

        if (filesToTransform.Length == 0)
        {
            Console.WriteLine("No .dll files found in the DynamicLibraries folder.");
            return;
        }

        using (Aes aesAlg = Aes.Create())
        {
            aesAlg.Key = key;
            aesAlg.IV = iv;

            ICryptoTransform transformer = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);

            foreach (var fileName in filesToTransform)
            {
                try
                {
                    Console.WriteLine($"Transforming file: {fileName}");
                    byte[] transformationFile = File.ReadAllBytes(fileName);

                    // Transform and store in memory
                    byte[] transformedFile = PerformDecryption(transformer, transformationFile);

                    string fileKey = Path.GetFileName(fileName); // Use the filename as the dictionary key
                    transformedFileContents[fileKey] = transformedFile;

                    Console.WriteLine($"File Transformed and stored in memory: {fileKey}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to Transform file {fileName}: {ex.Message}");
                }
            }
        }

        Console.WriteLine("Transformation completed.");
    }

    private static byte[] PerformDecryption(ICryptoTransform transformer, byte[] data)
    {
        using (MemoryStream ms = new MemoryStream())
        {
            using (CryptoStream cs = new CryptoStream(ms, transformer, CryptoStreamMode.Write))
            {
                cs.Write(data, 0, data.Length);
                cs.FlushFinalBlock();
            }
            return ms.ToArray();
        }
    }

    // Method to retrieve transformed content by file name
    public static byte[] GetTransformedFileContent(string fileName)
    {
        return transformedFileContents.ContainsKey(fileName) ? transformedFileContents[fileName] : null;
    }

    public static void Clear()
    {
        transformedFileContents.Clear();
    }

}
