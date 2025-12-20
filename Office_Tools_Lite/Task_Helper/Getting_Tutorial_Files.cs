namespace Office_Tools_Lite.Task_Helper;
public static class Getting_Tutorial_Files
{
    private static string? htmlFileUri;
    //private static string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
    //private static string targetTutorialsFolder = Path.Combine(localAppData, "Office_Tools_Lite", "Tutorials");
    private static string? targetFilePath;
    public static string? GettingTutorialFiles()
    {
        // Path to the Tutorials folder in the app's installation directory
        string sourceTutorialsFolder = Path.Combine(AppContext.BaseDirectory, "Tutorials");

        // Verify the Tutorials folder exists
        if (Directory.Exists(sourceTutorialsFolder))
        {
            // Create the target folder if it doesn't exist
            //Directory.CreateDirectory(targetTutorialsFolder);

            //// Copy the entire Tutorials folder and its contents to the target location
            //CopyDirectory(sourceTutorialsFolder, targetTutorialsFolder);

            //// Path to index.html in the target folder
            //targetFilePath = Path.Combine(targetTutorialsFolder, "index.html");
            //htmlFileUri = new Uri($"file:///{targetFilePath.Replace("\\", "/")}").AbsoluteUri;
            //return htmlFileUri;


            targetFilePath = Path.Combine(sourceTutorialsFolder, "index.html");
            htmlFileUri = new Uri($"file:///{targetFilePath.Replace("\\", "/")}").AbsoluteUri;
            return htmlFileUri;


        }
        else
        {
            return null;
        }
    }

    // Helper method to copy a directory and its contents recursively
    private static void CopyDirectory(string sourceDir, string targetDir)
    {
        // Create the target directory if it doesn't exist
        Directory.CreateDirectory(targetDir);

        // Copy all files in the source directory
        foreach (string file in Directory.GetFiles(sourceDir))
        {
            string fileName = Path.GetFileName(file);
            string targetFilePath = Path.Combine(targetDir, fileName);
            File.Copy(file, targetFilePath, true);
        }

        // Copy all subdirectories recursively
        foreach (string subDir in Directory.GetDirectories(sourceDir))
        {
            string subDirName = Path.GetFileName(subDir);
            string targetSubDir = Path.Combine(targetDir, subDirName);
            CopyDirectory(subDir, targetSubDir);
        }
    }
}
