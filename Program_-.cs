using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging; // For Open XML SDK
using System.Linq;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel; // For XSSFWorkbook
using NPOI.HSSF.UserModel; // For HSSFWorkbook (if used)
using NPOI.XWPF.UserModel; // For XWPFDocument (if used)
using NPOI.HPSF; // For SummaryInformation (if used)
using iTextSharp.text.pdf;
using MetadataExtractor;
using System.Security.Principal;
using System.IO.Compression;
using Newtonsoft.Json;

public class FileMetadata
{
    [JsonProperty(Order = 1)]
    public string Owner { get; set; }

    [JsonProperty(Order = 2)]
    public string Parent { get; set; }

    [JsonProperty(Order = 3)]
    public string Size { get; set; }

    [JsonProperty(Order = 4)]
    public DateTime CreatedDate { get; set; }  // changed to DateTime

    [JsonProperty(Order = 5)]
    public TimeSpan CreatedTime { get; set; }  // changed to TimeSpan

    [JsonProperty(Order = 6)]
    public DateTime AccessedDate { get; set; }  // changed to DateTime

    [JsonProperty(Order = 7)]
    public TimeSpan AccessedTime { get; set; }  // changed to TimeSpan

    [JsonProperty(Order = 8)]
    public DateTime ModifiedDate { get; set; }  // changed to DateTime

    [JsonProperty(Order = 9)]
    public TimeSpan ModifiedTime { get; set; }  // changed to TimeSpan

    [JsonIgnore]
    public ConcurrentDictionary<string, FileMetadata> Children { get; set; } = new ConcurrentDictionary<string, FileMetadata>();
}




public static class DirectoryScanner
{
    public static async Task<FileMetadata> ScanDirectoryAsync(string path)
{
    var info = new DirectoryInfo(path);
    var result = new FileMetadata();

    var fileTasks = info.GetFiles().Select(file => ProcessFileAsync(file)).ToList();
    var fileResults = await Task.WhenAll(fileTasks);

    foreach (var (fileName, fileInfo) in fileResults)
    {
        if (fileName != null && fileInfo != null)
        {
            result.Children[fileName] = fileInfo;
        }
    }

    foreach (var directory in info.GetDirectories())
    {
        var directoryInfo = await ScanDirectoryAsync(directory.FullName);
        foreach (var entry in directoryInfo.Children)
        {
            result.Children[entry.Key] = entry.Value;
        }
    }

    return result;
}

    
    public static async Task<(DateTime CreationTime, DateTime LastWriteTime, long Length)> GetFilePropertiesAsync(FileInfo file)
{
    const int maxRetries = 3;
    const int delayBase = 2; // Base delay in seconds, will be doubled every retry

    for (int i = 0; i < maxRetries; i++)
    {
        try
        {
            // Get file properties (size, creation date, modification date, etc.)
            DateTime creationTime = file.CreationTime;
            DateTime lastWriteTime = file.LastWriteTime;
            long length = file.Length;

            // Return the file properties
            return (creationTime, lastWriteTime, length);
        }
        catch (System.Net.NetworkInformation.NetworkInformationException)
        {
            // If a network disturbance occurs, wait for an exponentially increasing delay
            int delay = (int)Math.Pow(delayBase, i);
            await Task.Delay(TimeSpan.FromSeconds(delay));
            continue; // Try again
        }
        catch
        {
            // If getting the file properties fails for a non-network-related reason, handle appropriately
            // For now, we'll just rethrow the exception
            throw;
        }
    }
    // If all else fails, return default properties or throw an exception
    throw new Exception("Failed to get file properties after multiple attempts.");
}

    private static async Task<ConcurrentDictionary<string, FileMetadata>> GetCompressedFileContentsAsync(FileInfo compressedFile)
    {
        var contents = new ConcurrentDictionary<string, FileMetadata>();

        using (var archive = ZipFile.OpenRead(compressedFile.FullName))
        {
            await Task.WhenAll(archive.Entries.Select(async entry =>
            {
                contents[entry.Name] = new FileMetadata
                {
                    Size = GetReadableSize(entry.Length)
                };
            }));
        }

        return contents;
    }

    private static async Task<(string FileName, FileMetadata Info)> ProcessFileAsync(FileInfo file)
{
    try
    {
        if (file.Name.StartsWith("~$") || file.Extension == ".lnk")
        {
            return (null, null);
        }

        var (creationDate, modifiedDate, fileSize) = await GetFilePropertiesAsync(file);
        var creationTime = creationDate.TimeOfDay;
        var modifiedTime = modifiedDate.TimeOfDay;

        var fileInfo = new FileMetadata
        {
            Parent = file.DirectoryName,
            CreatedDate = creationDate.Date,  // Store the date portion as DateTime
            CreatedTime = creationTime,  // Store the time portion as TimeSpan
            AccessedDate = creationDate.Date,  // Store the date portion as DateTime
            AccessedTime = creationTime,  // Store the time portion as TimeSpan
            ModifiedDate = modifiedDate.Date,  // Store the date portion as DateTime
            ModifiedTime = modifiedTime,  // Store the time portion as TimeSpan
            Size = GetReadableSize(fileSize)
        };


        if (file.Extension == ".zip")
        {
            fileInfo.Children = await GetCompressedFileContentsAsync(file);
        }

        var ownerTask = GetOwnerAsync(file).ContinueWith(t =>
        {
            fileInfo.Owner = t.Result;
        });
        await ownerTask;
        
        return (file.Name, fileInfo);
    }
    catch (Exception ex)
    {
        // Handle the exception
        // For example, you can log the error and the file that caused it
        Console.WriteLine($"An error occurred while processing file {file.Name}: {ex.Message}");
        return (null, null);
    }
}





    private static async Task<long> GetDirectorySizeAsync(DirectoryInfo directoryInfo)
    {
        long size = 0;
        FileInfo[] files = directoryInfo.GetFiles();
        size += files.Sum(file => file.Length);

        var subdirectorySizes = await Task.WhenAll(directoryInfo.GetDirectories().Select(GetDirectorySizeAsync));
        size += subdirectorySizes.Sum();

        return size;
    }
    public static class SupportedFormats
    {
        public static List<string> ImageTypes = new List<string> { ".jpg", ".jpeg", ".png", /* other image types */ };
        public static List<string> VideoTypes = new List<string> { ".mp4", ".avi", /* other video types */ };
        public static List<string> AudioTypes = new List<string> { ".mp3", ".wav", /* other audio types */ };
    }
public static async Task<string> GetOwnerAsync(FileInfo file)
{
    const int maxRetries = 3;
    const int delayBase = 2; // Base delay in seconds, will be doubled every retry

    for (int i = 0; i < maxRetries; i++)
    {
        try
        {
            string[] supportedExtensions = new[] { ".docx", ".dotx", ".docm", ".dotm", ".xlsx", ".xltx", ".xlsm", ".xltm", ".pptx", ".potx", ".pptm", ".potm" };

            if (supportedExtensions.Contains(file.Extension))
            {
                using (Package package = Package.Open(file.FullName, FileMode.Open, FileAccess.Read))
                {
                    PackageProperties packageProperties = package.PackageProperties;
                    return packageProperties.Creator;
                }
            }
            
            else if (file.Extension == ".doc")
            {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            doc.LoadFromFile(file.FullName);
            string author = doc.BuiltinDocumentProperties.Author;
            return author;
            }
            else if (file.Extension == ".ppt")
            {
                Spire.Presentation.Presentation ppt = new Spire.Presentation.Presentation();
                ppt.LoadFromFile(file.FullName);
                string author = ppt.DocumentProperty.Application;
                return author;
            }
            else if (file.Extension == ".xls")
            {
                using (var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    HSSFWorkbook workbook = new HSSFWorkbook(stream);
                    SummaryInformation si = workbook.SummaryInformation;
                    return si.Author;
                }
            }            
            else if (file.Extension == ".pdf")
            {
                using var reader = new PdfReader(file.FullName);
                string author = reader.Info["Author"];
                return author;
            }
            else if (SupportedFormats.ImageTypes.Contains(file.Extension) ||
                     SupportedFormats.VideoTypes.Contains(file.Extension) ||
                     SupportedFormats.AudioTypes.Contains(file.Extension))
            {
                var directories = ImageMetadataReader.ReadMetadata(file.FullName);
                foreach (var directory in directories)
                {
                    foreach (var tag in directory.Tags)
                    {
                        if (tag.Name == "Author")
                        {
                            return tag.Description;
                        }
                    }
                }
                throw new Exception("No author tag found in metadata");
            }

        }
        catch
        {
            // If getting the author from the file properties fails, proceed to the next approach
        }
        try
        {
            // Next, try to get the owner from the file's access control
            var owner = file.GetAccessControl().GetOwner(typeof(NTAccount));
            string fullOwner = owner != null ? owner.ToString() : null;
            if (fullOwner != null)
            {
                string[] ownerParts = fullOwner.Split(new string[] {"\\", "//"}, StringSplitOptions.None);
                if (ownerParts.Length > 1)
                {
                    return ownerParts[ownerParts.Length - 1];  // Return only the username part
                }
                else
                {
                    return fullOwner;  // If there is no separator in the owner string, return the whole string
                }
            }
            else
            {
                // handle the case when fullOwner is null, return a default value or throw an exception
                return "Unknown";
            }
        }
        catch (System.Net.NetworkInformation.NetworkInformationException)
        {
            // If a network disturbance occurs, wait for an exponentially increasing delay
            int delay = (int)Math.Pow(delayBase, i);
            await Task.Delay(TimeSpan.FromSeconds(delay));
            continue; // Try again
        }
        catch
        {
            // If getting the owner from the file's access control fails for a non-network-related reason, proceed to the next approach
        }
    }
    // If all else fails, return "Unknown"
        return "Unknown";
}
    public static string GetReadableSize(long length)
    {
        string[] sizes = { "B", "KB", "MB", "GB", "TB" };
        int order = 0;
        while (length >= 1024 && order < sizes.Length - 1)
        {
            order++;
            length = length / 1024;
        }
        return $"{length:0.##} {sizes[order]}";
    }


}

public class Program
{
    public static async Task Main()
{
    var root = await DirectoryScanner.ScanDirectoryAsync(@"C:\\Users\\dimit\\Downloads");
    var json = JsonConvert.SerializeObject(root.Children, Newtonsoft.Json.Formatting.Indented);
    await File.WriteAllTextAsync(@"C:\\Users\\dimit\\source\\repos\\ConsoleApp1\\ConsoleApp1\\output.json", json);
}

}
