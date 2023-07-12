using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using Newtonsoft.Json;
using System.Security.AccessControl;
using System.Collections.Generic;
using System.IO.Compression;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using MetadataExtractor;
using System.Reflection.Metadata;
using System.Security.Principal;
using iTextSharp.text.pdf;

public class FileOrFolderInfo
{
    public string Type { get; set; }
    public string Parent { get; set; }
    public string Size { get; set; }
    public ConcurrentDictionary<string, FileOrFolderInfo> Children { get; set; } = new ConcurrentDictionary<string, FileOrFolderInfo>(); // Changed to ConcurrentDictionary
    public string Owner { get; set; }
    public DateTime CreatedDate { get; set; }
    public TimeSpan CreatedTime { get; set; }
    public DateTime AccessedDate { get; set; }
    public TimeSpan AccessedTime { get; set; }
    public DateTime ModifiedDate { get; set; }
    public TimeSpan ModifiedTime { get; set; }


}

public static class DirectoryScanner
{
    public static async Task<FileOrFolderInfo> ScanDirectoryAsync(string path)
    {
        var info = new DirectoryInfo(path);
        var result = new FileOrFolderInfo
        {
            Type = "folder",
            Parent = info.Parent?.FullName,
            Size = (await GetDirectorySizeAsync(info)).ToString()
        };

        var fileTasks = info.GetFiles().Select(async file =>
        {
            if (file.Name.StartsWith("~$") || file.Extension == ".lnk")
            {
                return;
            }

            var creationDate = file.CreationTime.Date;
            var creationTime = file.CreationTime.TimeOfDay;
            var modifiedDate = file.LastWriteTime.Date;
            var modifiedTime = file.LastWriteTime.TimeOfDay;
            var owner = await GetOwnerAsync(file);  // Await the GetOwnerAsync call


            if (file.Extension == ".zip")
            {
                result.Children[file.Name] = new FileOrFolderInfo
                {
                    Type = "compressed file",
                    Parent = info.FullName,
                    Owner = owner,
                    CreatedDate = creationDate,
                    CreatedTime = creationTime,
                    AccessedDate = creationDate,
                    AccessedTime = creationTime,
                    ModifiedDate = modifiedDate,
                    ModifiedTime = modifiedTime,
                    Size = GetReadableSize(file.Length),
                    Children = await GetCompressedFileContentsAsync(file)
                };
            }
            else
            {
                result.Children[file.Name] = new FileOrFolderInfo
                {
                    Type = "file",
                    Parent = info.FullName,
                    Owner = owner,
                    CreatedDate = creationDate,
                    CreatedTime = creationTime,
                    AccessedDate = creationDate,
                    AccessedTime = creationTime,
                    ModifiedDate = modifiedDate,
                    ModifiedTime = modifiedTime,
                    Size = (file.Length / 1024f / 1024f).ToString() + " MB"
                };
            }
        });

        var directoryTasks = info.GetDirectories().Select(async directory =>
        {
            result.Children[directory.Name] = await ScanDirectoryAsync(directory.FullName);
        });

        await Task.WhenAll(fileTasks.Concat(directoryTasks));

        return result;
    }



    private static async Task<ConcurrentDictionary<string, FileOrFolderInfo>> GetCompressedFileContentsAsync(FileInfo compressedFile)
    {
        var contents = new ConcurrentDictionary<string, FileOrFolderInfo>();

        using (var archive = ZipFile.OpenRead(compressedFile.FullName))
        {
            await Task.WhenAll(archive.Entries.Select(async entry =>
            {
                contents[entry.Name] = new FileOrFolderInfo
                {
                    Type = "file in compressed file",
                    Size = GetReadableSize(entry.Length)
                };
            }));
        }

        return contents;
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
        try
        {
            if (file.Extension == ".docx" || file.Extension == ".xlsx" || file.Extension == ".pptx")
            {
                using (var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var document = WordprocessingDocument.Open(stream, false))
                {
                    return document.PackageProperties.Creator;
                }
            }
            else if (file.Extension == ".pdf")
            {
                using (var reader = new iTextSharp.text.pdf.PdfReader(file.FullName))
                {
                    var author = reader.Info["Author"];
                    return author;
                }
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
                return "Unknown";
            }
            else
            {
                // This operation is inherently synchronous
                return file.GetAccessControl().GetOwner(typeof(NTAccount)).ToString();
            }
        }
        catch
        {
            return "Unable to retrieve author";
        }
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
        var result = new Dictionary<string, FileOrFolderInfo>
        {
            ["root"] = await DirectoryScanner.ScanDirectoryAsync(@"C:\Users\dimit\OneDrive")
        };
        var json = JsonConvert.SerializeObject(result, Newtonsoft.Json.Formatting.Indented);
        await File.WriteAllTextAsync(@"C:\\Users\\dimit\\Documents\\output.json", json);
    }
}

