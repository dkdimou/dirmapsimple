using System;
using System.IO;
using Newtonsoft.Json;
using System.Security.AccessControl;
using System.Collections.Generic;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Words;
using Aspose.Slides;
using Aspose.Pdf;
using System.Reflection.Metadata;
using System.Collections.Concurrent;

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
    public static FileOrFolderInfo ScanDirectory(string path)
    {
        var info = new DirectoryInfo(path);
        var result = new FileOrFolderInfo
        {
            Type = "folder",
            Parent = info.Parent?.FullName,
            Size = GetDirectorySize(info).ToString(),
            Children = new ConcurrentDictionary<string, FileOrFolderInfo>()  // Changed to ConcurrentDictionary
        };

        Parallel.ForEach(info.GetFiles(), file =>  // Changed to Parallel.ForEach
        {
            if (file.Extension == ".zip")
            {
                result.Children[file.Name] = new FileOrFolderInfo
                {
                    Type = "compressed file",
                    Parent = info.FullName,
                    Owner = GetOwner(file),
                    CreatedDate = file.CreationTime.Date,
                    CreatedTime = file.CreationTime.TimeOfDay,
                    AccessedDate = file.CreationTime.Date,
                    AccessedTime = file.CreationTime.TimeOfDay,
                    ModifiedDate = file.LastWriteTime.Date,
                    ModifiedTime = file.LastWriteTime.TimeOfDay,
                    Size = GetReadableSize(file.Length),
                    Children = GetCompressedFileContents(file)
                };
            }
            else
            {
                result.Children[file.Name] = new FileOrFolderInfo
                {
                    Type = "file",
                    Parent = info.FullName,
                    Owner = GetOwner(file),
                    CreatedDate = file.CreationTime.Date,
                    CreatedTime = file.CreationTime.TimeOfDay,
                    AccessedDate = file.CreationTime.Date,
                    AccessedTime = file.CreationTime.TimeOfDay,
                    ModifiedDate = file.LastWriteTime.Date,
                    ModifiedTime = file.LastWriteTime.TimeOfDay,
                    Size = (file.Length / 1024f / 1024f).ToString() + " MB"
                };
            }
        });

        Parallel.ForEach(info.GetDirectories(), directory =>  // Changed to Parallel.ForEach
        {
            result.Children[directory.Name] = ScanDirectory(directory.FullName);
        });

        return result;
    }



    private static ConcurrentDictionary<string, FileOrFolderInfo> GetCompressedFileContents(FileInfo compressedFile)  // Changed return type to ConcurrentDictionary
    {
        var contents = new ConcurrentDictionary<string, FileOrFolderInfo>();  // Changed to ConcurrentDictionary

        using (var archive = ZipFile.OpenRead(compressedFile.FullName))
        {
            foreach (var entry in archive.Entries)
            {
                contents[entry.Name] = new FileOrFolderInfo
                {
                    Type = "file in compressed file",
                    Size = GetReadableSize(entry.Length)
                };
            }
        }

        return contents;
    }



    private static long GetDirectorySize(DirectoryInfo directoryInfo)
    {
        long size = 0;
        FileInfo[] files = directoryInfo.GetFiles();
        foreach (FileInfo file in files)
        {
            size += file.Length;
        }
        DirectoryInfo[] directories = directoryInfo.GetDirectories();
        foreach (DirectoryInfo directory in directories)
        {
            size += GetDirectorySize(directory);
        }
        return size;
    }
    public static string GetOwner(FileInfo file)
    {
        try
        {
            if (file.Extension == ".doc" || file.Extension == ".docx" || file.Extension == ".docm")
            {
                var doc = new Aspose.Words.Document(file.FullName);
                return doc.BuiltInDocumentProperties.Author;
            }
            else if (file.Extension == ".xls" || file.Extension == ".xlsx" || file.Extension == ".xlsm")
            {
                var workbook = new Aspose.Cells.Workbook(file.FullName);
                return workbook.BuiltInDocumentProperties.Author;
            }
            else if (file.Extension == ".ppt" || file.Extension == ".pptx" || file.Extension == ".pptm")
            {
                var presentation = new Aspose.Slides.Presentation(file.FullName);
                return presentation.DocumentProperties.Author;
            }
            else if (file.Extension == ".pbix")
            {
                // No direct way to extract the author of a .pbix file.
                return "Unknown";
            }
            else
            {
                return file.GetAccessControl().GetOwner(typeof(System.Security.Principal.NTAccount)).ToString();
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
    public static void Main()
    {
        var result = new Dictionary<string, FileOrFolderInfo>
        {
            ["root"] = DirectoryScanner.ScanDirectory(@"C:\\Users\\dimit\\Downloads")
        };
        var json = JsonConvert.SerializeObject(result, Formatting.Indented);
        File.WriteAllText(@"C:\\Users\\dimit\\Documents\\output.json", json);


    }
}
