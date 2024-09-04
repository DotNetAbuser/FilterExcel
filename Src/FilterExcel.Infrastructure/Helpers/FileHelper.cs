namespace FilterExcel.Infrastructure.Helpers;

public class FileHelper : IFileHelper
{
    public async Task<string> CreateFileByFormAsync(
        string routeFilePath, string saveDirectoryPath, 
        IFormFile file)
    {
        var rootFilePath = routeFilePath + "/" + saveDirectoryPath;
        
        if (!Directory.Exists(rootFilePath)) Directory.CreateDirectory(rootFilePath);
        
        var filePath = rootFilePath + $"/{file.FileName}";
        await using var stream = new FileStream(filePath, FileMode.Create);
        await file.CopyToAsync(stream);
        return filePath;
    }

    public async Task DeleteFileAsync(string filePath)
    {
        if (!File.Exists(filePath)) return;
        File.Delete(filePath);
    }
}