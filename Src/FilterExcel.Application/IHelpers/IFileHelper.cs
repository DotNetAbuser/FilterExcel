namespace FilterExcel.Application.IHelpers;

public interface IFileHelper
{
    Task<string> CreateFileByFormAsync(string routeFilePath, string saveDirectoryPath, IFormFile file);
    Task DeleteFileAsync(string filePath);
}