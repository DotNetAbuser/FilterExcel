namespace FilterExcel.Application.IServices;

public interface ICountsService
{
    Task<string> CountsAsync(IFormFile xlsxInputFile);
}