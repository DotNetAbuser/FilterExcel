namespace FilterExcel.Application.IServices;

public interface ISearchValuesFromTxtService
{
    Task<string> SearchFromTxtAsync(IFormFile xlsxFile, IFormFile txtFileWithValues);
}