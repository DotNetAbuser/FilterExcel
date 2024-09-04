namespace FilterExcel.Application.IServices;

public interface IFilterService
{
    Task<string> FilterByUniqueAsync(int filterColumn, IFormFile xlsxFile);
}