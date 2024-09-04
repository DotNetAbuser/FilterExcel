namespace FilterExcel.Application.IServices;

public interface ISearchBySupplierService
{
    Task<string> SearchBySupplierAsync(IFormFile xlsxSupplierList, IFormFile xlsxInputFile);
}