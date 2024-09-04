namespace FilterExcel.Api.Controllers;

public class ExcelManipulationController : BaseController
{
    private readonly IFilterService _filterService;
    private readonly ISearchValuesFromTxtService _searchValuesFromTxtService;
    private readonly ISearchBySupplierService _searchBySupplierService;
    private readonly ICountsService _countsService;
        
    public ExcelManipulationController(
        IFilterService filterService,
        ISearchValuesFromTxtService searchValuesFromTxtService,
        ISearchBySupplierService searchBySupplierService,
        ICountsService countsService)
    {
        _filterService = filterService;
        _searchValuesFromTxtService = searchValuesFromTxtService;
        _searchBySupplierService = searchBySupplierService;
        _countsService = countsService;
    }
    
    [HttpPost("delete-dublicates")]
    public async Task<IActionResult> FilterFileAndSaveOnServerAsync(
        int filterColumnIndex, IFormFile xlsxInputFile)
    {
        return Ok(await _filterService.FilterByUniqueAsync(filterColumnIndex, xlsxInputFile));
    }

    [HttpPost("search-by-values-from-txt")]
    public async Task<IActionResult> SearchValuesFromTxtAsync(
        IFormFile txtFileWithSearchValues, IFormFile xlsxInputFile)
    {
        return Ok(await _searchValuesFromTxtService.SearchFromTxtAsync(xlsxInputFile, txtFileWithSearchValues));
    }
    
    [HttpPost("search-by-suppliers-xlsx")]
    public async Task<IActionResult> FindBySuppliersAsync(
        IFormFile xlsxSupplierList, IFormFile xlsxInputFile)
    {
        return Ok(await _searchBySupplierService.SearchBySupplierAsync(xlsxSupplierList, xlsxInputFile));
    }

    [HttpPost("counts")]
    public async Task<IActionResult> CountsAsync(
        IFormFile xlsxInputFile)
    {
        return Ok(await _countsService.CountsAsync(xlsxInputFile));
    }

    
}