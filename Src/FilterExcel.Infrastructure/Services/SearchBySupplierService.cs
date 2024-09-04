namespace FilterExcel.Infrastructure.Services;

public class SearchBySupplierService : ISearchBySupplierService
{
    private readonly IFileHelper _fileHelper;
    private readonly IExcelOperationsHelper _excelOperationsHelper;

    public SearchBySupplierService(
        IFileHelper fileHelper,
        IExcelOperationsHelper excelOperationsHelper)
    {
        _fileHelper = fileHelper;
        _excelOperationsHelper = excelOperationsHelper;
    }
    
    public async Task<string> SearchBySupplierAsync(
        IFormFile xlsxSupplierList, IFormFile xlsxInputFile)
    {
        var directoryUniqueName = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
        
        var xlsxSupplierListPath = await _fileHelper.CreateFileByFormAsync(
            FileSystemConstants.SearchBySupplierList, directoryUniqueName,
            xlsxSupplierList);
        var xlsxInputPath = await _fileHelper.CreateFileByFormAsync(
            FileSystemConstants.SearchBySupplierList, directoryUniqueName,
            xlsxInputFile);

        using var suppliersWorkbook = new XLWorkbook(xlsxSupplierListPath);
        using var inputWorkbook = new XLWorkbook(xlsxInputPath);
        
        var suppliersWorksheet = suppliersWorkbook.Worksheets.FirstOrDefault();
        if (suppliersWorksheet == null) return "Ошибка в xlsx файле с поставщиками нет worksheets!";
        
        var inputWorksheet = inputWorkbook.Worksheets.FirstOrDefault();
        if (inputWorksheet == null) return "Ошибка в xlsx файле с полезной нагрузкой нет worksheets!";

        var remainingWorkbook = new XLWorkbook();
        var remainingWorksheet = remainingWorkbook.AddWorksheet("Остаток");

        var organizationWorkbooks = new Dictionary<string, XLWorkbook>();
        var organizations = ReadSuppliersData(suppliersWorksheet);

        var copies = new Dictionary<string, List<IXLRow>>();
        foreach (var organization in organizations)
        {
            var orgWorkbook = new XLWorkbook();
            var orgWorksheet = orgWorkbook.AddWorksheet("Data");
            _excelOperationsHelper.CopyRow(inputWorksheet.Row(1), orgWorksheet.Row(1));
            organizationWorkbooks[organization.Title] = orgWorkbook;
            copies[organization.Title] = inputWorksheet.RowsUsed().Skip(1).ToList();
        }
        copies["remains"] = inputWorksheet.RowsUsed().Skip(1).ToList();

        _excelOperationsHelper.CopyRow(inputWorksheet.Row(1), remainingWorksheet.Row(1));

        var rowsToDelete = new HashSet<int>();

        // Process each row for each organization
        foreach (var org in organizations)
        {
            var orgRows = copies[org.Title];
            for (var i = 0; i < orgRows.Count; i++)
            {
                var productRow = orgRows[i];
                if (!productRow.Cell(5).GetString().Equals(org.Title, StringComparison.Ordinal) &&
                    !org.SupplierList.Any(v => productRow.Cell(5).GetString().Equals(v, StringComparison.Ordinal))) 
                    continue;
                
                var orgWorksheet = organizationWorkbooks[org.Title].Worksheet("Data");
                var newRowNumber = orgWorksheet.LastRowUsed()?.RowNumber() + 1 ?? 2;
                _excelOperationsHelper.CopyRow(productRow, orgWorksheet.Row(newRowNumber));
                rowsToDelete.Add(productRow.RowNumber());

                while (i + 1 < orgRows.Count && _excelOperationsHelper.AllOtherCellsEmpty(
                           orgRows[i + 1], 2))
                {
                    i++;
                    productRow = orgRows[i];
                        
                    if (productRow.Cell(2).Style.Font.Bold) continue;
                        
                    _excelOperationsHelper.CopyRow(productRow, orgWorksheet.Row(++newRowNumber));
                    rowsToDelete.Add(productRow.RowNumber());
                }
            }
        }

        var remainingRows = copies["remains"];
        for (var i = 0; i < remainingRows.Count; i++)
        {
            var productRow = remainingRows[i];
            if (rowsToDelete.Contains(productRow.RowNumber())) continue;
            
            var newRowNumber = remainingWorksheet.LastRowUsed()?.RowNumber() + 1 ?? 2;
            _excelOperationsHelper.CopyRow(productRow, remainingWorksheet.Row(newRowNumber));

            while (i + 1 < remainingRows.Count && _excelOperationsHelper.AllOtherCellsEmpty(
                       remainingRows[i + 1], 2))
            {
                i++;
                productRow = remainingRows[i];
                if (!productRow.Cell(2).Style.Font.Bold) // Check for bold font
                {
                    _excelOperationsHelper.CopyRow(productRow, remainingWorksheet.Row(++newRowNumber));
                }
            }
        }

        foreach (var rowNumber in rowsToDelete.OrderByDescending(r => r))
            inputWorksheet.Row(rowNumber).Delete();

        foreach (var kvp in organizationWorkbooks)
        {
            var fileName = _excelOperationsHelper.SaveWorkbook(kvp.Value, xlsxSupplierListPath, kvp.Key);
            _excelOperationsHelper.InsertTableWithBorders(kvp.Value.Worksheet("Data"));
            _excelOperationsHelper.EditWorkbookForBold(kvp.Value, fileName);
        }

        var remainingFileName = _excelOperationsHelper.SaveWorkbook(remainingWorkbook, xlsxSupplierListPath, "остаток");
        _excelOperationsHelper.InsertTableWithBorders(remainingWorkbook.Worksheet("Остаток"));
        _excelOperationsHelper.EditWorkbookForBold(remainingWorkbook, remainingFileName);
        
        return "Операция завершена успешно!";
    }

    private List<OrganizationModel> ReadSuppliersData(IXLWorksheet worksheet)
    {
        var organizations = new List<OrganizationModel>();
        
        var lastColumn = worksheet.LastColumnUsed().ColumnNumber();
        for (var col = 2; col <= lastColumn; col++)
        {
            var organization = worksheet.Cell(1, col).GetString();
            
            if (string.IsNullOrEmpty(organization)) continue;
            
            var supplierList = new List<string>();
            for (var row = 8; row <= worksheet.LastRowUsed().RowNumber(); row++)
            {
                var supplier = worksheet.Cell(row, col).GetString();
                if (!string.IsNullOrEmpty(supplier))
                {
                    supplierList.Add(supplier);
                }
            }
            organizations.Add(new OrganizationModel(organization, supplierList));
        }
        
        return organizations;
    }
}