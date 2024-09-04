namespace FilterExcel.Infrastructure.Services;

public class FilterService : IFilterService
{
    private readonly IFileHelper _fileHelper;
    
    public FilterService(
        IFileHelper fileHelper)
    {
        _fileHelper = fileHelper;
    }
    
    public async Task<string> FilterByUniqueAsync(
        int filterColumnIndex, IFormFile xlsxFile)
    {
        var directoryUniqueName = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
        
        var xlsxFilePath = await _fileHelper.CreateFileByFormAsync(
            FileSystemConstants.RemovesDublicates, directoryUniqueName,
            xlsxFile);

        using var workbook = new XLWorkbook(xlsxFilePath);
        var worksheet = workbook.Worksheets.FirstOrDefault();
        if (worksheet == null) return "Загруженный файл xlsx не содержит worksheet!";

        var rowCount = worksheet.Rows().Count();
        var list = new Dictionary<string, int>();
            
        for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
        {
            var value = worksheet.Cell(rowIndex, filterColumnIndex).Value.ToString();
            if (list.ContainsKey(value)) list[value]++;
            else list[value] = 1;
        }
        for (var rowIndex = rowCount; rowIndex >= 1; rowIndex--)
        {
            var value = worksheet.Cell(rowIndex, filterColumnIndex).Value.ToString();
            if (list[value] > 1) worksheet.Row(rowIndex).Delete();
        }
            
        workbook.SaveAs(xlsxFilePath);

        return "Из файла успешно удаленны значения дубликаты!";
    }
}