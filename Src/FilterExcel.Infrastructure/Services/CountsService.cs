namespace FilterExcel.Infrastructure.Services;

public class CountsService : ICountsService
{
    private readonly IFileHelper _fileHelper;
    private readonly IExcelOperationsHelper _excelOperationsHelper;
    private readonly IRegexOperationsHelper _regexOperationsHelper;

    public CountsService(
        IFileHelper fileHelper,
        IExcelOperationsHelper excelOperationsHelper,
        IRegexOperationsHelper regexOperationsHelper)
    {
        _fileHelper = fileHelper;
        _excelOperationsHelper = excelOperationsHelper;
        _regexOperationsHelper = regexOperationsHelper;
    }
    
    public async Task<string> CountsAsync(IFormFile xlsxInputFile)
    {
        var directoryUniqueName = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
            
        var xlsxInputPath = await _fileHelper.CreateFileByFormAsync(
            FileSystemConstants.Counts, directoryUniqueName,
            xlsxInputFile);

        await InOrderFirst(xlsxInputPath);
        await BalanceOfOrder(xlsxInputPath);

        return "Операция завершена успешно!";
    }

    private async Task InOrderFirst(string xlsxInputPath)
    {
        using var inputWorkbook = new XLWorkbook(xlsxInputPath);
        using var resultWorkbook = new XLWorkbook();
        var inputWorksheet = inputWorkbook.Worksheets.FirstOrDefault();
        if (inputWorksheet == null) return;

        var resultWorksheet = resultWorkbook.AddWorksheet("Заказ");

        _excelOperationsHelper.CopyRow(inputWorksheet.Row(1), resultWorksheet.Row(1));

        var dstRowCount = 2;
        foreach (var row in inputWorksheet.RowsUsed().Skip(1))
        {
            var cellValue = row.Cell(7).FormulaA1;
            if (string.IsNullOrEmpty(cellValue)) cellValue = row.Cell(7).GetValue<string>();
            

            if (Regex.IsMatch(cellValue, @"=*\d+\s*[-+]\s*\d+") || 
                Regex.IsMatch(cellValue, @"\d+\s*[-+]\s*\d+"))
            {
                if (cellValue.StartsWith("(") &&
                    cellValue.EndsWith(")") &&
                    Regex.IsMatch(cellValue, @"\([^()]+\)")) continue;
                

                while (Regex.IsMatch(cellValue, @"\([^()]+\)"))
                {
                    cellValue = Regex.Replace(cellValue, @"\(([^()]+)\)", match =>
                    {
                        var expression = match.Groups[1].Value;
                        var result = _regexOperationsHelper.CalculateExpression(expression);
                        return result.ToString();
                    });
                }

                var sumOfSubtractedNumbers = _regexOperationsHelper.SumSubtractedNumbers(cellValue);
                row.Cell(7).Value = sumOfSubtractedNumbers;
                _excelOperationsHelper.CopyRow(row, resultWorksheet.Row(dstRowCount++));
            }
        }

        _excelOperationsHelper.InsertTableWithBorders(resultWorksheet);
        _excelOperationsHelper.SaveWorkbook(resultWorkbook, xlsxInputPath, "результат_в_заказ");
    }

    private async Task BalanceOfOrder(string xlsxInputPath)
    {
        using var inputWorkbook = new XLWorkbook(xlsxInputPath);
        using var resultWorkbook = new XLWorkbook();
        var inputWorksheet = inputWorkbook.Worksheets.FirstOrDefault();
        if (inputWorksheet == null) return;

        var resultWorksheet = resultWorkbook.AddWorksheet("Остаток на заказ");

        _excelOperationsHelper.CopyRow(inputWorksheet.Row(1), resultWorksheet.Row(1));

        var resultRowCount = 2;
        foreach (var row in inputWorksheet.RowsUsed().Skip(1))
        {
            var cell = row.Cell(7);
            var cellValue = cell.HasFormula ? cell.FormulaA1 : cell.GetValue<string>();

            if (Regex.IsMatch(cellValue, @"\d+\s*[-+]\s*\d+") && !cellValue.StartsWith("(") ||
                Regex.IsMatch(cellValue, @"\d+\s*[-+]\s*\d+") && !cellValue.EndsWith(")"))
            {
                cell.Value = $"({cellValue})";
            }
            else
            {
                cell.Value = cellValue;
            }

            _excelOperationsHelper.CopyRow(row, resultWorksheet.Row(resultRowCount++));
        }

        _excelOperationsHelper.InsertTableWithBorders(resultWorksheet);
        _excelOperationsHelper.SaveWorkbook(resultWorkbook, xlsxInputPath, "результат_остаток_от_заказа");
    }
}