namespace FilterExcel.Infrastructure.Services;

public class SearchValuesFromTxtService : ISearchValuesFromTxtService
{
    private readonly IFileHelper _fileHelper;
    private readonly IExcelOperationsHelper _excelOperationsHelper;

    public SearchValuesFromTxtService(
        IFileHelper fileHelper,
        IExcelOperationsHelper excelOperationsHelper)
    {
        _fileHelper = fileHelper;
        _excelOperationsHelper = excelOperationsHelper;
    }

    public async Task<string> SearchFromTxtAsync(
        IFormFile xlsxFile, IFormFile txtFileWithValues)
    {
        var directoryUniqueName = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
        
        var xlsxFilePath = await _fileHelper.CreateFileByFormAsync(
            FileSystemConstants.SearchByValuesFromTxt, directoryUniqueName, 
            xlsxFile);
        var txtFilePath = await _fileHelper.CreateFileByFormAsync(
            FileSystemConstants.SearchByValuesFromTxt, directoryUniqueName, 
            txtFileWithValues);

        var searchWordsList = (await File.ReadAllLinesAsync(txtFilePath))
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();

        if (searchWordsList.Count == 0) return "Файл со словами фильтрами пуст!";

        using var inputWorkBook = new XLWorkbook(xlsxFilePath);
        using var resultWorkBookFiltered = new XLWorkbook();
        using var resultWorkBookRemaining = new XLWorkbook();
        
        var inputWorkSheet = inputWorkBook.Worksheets.FirstOrDefault();
        if (inputWorkSheet == null) return "В загруженном xlsx файл отсутвует worksheet!";
        
        var resultWorksheetFiltered = resultWorkBookFiltered.AddWorksheet("С фильтром");
        var resultWorksheetRemaining = resultWorkBookRemaining.AddWorksheet("Остаток");
        
        _excelOperationsHelper.CopyRow(inputWorkSheet.Row(1), resultWorksheetFiltered.Row(1));
        _excelOperationsHelper.CopyRow(inputWorkSheet.Row(1), resultWorksheetRemaining.Row(1));

        var filteredRowCount = 2;
        var remainingRowCount = 2;

        var deleteRowIndexList = new List<int>();

        foreach (var row in inputWorkSheet.RowsUsed().Skip(1))
        {
            var isFiltered = false;
            var cellValue = row.Cell(2).GetString();

            if (string.IsNullOrWhiteSpace(cellValue) ||
                cellValue.StartsWith($"-")) continue;

            foreach (var value in searchWordsList)
            {
                var word = value.Trim();
                var wordPattern = $@"\b{Regex.Escape(word)}\b";
                var wordWithHyphenPatternBefore = $@"-{Regex.Escape(word)}\b";
                var wordWithHyphenPatternAfter = $@"\b{Regex.Escape(word)}-";

                if (Regex.IsMatch(cellValue, wordPattern, RegexOptions.IgnoreCase) &&
                    !Regex.IsMatch(cellValue, wordWithHyphenPatternBefore, RegexOptions.IgnoreCase) &&
                    !Regex.IsMatch(cellValue, wordWithHyphenPatternAfter, RegexOptions.IgnoreCase))
                {
                    isFiltered = true;
                    break;
                }
            }

            if (!isFiltered) continue;

            if (row.Cell(2).Style.Font.Bold &&
                _excelOperationsHelper.AllOtherCellsEmpty(row, 2)) continue;

            _excelOperationsHelper.CopyRow(row, resultWorksheetFiltered.Row(filteredRowCount++));
            deleteRowIndexList.Add(row.RowNumber());
            
            var currentRow = row.RowNumber();
            while (currentRow < inputWorkSheet.LastRowUsed().RowNumber())
            {
                var nextRow = inputWorkSheet.Row(++currentRow);
                if (string.IsNullOrWhiteSpace(nextRow.Cell(2).GetString()) ||
                    !_excelOperationsHelper.AllOtherCellsEmpty(nextRow, 2)) break;

                _excelOperationsHelper.CopyRow(nextRow, resultWorksheetFiltered.Row(filteredRowCount++));
                deleteRowIndexList.Add(nextRow.RowNumber());
            }
        }

        foreach (var rowIndex in deleteRowIndexList.OrderByDescending(r => r))
            inputWorkSheet.Row(rowIndex).Delete();


        foreach (var row in inputWorkSheet.RowsUsed().Skip(1))
            _excelOperationsHelper.CopyRow(row, resultWorksheetRemaining.Row(remainingRowCount++));

        _excelOperationsHelper.InsertTableWithBorders(resultWorksheetFiltered);
        _excelOperationsHelper.InsertTableWithBorders(resultWorksheetRemaining);

        var filteredFileName = _excelOperationsHelper.SaveWorkbook(
            resultWorkBookFiltered, xlsxFilePath, "Результат");
        var remainingFileName = _excelOperationsHelper.SaveWorkbook(
            resultWorkBookRemaining, xlsxFilePath, "Результат_остаток");

        _excelOperationsHelper.EditWorkbookForBold(resultWorkBookRemaining, remainingFileName);

        return "Операция завершенна успешно!";
    }
}