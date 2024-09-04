namespace FilterExcel.Infrastructure.Helpers;

public class ExcelOperationsHelper : IExcelOperationsHelper
{
    public void InsertTableWithBorders(IXLWorksheet worksheet)
    {
        var firstRow = worksheet.FirstRowUsed().RowNumber();
        var lastRow = worksheet.LastRowUsed().RowNumber();
        var firstColumn = worksheet.FirstColumnUsed().ColumnNumber();
        var lastColumn = worksheet.LastColumnUsed().ColumnNumber();

        var range = worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        range.CreateTable();
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
    }
        
    public void EditWorkbookForBold(XLWorkbook workbook, string fileName)
    {
        var worksheet = workbook.Worksheets.First();
        for (var i = worksheet.LastRowUsed().RowNumber(); i > 2; i--)
        {
            var cell = worksheet.Row(i).Cell(2);
            if (cell.Style.Font.Bold) worksheet.Row(i).InsertRowsAbove(2);
        }

        workbook.SaveAs(fileName);
    }
        
    public bool AllOtherCellsEmpty(IXLRow row, int columnToExclude)
    {
        return row.CellsUsed(c => c.Address.ColumnNumber != columnToExclude)
            .All(c => string.IsNullOrWhiteSpace(c.GetString()));
    }
        
    public void CopyRow(IXLRow srcRow, IXLRow dstRow)
    {
        foreach (var cell in srcRow.CellsUsed())
        {
            var dstCell = dstRow.Cell(cell.Address.ColumnNumber);
            dstCell.Value = cell.Value;
            dstCell.Style = cell.Style;
            dstCell.WorksheetColumn().Width = cell.WorksheetColumn().Width;
            dstCell.WorksheetRow().Height = cell.WorksheetRow().Height;
        }
    }
        
    public string SaveWorkbook(XLWorkbook workbook, string originalFileName, string suffix) 
    {
        var directory = Path.GetDirectoryName(originalFileName); 
        var baseFileName = Path.GetFileNameWithoutExtension(originalFileName);
        var extension = Path.GetExtension(originalFileName);
        var dstWorkbookFileName = Path.Combine(directory, $"{baseFileName}_{suffix}{extension}");

        var counter = 1;
        while (File.Exists(dstWorkbookFileName))
        {
            dstWorkbookFileName = Path.Combine(directory, $"{baseFileName}_{suffix}_{counter++}{extension}");
        }

        workbook.SaveAs(dstWorkbookFileName);
        return dstWorkbookFileName;
    }
}