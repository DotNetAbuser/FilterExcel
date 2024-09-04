using ClosedXML.Excel;

namespace FilterExcel.Application.IHelpers;

public interface IExcelOperationsHelper
{
    void InsertTableWithBorders(IXLWorksheet worksheet);
    void EditWorkbookForBold(XLWorkbook workbook, string fileName);
    bool AllOtherCellsEmpty(IXLRow row, int columnToExclude);
    void CopyRow(IXLRow srcRow, IXLRow dstRow);
    // void CopyRow(IXLRow srcRow, IXLWorksheet dstWorksheet, int dstRowNumber);
    string SaveWorkbook(XLWorkbook workbook, string originalFileName, string suffix);
}