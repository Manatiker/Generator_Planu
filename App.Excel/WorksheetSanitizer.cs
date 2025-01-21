using OfficeOpenXml;

namespace App.Excel;

public class WorksheetSanitizer
{
    public void HideWorksheet(ExcelPackage package, string sheetName)
    {
        var worksheet = package.Workbook.Worksheets[sheetName];
        worksheet.Hidden = eWorkSheetHidden.Hidden;
    }
    
    public void ChangeColumnWidth(ExcelWorksheet worksheet, int columnNumber, int width)
    {
        worksheet.Column(columnNumber).Width = width;
    }
}