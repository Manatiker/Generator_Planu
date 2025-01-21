using App.Console;
using OfficeOpenXml;

namespace App.Excel;

public class FileGenerator
{
    public static void GenerateFile()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        using var package = new ExcelPackage();

        var date = DateTime.Now.AddMonths(1);
        
        var days = GetWorkingDaysForAMonth(date);
        
        const string REF_SHEET_NAME = "Odniesienia";

        var refWorksheet = package.Workbook.Worksheets.Add(REF_SHEET_NAME);
        refWorksheet.Hidden = eWorkSheetHidden.Hidden;
        
        var populator = new WorksheetPopulator();
        populator.PopulateReferenceSheet(refWorksheet);

        foreach (var day in days)
        {
            var worksheet = package.Workbook.Worksheets.Add(day.ToString("dd.MM"));
            populator = new WorksheetPopulator();
            populator.PopulateWorksheet(worksheet);
        }
        
        var sanitizer = new WorksheetSanitizer();
        sanitizer.HideWorksheet(package, REF_SHEET_NAME);

        package.SaveAs(new FileInfo("C:\\CodePlayground\\Excel_Ania\\Plan_"+MonthDictionary.GetMonthName(date.Month)+"_"+date.Year+".xlsx"));
    }
    
    private static List<DateTime> GetWorkingDaysForAMonth(DateTime date)
    {
        var days = new List<DateTime>();
        var daysInMonth = DateTime.DaysInMonth(date.Year, date.Month);
        for (int i = 1; i <= daysInMonth; i++)
        {
            var day = new DateTime(date.Year, date.Month, i);
            if (day.DayOfWeek != DayOfWeek.Saturday && day.DayOfWeek != DayOfWeek.Sunday)
            {
                days.Add(day);
            }
        }

        return days;
    }
}
