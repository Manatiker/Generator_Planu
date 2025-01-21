using App.Generator_Planu;
using OfficeOpenXml;

namespace App.Excel;

public class FileGenerator
{
    public static void GenerateFile(int monthNumber)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        using var package = new ExcelPackage();
        
        var days = GetWorkingDaysForAMonth(monthNumber);
        
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

        package.SaveAs(new FileInfo(@".\Plany\Plan_"+MonthDictionary.GetMonthName(monthNumber)+"_2025.xlsx"));
    }
    
    private static List<DateTime> GetWorkingDaysForAMonth(int monthNumber)
    {
        var days = new List<DateTime>();
        var daysInMonth = DateTime.DaysInMonth(2025, monthNumber);
        
        for (int i = 1; i <= daysInMonth; i++)
        {
            var day = new DateTime(2025, monthNumber, i);
            if (day.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) continue;
            
            days.Add(day);
        }

        return days;
    }
}
