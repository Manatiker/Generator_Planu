using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace App.Excel;

public class WorksheetPopulator
{
    public void PopulateWorksheet(ExcelWorksheet worksheet)
    {
        
        worksheet.Cells["A1"].Value = "Lp.";
        worksheet.Cells["B1"].Value = "Dostawa/Załadunek";
        worksheet.Cells["C1"].Value = "Status";
        worksheet.Cells["D1"].Value = "Klient";
        worksheet.Cells["E1"].Value = "Magazyn";
        worksheet.Cells["F1"].Value = "Godzina Zał.";
        worksheet.Cells["G1"].Value = "Godzina Dost.";
        worksheet.Cells["H1"].Value = "Pojazd";
        worksheet.Cells["I1"].Value = "Nr Kontenera";
        worksheet.Cells["J1"].Value = "Kierowca";
        worksheet.Cells["K1"].Value = "Tel. Kierowcy";
        worksheet.Cells["L1"].Value = "Nr Dokumentu";
        worksheet.Cells["M1"].Value = "Towar";
        worksheet.Cells["N1"].Value = "Uwagi";
        worksheet.Cells["N1"].Style.Font.Color.SetColor(Color.Red);

        for (int i = 2; i <= 201; i++)
        {
            worksheet.Cells["A" + i.ToString()].Value = i - 1;
        }
        
        worksheet.Cells.AutoFitColumns();

        AddStylesToHeader(worksheet);
        AddStylesToWholeWorksheet(worksheet);
        AddDropdownToWorksheet(worksheet, "Odniesienia");
        AddConditionalFormatting(worksheet);
        
        var sanitaizer = new WorksheetSanitizer();
        sanitaizer.ChangeColumnWidth(worksheet, 1, 5);
    }
    
    public void PopulateReferenceSheet(ExcelWorksheet worksheet)
    {
        worksheet.Cells["A1"].Value = "Dostawa";
        worksheet.Cells["A2"].Value = "Załadunek";
        worksheet.Cells["A3"].Value = "Przerzuty";
    }

    private void AddConditionalFormatting(ExcelWorksheet worksheet)
    {
        var cfLoading = worksheet.ConditionalFormatting.AddEqual(worksheet.Cells["B2:B200"]);
        cfLoading.Formula = "\"Załadunek\"";
        cfLoading.Style.Fill.BackgroundColor.SetColor(0,218,233,248);

        var cfUnloading = worksheet.ConditionalFormatting.AddEqual(worksheet.Cells["B2:B200"]);
        cfUnloading.Formula = "\"Dostawa\"";
        cfUnloading.Style.Fill.BackgroundColor.SetColor(0,218,242,208);
    }
    

    private void AddStylesToHeader(ExcelWorksheet worksheet)
    {
        using var range = worksheet.Cells[1, 1, 1, 14];
        range.Style.Font.Bold = true;
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(0,255, 255, 0);
        
        worksheet.Rows[1].Height = 45;
        worksheet.Columns[1].Width = 5; //Lp.
        worksheet.Columns[8].Width = 20; //Pojazd
        worksheet.Columns[10].Width = 20; //Kierowca
        worksheet.Columns[13].Width = 40; //Towar
    }

    private void AddStylesToWholeWorksheet(ExcelWorksheet worksheet)
    {
        using var range = worksheet.Cells[1, 1, 200, 14];
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }
    
    private void AddDropdownToWorksheet(ExcelWorksheet worksheet, string referenceSheetName)
    {
        var validation = worksheet.DataValidations.AddListValidation("B2:B200");
        validation.Formula.ExcelFormula = $"{referenceSheetName}!$A$1:$A$3";
    }

    
    
}