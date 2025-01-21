using App.Console;
using App.Excel;

Console.WriteLine("Generowanie pliku na miesiąc "+MonthDictionary.GetMonthName(DateTime.Now.AddMonths(1).Month)+"...");

try
{
    FileGenerator.GenerateFile();
    Console.WriteLine("Wygenerowano plik pomyślnie.. \n");
}
catch (InvalidOperationException e)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Plik w użyciu. Zamknij plik i spróbuj ponownie.\n");
}
catch (Exception e)
{
    Console.WriteLine(e.Message+"\n");
    throw;
}

Console.ResetColor();
Console.WriteLine("Naciśnij dowolny klawisz aby zakończyć...");
Console.ReadKey();
