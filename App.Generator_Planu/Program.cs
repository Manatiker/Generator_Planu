using App.Generator_Planu;
using App.Excel;

Console.WriteLine("Podaj numer miesiąca dla którego chcesz wygenerować plik (1-styczen, 2-luty, itd.): ");
var month = Console.ReadLine();

while (string.IsNullOrEmpty(month))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Podaj numer miesiąca z przedziału od 1 do 12.");
    month = Console.ReadLine();
}

while ((int.TryParse(month, out var x) == false) || (int.Parse(month) is < 1 or > 12))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Podaj poprawny numer miesiąca z przedziału od 1 do 12.");
    Console.ResetColor();
    month = Console.ReadLine();
}

var monthNo = int.Parse(month);

Console.WriteLine("Generowanie pliku na miesiąc "+MonthDictionary.GetMonthName(monthNo)+"...");

try
{
    FileGenerator.GenerateFile(monthNo);
    Console.WriteLine("Wygenerowano plik pomyślnie.. \n");
}
catch (InvalidOperationException)
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
