namespace App.Generator_Planu;

public static class MonthDictionary
{
    private static readonly Dictionary<int, string> _monthNames = new Dictionary<int, string>
    {
        { 1, "Styczeń" },
        { 2, "Luty" },
        { 3, "Marzec" },
        { 4, "Kwiecień" },
        { 5, "Maj" },
        { 6, "Czerwiec" },
        { 7, "Lipiec" },
        { 8, "Sierpień" },
        { 9, "Wrzesień" },
        { 10, "Październik" },
        { 11, "Listopad" },
        { 12, "Grudzień" }
    };
    
    public static string GetMonthName(int month)
    {
        if (month == 13) month = 1;
        
        if (month < 1 || month > 13)
        {
            throw new ArgumentOutOfRangeException(nameof(month), "Month must be between 1 and 12.");
        }

        return _monthNames[month];
    }
    
    public static string GetMonthName(DateTime date)
    {
        return GetMonthName(date.Month);
    }
}