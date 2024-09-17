using MassBlattLinkerr.Core;
using Serilog;
using System.Diagnostics;

class Program
{
    public static void Main(string[] args)
    {
        ILogger log = InitializeLogger();

        log.Information("Zahajuji operaci...");

        if (args.Length == 0)
        {
            new ExcelHandler(log).
                ProcessExcelFile(
                @"Z:\Plan\Plan vyroby\Plánování\2024\plánování_2024_nové.xlsm",
                Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsm",
                Environment.ProcessPath
                );
        }
        else
        {
            var dbHandler = new DatabaseHandler();

            if (args[0] == "!")
            {
                dbHandler.InitializeAndPopulateDatabase();
            }

            else
            {
                string cellValue = args[0];
                log.Information($"Hodnota buňky: {cellValue}");

                var articleHandler = new ArticleHandler();
                var articleWithPrintVersion = articleHandler.ExtractArticleAndVersion(cellValue);
                log.Information("Naformátovaná hodnota: {article}.{printversion}", articleWithPrintVersion.Item1, articleWithPrintVersion.Item2);
                var matchedPrintVersionPaths = dbHandler.SearchByArticleAndVersion(articleWithPrintVersion.Item1, articleWithPrintVersion.Item2);

                foreach (var matchedPrintVersionPath in matchedPrintVersionPaths)
                {
                    log.Information("Otevírám: {matchedPrintVersionPath}", matchedPrintVersionPath);
                    try
                    {
                        Process.Start(new ProcessStartInfo(matchedPrintVersionPath) { UseShellExecute = true });
                    }

                    catch
                    {
                        string? matchedPrintVersionPathDirectory = Path.GetDirectoryName(matchedPrintVersionPath);
                        log.Warning("Nepodařilo se otevřít soubor. Otevírám složku {matchedPrintVersionPathDirectory}", matchedPrintVersionPathDirectory);
                        Process.Start(new ProcessStartInfo(matchedPrintVersionPathDirectory ?? throw new ArgumentNullException(nameof(matchedPrintVersionPathDirectory)))
                        {
                            UseShellExecute = true
                        });

                    }
                }
            }
        }
    }

    static ILogger InitializeLogger()
    {
        string exePath = AppDomain.CurrentDomain.BaseDirectory;
        string logFilePath = Path.Combine(exePath, "log.txt");

        return new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.File(logFilePath)
            .CreateLogger();
    }

}