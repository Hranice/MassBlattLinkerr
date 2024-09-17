using Serilog;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace MassBlattLinkerr.Core
{
    /// <summary>
    /// Class for handling Excel file operations.
    /// </summary>
    public class ExcelHandler
    {
        private readonly ILogger? _logger;

        /// <summary>
        /// Initializes the ExcelHandler with the application's logger.
        /// </summary>
        /// <param name="logger">The logger instance from the main application.</param>
        public ExcelHandler(ILogger? logger = null)
        {
            if (logger is not null)
            {
                _logger = logger;
            }
        }

        /// <summary>
        /// Processes the Excel file by adding a VBA macro and saves a copy.
        /// </summary>
        /// <param name="sourceFilePath">The path to the source Excel file.</param>
        /// <param name="destinationFilePath">The destination path for the processed Excel file.</param>
        /// <param name="executablePath">The path to the executable called by the macro.</param>
        /// <param name="macroCode">Optional custom VBA macro code. If not provided, a default macro will be used.</param>
        public void ProcessExcelFile(
            string sourceFilePath,
            string destinationFilePath,
            string? executablePath = null,
            string? macroCode = null)
        {
            if (!File.Exists(sourceFilePath))
            {
                _logger?.Error("The specified source file does not exist: {SourceFilePath}", sourceFilePath);
                return;
            }

            // Create destination directory if it doesn't exist
            string destinationDirectory = Path.GetDirectoryName(destinationFilePath);
            if (!Directory.Exists(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }

            var excelApp = new Excel.Application { DisplayAlerts = false };

            try
            {
                _logger?.Information("Opening Excel file: {SourceFilePath}", sourceFilePath);
                var workbook = excelApp.Workbooks.Open(sourceFilePath, ReadOnly: true);

                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    var vbComponent = workbook.VBProject.VBComponents.Item(sheet.CodeName);

                    // Use provided macro code or fall back to the default
                    string macro = macroCode ?? GetDefaultVbaMacro(executablePath);
                    vbComponent.CodeModule.AddFromString(macro);
                }

                workbook.SaveAs(destinationFilePath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                _logger?.Information("Macro successfully added. File saved at: {DestinationFilePath}, Macro: '{Macro}'", destinationFilePath, GetDefaultVbaMacro(executablePath));

                // Optionally, launch the file after saving
                LaunchFile(destinationFilePath);
            }
            catch (Exception ex)
            {
                _logger?.Error("Error processing Excel file: {Message}\nStack Trace: {StackTrace}", ex.Message, ex.StackTrace);
            }
            finally
            {
                excelApp.Quit();
            }
        }

        /// <summary>
        /// Returns the default VBA macro code.
        /// </summary>
        /// <param name="executablePath">The path to the executable that the macro will call.</param>
        private string GetDefaultVbaMacro(string executablePath)
        {
            return $@"Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim cellValue As String
    Dim shellCommand As String

    Cancel = True
    cellValue = Target.Value
    shellCommand = ""{executablePath} """""" & cellValue & """""""" 
    Call Shell(shellCommand, vbNormalFocus)
End Sub";
        }

        /// <summary>
        /// Optionally launches the Excel file after saving.
        /// </summary>
        /// <param name="filePath">The file path to launch.</param>
        private void LaunchFile(string filePath)
        {
            try
            {
                _logger?.Information("Launching file: {FilePath}", filePath);
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                _logger?.Error("Failed to launch file: {Message}\nStack Trace: {StackTrace}", ex.Message, ex.StackTrace);
            }
        }
    }
}
