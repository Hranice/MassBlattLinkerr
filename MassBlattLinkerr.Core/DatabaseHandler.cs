using Serilog;
using System.Collections.Concurrent;
using System.Data.SQLite;
using System.Diagnostics;

namespace MassBlattLinkerr.Core
{
    /// <summary>
    /// Class for handling database-related operations such as initialization, clearing, and querying.
    /// </summary>
    public class DatabaseHandler
    {
        private readonly string DatabaseFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "articles.db");
        private readonly ILogger? _logger;

        /// <summary>
        /// Initializes the ExcelHandler with the application's logger.
        /// </summary>
        /// <param name="logger">The logger instance from the main application.</param>
        public DatabaseHandler(ILogger? logger = null)
        {
            if (logger is not null)
            {
                _logger = logger;
            }
        }

        /// <summary>
        /// Initializes and populates the SQLite database with data from the file system.
        /// </summary>
        public void InitializeAndPopulateDatabase()
        {
            string directoryPath = @"Z:\Zdenek\Maßblatty-PDF";

            ClearDatabase();
            InitializeDatabase();

            var articlesData = new ConcurrentBag<(string ArticleName, string PrintVersionName, string FilePath)>();

            Parallel.ForEach(Directory.EnumerateFiles(directoryPath, "*.pdf", SearchOption.AllDirectories), file =>
            {
                var data = new ArticleHandler().ExtractDataFromPath(file);
                if (data != null)
                {
                    articlesData.Add(data.Value);
                }
            });

            InsertDataIntoDatabase(articlesData);
            _logger?.Information("Naplnění databáze bylo úspěšné.");
        }

        /// <summary>
        /// Initializes the SQLite database, creating the required table.
        /// </summary>
        private void InitializeDatabase()
        {
            // Generate a path for the temporary database in the system's temp folder
            string tempDatabaseFile = Path.Combine(Path.GetTempPath(), "temp_articles.db");

            if (!File.Exists(DatabaseFile))
            {
                try
                {
                    // Create the database in the temp folder first
                    using (var connection = new SQLiteConnection($"Data Source={tempDatabaseFile};Version=3;New=True;"))
                    {
                        connection.Open();
                        string createTableQuery = @"CREATE TABLE Articles (
                                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                            ArticleName TEXT NOT NULL,
                                            PrintVersionName TEXT NOT NULL,
                                            FilePath TEXT NOT NULL
                                        )";
                        using (var command = new SQLiteCommand(createTableQuery, connection))
                        {
                            command.ExecuteNonQuery();
                        }
                        connection.Close();
                    }

                    // Copy the temp database to the desired final location
                    File.Copy(tempDatabaseFile, DatabaseFile, true);
                    _logger?.Information($"Database successfully created and moved to: {DatabaseFile}");
                }
                catch (SQLiteException ex)
                {
                    _logger?.Error($"SQLiteException: Error during database initialization.\n" +
                                  $"Exception Message: '{ex.Message}'.\n" +
                                  $"Exception StackTrace: '{ex.StackTrace}'.");
                }
                catch (Exception ex)
                {
                    _logger?.Error($"Exception: Error during database initialization.\n" +
                                  $"Exception Message: '{ex.Message}'.\n" +
                                  $"Exception StackTrace: '{ex.StackTrace}'.");
                }
                finally
                {
                    if (File.Exists(tempDatabaseFile)) File.Delete(tempDatabaseFile);
                }
            }
        }

        /// <summary>
        /// Clears the database by deleting all data from the Articles table.
        /// </summary>
        private void ClearDatabase()
        {
            // Generate a path for the temporary database in the system's temp folder
            string tempDatabaseFile = Path.Combine(Path.GetTempPath(), "temp_articles.db");

            if (File.Exists(DatabaseFile))
            {
                File.Copy(DatabaseFile, tempDatabaseFile, true);

                using (var connection = new SQLiteConnection($"Data Source={tempDatabaseFile};Version=3;"))
                {
                    connection.Open();
                    string clearTableQuery = "DELETE FROM Articles";
                    using (var command = new SQLiteCommand(clearTableQuery, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                    connection.Close();
                }

                File.Copy(tempDatabaseFile, DatabaseFile, true);
                if (File.Exists(tempDatabaseFile)) File.Delete(tempDatabaseFile);
            }
        }

        /// <summary>
        /// Inserts the extracted data into the database.
        /// </summary>
        private void InsertDataIntoDatabase(ConcurrentBag<(string ArticleName, string PrintVersionName, string FilePath)> articlesData)
        {
            string tempDatabaseFile = Path.Combine(Path.GetTempPath(), "temp_articles.db");
            File.Copy(DatabaseFile, tempDatabaseFile, true);

            using (var connection = new SQLiteConnection($"Data Source={tempDatabaseFile};Version=3;"))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    foreach (var data in articlesData)
                    {
                        string insertQuery = "INSERT INTO Articles (ArticleName, PrintVersionName, FilePath) VALUES (@ArticleName, @PrintVersionName, @FilePath)";
                        using (var command = new SQLiteCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ArticleName", data.ArticleName);
                            command.Parameters.AddWithValue("@PrintVersionName", data.PrintVersionName);
                            command.Parameters.AddWithValue("@FilePath", data.FilePath);
                            command.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }
                connection.Close();
            }

            File.Copy(tempDatabaseFile, DatabaseFile, true);
            if (File.Exists(tempDatabaseFile)) File.Delete(tempDatabaseFile);
        }

        /// <summary>
        /// Searches the database for files by article name and version.
        /// If no results are found, it calls InitializeAndPopulateDatabase and tries again.
        /// If still no results are found, opens the folder where the files should be located.
        /// </summary>
        /// <param name="articleName">Article name to search for.</param>
        /// <param name="printVersionName">Print version name to search for.</param>
        /// <returns>List of file paths that match the search criteria.</returns>
        public List<string> SearchByArticleAndVersion(string articleName, string printVersionName)
        {
            // Validate input parameters
            if (string.IsNullOrWhiteSpace(articleName) || string.IsNullOrWhiteSpace(printVersionName))
            {
                throw new ArgumentException("Article name and print version name cannot be null or empty.");
            }

            var results = new List<string>();
            string tempDatabaseFile = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.db");

            try
            {
                // Check if the original database file exists
                if (!File.Exists(DatabaseFile))
                {
                    _logger?.Error($"Database file does not exist: {DatabaseFile}");
                    InitializeAndPopulateDatabase();
                }

                // First attempt to search the database
                results = SearchInDatabase(articleName, printVersionName, tempDatabaseFile);

                // If no results are found, reinitialize and repopulate the database and search again
                if (results.Count == 0)
                {
                    _logger?.Warning("No results found, initializing and repopulating the database.");
                    InitializeAndPopulateDatabase();

                    // Try searching the newly populated database
                    results = SearchInDatabase(articleName, printVersionName, tempDatabaseFile);
                }

                // If still no results, open the directory where the files should be located
                if (results.Count == 0)
                {
                    _logger?.Warning("No exact match found. Opening folder where the files should be located.");

                    string fallbackQuery = "SELECT FilePath FROM Articles WHERE ArticleName = @ArticleName";

                    using (var connection = new SQLiteConnection($"Data Source={tempDatabaseFile};Version=3;"))
                    {
                        connection.Open();
                        using (var command = new SQLiteCommand(fallbackQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ArticleName", articleName);

                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string filePath = reader["FilePath"] != DBNull.Value ? reader["FilePath"].ToString() : string.Empty;

                                    if (!string.IsNullOrEmpty(filePath))
                                    {
                                        string matchedPrintVersionPathDirectory = Path.GetDirectoryName(filePath);
                                        _logger?.Warning("Nepodařilo se otevřít soubor. Otevírám složku {matchedPrintVersionPathDirectory}", matchedPrintVersionPathDirectory);

                                        Process.Start(new ProcessStartInfo(matchedPrintVersionPathDirectory) { UseShellExecute = true });
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Error($"Exception: {ex.Message}\nStack Trace: {ex.StackTrace}");
            }
            finally
            {
                if (File.Exists(tempDatabaseFile))
                {
                    try
                    {
                        File.Delete(tempDatabaseFile);
                    }
                    catch (Exception ex)
                    {
                        _logger?.Warning($"Failed to delete temporary database file '{tempDatabaseFile}'. Exception: {ex.Message}");
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// Searches the temporary database for matching article and print version.
        /// </summary>
        /// <param name="articleName">Article name to search for.</param>
        /// <param name="printVersionName">Print version name to search for.</param>
        /// <param name="tempDatabaseFile">Path to the temporary database file.</param>
        /// <returns>List of matching file paths.</returns>
        private List<string> SearchInDatabase(string articleName, string printVersionName, string tempDatabaseFile)
        {
            var results = new List<string>();

            File.Copy(DatabaseFile, tempDatabaseFile, true);

            using (var connection = new SQLiteConnection($"Data Source={tempDatabaseFile};Version=3;"))
            {
                connection.Open();

                string searchQuery = "SELECT FilePath FROM Articles WHERE ArticleName = @ArticleName AND PrintVersionName = @PrintVersionName";
                using (var command = new SQLiteCommand(searchQuery, connection))
                {
                    command.Parameters.AddWithValue("@ArticleName", articleName);
                    command.Parameters.AddWithValue("@PrintVersionName", printVersionName);

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            results.Add(reader["FilePath"].ToString());
                        }
                    }
                }

                connection.Close();
            }

            return results;
        }
    }
}
